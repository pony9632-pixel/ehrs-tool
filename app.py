"""文中 eHRS 排班/打卡 桌面工具(CustomTkinter)。

啟動:
    python3 app.py

第一次用請到「設定」分頁填帳號密碼並儲存(存本機 ~/.ehrs_tool,不進 git)。
排班分頁:選年月→載入→雙擊任一格可改班/刪班(會二次確認後才寫入正式班表)。
打卡分頁:選日期區間→查詢。
"""
import calendar as _cal
import datetime as dt
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import load_config, save_config
from ehrs_client import DEFAULT_BASE_URL, EhrsClient

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

_WEEK = "一二三四五六日"


class EhrsApp(ctk.CTk):
    def __init__(self, auto_login: bool = True) -> None:
        super().__init__()
        self.title("文中排班工具")
        self.geometry("1180x700")

        self.client: EhrsClient | None = None
        self.account: str = ""
        self.cfg = load_config()
        self.schedule: list = []
        self.row_emp: dict = {}
        self.shift_codes: dict[str, tuple[str, int]] = {}

        self._build_ui()
        if auto_login:
            self.after(250, self._do_login)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        self._build_schedule_tab(self.tabview.add("排班"))
        self._build_punch_tab(self.tabview.add("打卡"))
        self._build_settings_tab(self.tabview.add("設定"))

        self.status = ctk.CTkLabel(self, text="就緒", anchor="w")
        self.status.pack(side="bottom", fill="x", padx=12, pady=(0, 8))

    def _build_schedule_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab)
        bar.pack(fill="x", padx=6, pady=6)
        now = dt.date.today()
        self.year_var = ctk.StringVar(value=str(now.year))
        self.month_var = ctk.StringVar(value=str(now.month))
        ctk.CTkLabel(bar, text="年").pack(side="left", padx=(8, 2))
        ctk.CTkEntry(bar, textvariable=self.year_var, width=70).pack(side="left")
        ctk.CTkLabel(bar, text="月").pack(side="left", padx=(10, 2))
        ctk.CTkEntry(bar, textvariable=self.month_var, width=50).pack(side="left")
        ctk.CTkButton(bar, text="載入排班", command=self._load_schedule).pack(
            side="left", padx=12
        )
        ctk.CTkLabel(bar, text="提示:雙擊格子可改班/刪班").pack(side="left", padx=8)

        wrap = tk.Frame(tab)
        wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.grid_tv = ttk.Treeview(wrap, show="headings", height=18)
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.grid_tv.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.grid_tv.xview)
        self.grid_tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.grid_tv.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.grid_tv.bind("<Double-1>", self._on_grid_double)

    def _build_punch_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab)
        bar.pack(fill="x", padx=6, pady=6)
        today = dt.date.today()
        self.p_start = ctk.CTkEntry(bar, width=110)
        self.p_start.insert(0, (today - dt.timedelta(days=7)).isoformat())
        self.p_end = ctk.CTkEntry(bar, width=110)
        self.p_end.insert(0, today.isoformat())
        self.p_emp = ctk.CTkEntry(bar, width=90, placeholder_text="員工代號")
        self.p_all = ctk.CTkCheckBox(bar, text="全員")
        ctk.CTkLabel(bar, text="起").pack(side="left", padx=(8, 2))
        self.p_start.pack(side="left")
        ctk.CTkLabel(bar, text="迄").pack(side="left", padx=(10, 2))
        self.p_end.pack(side="left")
        ctk.CTkLabel(bar, text="員工").pack(side="left", padx=(10, 2))
        self.p_emp.pack(side="left")
        self.p_all.pack(side="left", padx=10)
        ctk.CTkButton(bar, text="查詢打卡", command=self._query_punch).pack(
            side="left", padx=12
        )

        wrap = tk.Frame(tab)
        wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        cols = [
            ("emp", "員工代號", 90),
            ("name", "姓名", 90),
            ("date", "出勤日期", 110),
            ("time", "刷卡時間", 90),
            ("src", "來源", 80),
            ("clock", "卡鐘", 70),
        ]
        self.punch_tv = ttk.Treeview(
            wrap, show="headings", columns=[c[0] for c in cols], height=18
        )
        for key, title, width in cols:
            self.punch_tv.heading(key, text=title)
            self.punch_tv.column(key, width=width, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.punch_tv.yview)
        self.punch_tv.configure(yscrollcommand=ysb.set)
        self.punch_tv.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

    def _build_settings_tab(self, tab) -> None:
        frm = ctk.CTkFrame(tab)
        frm.pack(padx=20, pady=20, anchor="nw")
        ctk.CTkLabel(frm, text="帳號").grid(row=0, column=0, sticky="e", padx=8, pady=8)
        self.s_acc = ctk.CTkEntry(frm, width=240)
        self.s_acc.insert(0, self.cfg.get("account", ""))
        self.s_acc.grid(row=0, column=1, pady=8)
        ctk.CTkLabel(frm, text="密碼").grid(row=1, column=0, sticky="e", padx=8, pady=8)
        self.s_pwd = ctk.CTkEntry(frm, width=240, show="*")
        self.s_pwd.insert(0, self.cfg.get("password", ""))
        self.s_pwd.grid(row=1, column=1, pady=8)
        ctk.CTkLabel(frm, text="站台網址").grid(
            row=2, column=0, sticky="e", padx=8, pady=8
        )
        self.s_base = ctk.CTkEntry(frm, width=360)
        self.s_base.insert(0, self.cfg.get("base_url", DEFAULT_BASE_URL))
        self.s_base.grid(row=2, column=1, pady=8)
        ctk.CTkButton(frm, text="儲存並登入", command=self._save_and_login).grid(
            row=3, column=1, sticky="w", pady=12
        )
        ctk.CTkLabel(
            frm,
            text="帳密只存在本機 ~/.ehrs_tool/config.json(限本人讀取)",
            text_color="gray",
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=8)

    # -------------------------------------------------------------- helpers
    def _set_status(self, text: str) -> None:
        self.status.configure(text=text)

    def _run_async(self, work, on_done, on_error=None) -> None:
        def worker():
            try:
                result = work()
            except Exception as exc:  # noqa: BLE001 - 回報到 UI
                self.after(0, lambda e=exc: (on_error or self._error)(e))
                return
            self.after(0, lambda: on_done(result))

        threading.Thread(target=worker, daemon=True).start()

    def _error(self, exc: Exception) -> None:
        self._set_status(f"錯誤:{exc}")
        messagebox.showerror("錯誤", str(exc))

    def _need_client(self) -> bool:
        if self.client is None:
            messagebox.showwarning("尚未登入", "請先到「設定」分頁登入。")
            self.tabview.set("設定")
            return False
        return True

    # ---------------------------------------------------------------- login
    def _save_and_login(self) -> None:
        self.cfg = {
            "account": self.s_acc.get().strip(),
            "password": self.s_pwd.get(),
            "base_url": self.s_base.get().strip() or DEFAULT_BASE_URL,
        }
        save_config(self.cfg)
        self._do_login()

    def _do_login(self) -> None:
        acc = self.cfg.get("account")
        pwd = self.cfg.get("password")
        base = self.cfg.get("base_url", DEFAULT_BASE_URL)
        if not acc or not pwd:
            self._set_status("尚未設定帳密,請到「設定」分頁")
            self.tabview.set("設定")
            return
        self._set_status("登入中…")

        def work():
            return EhrsClient(base_url=base).login(acc, pwd)

        def done(client):
            self.client = client
            self.account = acc
            self._set_status(f"已登入:{acc}")

        self._run_async(work, done)

    # ------------------------------------------------------------- schedule
    def _load_schedule(self) -> None:
        if not self._need_client():
            return
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            messagebox.showwarning("輸入錯誤", "年/月請填數字。")
            return
        self._set_status(f"載入 {year}-{month:02d} 排班中…")

        def work():
            return self.client.get_schedule(year, month)

        def done(sched):
            self.schedule = sched
            self._build_code_map(sched)
            self._populate_grid(year, month, sched)
            self._set_status(f"已載入 {year}-{month:02d} 排班(共 {len(sched)} 人)")

        self._run_async(work, done)

    def _build_code_map(self, sched) -> None:
        codes: dict[str, tuple[str, int]] = {}
        for emp in sched:
            for s in emp.shifts:
                if s.code:
                    codes[s.code] = (s.name, s.kind if s.kind is not None else 3)
        self.shift_codes = codes

    def _populate_grid(self, year: int, month: int, sched) -> None:
        ndays = _cal.monthrange(year, month)[1]
        cols = ["emp"] + [f"d{d}" for d in range(1, ndays + 1)]
        tv = self.grid_tv
        tv.delete(*tv.get_children())
        tv["columns"] = cols
        tv.heading("emp", text="員工")
        tv.column("emp", width=150, anchor="w", stretch=False)
        for d in range(1, ndays + 1):
            wd = _WEEK[_cal.weekday(year, month, d)]
            tv.heading(f"d{d}", text=f"{d}/{wd}")
            tv.column(f"d{d}", width=52, anchor="center", stretch=False)
        self.row_emp = {}
        for emp in sched:
            by_day = {int(s.date[8:10]): s.code for s in emp.shifts}
            values = [f"{emp.emp_id} {emp.name}"] + [
                by_day.get(d, "") for d in range(1, ndays + 1)
            ]
            iid = tv.insert("", "end", values=values)
            self.row_emp[iid] = emp

    def _on_grid_double(self, event) -> None:
        tv = self.grid_tv
        if tv.identify("region", event.x, event.y) != "cell":
            return
        col = tv.identify_column(event.x)
        row = tv.identify_row(event.y)
        if not row or col in ("", "#1"):
            return
        day = int(col[1:]) - 1
        emp = self.row_emp.get(row)
        if emp is None or day < 1:
            return
        self._open_editor(emp, day)

    def _open_editor(self, emp, day: int) -> None:
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        date_str = f"{year:04d}-{month:02d}-{day:02d}"
        cur = emp.shift_on(date_str)

        top = ctk.CTkToplevel(self)
        top.title("編輯排班")
        top.geometry("380x300")
        top.transient(self)
        top.after(60, top.grab_set)

        ctk.CTkLabel(
            top, text=f"{emp.emp_id} {emp.name}", font=("", 16, "bold")
        ).pack(pady=(18, 2))
        ctk.CTkLabel(top, text=date_str).pack()
        cur_txt = f"目前:{cur.code} {cur.name}" if cur else "目前:(空白)"
        ctk.CTkLabel(top, text=cur_txt, text_color="gray").pack(pady=6)

        options = ["(空白/刪除)"] + [
            f"{c} {n}" for c, (n, _k) in sorted(self.shift_codes.items())
        ]
        default = (
            f"{cur.code} {self.shift_codes.get(cur.code, (cur.name, 3))[0]}"
            if cur
            else options[0]
        )
        var = ctk.StringVar(value=default if default in options else options[0])
        ctk.CTkOptionMenu(top, values=options, variable=var, width=300).pack(pady=10)

        def save():
            choice = var.get()
            top.destroy()
            if choice.startswith("(空白"):
                self._apply_delete(emp, date_str, cur)
            else:
                self._apply_set(emp, date_str, choice.split()[0])

        ctk.CTkButton(top, text="儲存", command=save).pack(pady=(6, 4))
        ctk.CTkButton(top, text="取消", command=top.destroy, fg_color="gray").pack()

    def _apply_set(self, emp, date_str: str, code: str) -> None:
        name, kind = self.shift_codes.get(code, ("", 3))
        if not messagebox.askyesno(
            "確認寫入",
            f"將把 {emp.emp_id} {emp.name}\n{date_str} 設為 {code} {name}\n\n"
            "這會寫入正式班表,確定?",
        ):
            return
        year, month, _ = (int(x) for x in date_str.split("-"))
        self._set_status("寫入中…")

        def work():
            return self.client.set_shift(
                year, month, emp.emp_id, date_str, code, kind=kind, dry_run=False
            )

        def done(res):
            self._set_status(f"已寫入 {emp.emp_id} {date_str} → {code}")
            self._load_schedule()

        self._run_async(work, done)

    def _apply_delete(self, emp, date_str: str, cur) -> None:
        if cur is None:
            return
        if not messagebox.askyesno(
            "確認刪除",
            f"將刪除 {emp.emp_id} {emp.name}\n{date_str} 的班 {cur.code}\n\n確定?",
        ):
            return
        year, month, _ = (int(x) for x in date_str.split("-"))
        self._set_status("刪除中…")

        def work():
            return self.client.delete_shift(
                year, month, emp.emp_id, date_str, dry_run=False
            )

        def done(res):
            self._set_status(f"已刪除 {emp.emp_id} {date_str}")
            self._load_schedule()

        self._run_async(work, done)

    # ---------------------------------------------------------------- punch
    def _query_punch(self) -> None:
        if not self._need_client():
            return
        start = self.p_start.get().strip()
        end = self.p_end.get().strip()
        emp = None if self.p_all.get() else (self.p_emp.get().strip() or self.account)
        self._set_status("查詢打卡中…")

        def work():
            return self.client.get_punch_records(
                start, end, emp_start=emp, emp_end=emp, readable=True
            )

        def done(rows):
            self._populate_punch(rows)
            scope = "全員" if self.p_all.get() else (emp or "")
            self._set_status(f"打卡 {start}~{end}({scope}) 共 {len(rows)} 筆")

        self._run_async(work, done)

    def _populate_punch(self, rows) -> None:
        tv = self.punch_tv
        tv.delete(*tv.get_children())
        for r in rows:
            tv.insert(
                "",
                "end",
                values=(
                    r.get("員工代號", ""),
                    r.get("員工姓名", ""),
                    str(r.get("出勤日期", ""))[:10],
                    str(r.get("刷卡時間", ""))[11:16],
                    r.get("來源", ""),
                    r.get("卡鐘代號", ""),
                ),
            )


def main() -> None:
    EhrsApp().mainloop()


if __name__ == "__main__":
    main()
