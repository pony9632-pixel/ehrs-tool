"""文中 eHRS 排班/打卡 桌面工具(CustomTkinter)。

啟動:
    python3 app.py

第一次用請到「設定」分頁填帳號密碼並儲存(存本機 ~/.ehrs_tool,不進 git)。
排班分頁:選年月→載入→雙擊任一格可改班/刪班(會二次確認後才寫入正式班表)。
打卡分頁:選日期區間→查詢。
"""
import calendar as _cal
import datetime as dt
import re
import threading
import tkinter as tk
from itertools import groupby
from tkinter import messagebox, ttk

import customtkinter as ctk

import tkinter.filedialog as filedialog

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from config import load_config, save_config
from ehrs_client import DEFAULT_BASE_URL, EhrsClient
from shift_codes_ref import SHIFT_CODES_REF
from version import __version__

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ─── UI 色彩字典（參考 HTML 設計稿配色）────────────────────────────────────────
_C = {
    "accent":  "#4A6FE3",   # 主色：靛藍
    "acc_h":   "#3B5DD0",   # 主色 hover
    "ink":     "#2D3048",   # 主文字：深藍灰
    "muted":   "#7B8097",   # 次文字：中灰
    "app_bg":  "#F5F7FB",   # 應用底色：極淡藍灰
    "tab_bg":  "#ECEEF5",   # 分頁列容器
    "line":    "#E8EBF4",   # 邊框 / 分隔線
    "white":   "#FFFFFF",
    "tbl_hdr": "#EEF2FB",   # 表格表頭
    "row_alt": "#F8F9FD",   # 隔行底色
    "red":     "#CC2200",
    "green":   "#2A8B5A",
}

_WEEK = "一二三四五六日"


def _get_schedule_periods(roc_year: int) -> list[tuple[dt.date, dt.date]]:
    """計算指定民國年度的 14 個四週排班期間（每期 28 天）。
    民國 115 年基準起始日：2025-12-22（週一）。
    相鄰年度以 ±52 週（364 天）近似推算，維持週一起始。
    回傳 [(start, end), ...] × 14。
    """
    _ANCHOR_ROC  = 115
    _ANCHOR_DATE = dt.date(2025, 12, 22)
    diff   = roc_year - _ANCHOR_ROC
    anchor = _ANCHOR_DATE + dt.timedelta(weeks=52 * diff)
    return [
        (anchor + dt.timedelta(days=28 * i),
         anchor + dt.timedelta(days=28 * i + 27))
        for i in range(14)
    ]


def _is_rest(code: str) -> bool:
    """排休班別判斷：991 及 991-x 開頭的都算排休。"""
    return bool(code) and code.startswith("991")


def _shift_times(code: str) -> tuple[str, str]:
    """從 SHIFT_CODES_REF 解析班別的表定上班/下班時間（HH:MM）。"""
    name = SHIFT_CODES_REF.get(code, ("", 3, 0.0))[0]
    m = re.search(r"(\d{2}:\d{2})-(\d{2}:\d{2})", name)
    return (m.group(1), m.group(2)) if m else ("", "")


def _shift_hours(code: str) -> float:
    """從 SHIFT_CODES_REF 直接取得班別時數（已含休息扣除）。"""
    return SHIFT_CODES_REF.get(code, ("", 3, 0.0))[2]


def _punch_mins(t: str) -> int:
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def _split_punches(punches: list[str], sched_in: str, sched_out: str) -> tuple[str, str]:
    """多筆打卡依距表定時間分為上班群/下班群，取上班群最早、下班群最晚。
    無排班資料時直接取第一筆和最後一筆。"""
    if not sched_in or not sched_out:
        return punches[0], punches[-1]
    si = _punch_mins(sched_in)
    so = _punch_mins(sched_out)
    if so < si:
        so += 1440
    in_grp, out_grp = [], []
    for p in punches:
        pm = _punch_mins(p)
        if pm < si - 120:
            pm += 1440
        (out_grp if abs(pm - so) <= abs(pm - si) else in_grp).append(p)
    return (in_grp[0] if in_grp else ""), (out_grp[-1] if out_grp else "")


def _split_punches_jiezhuan(punches: list[str], sched_in: str, sched_out: str) -> tuple[str, str]:
    """結轉打卡專用分群：不做 2 小時截斷，直接以距表定時間近遠判斷上/下班。
    支援跨夜班（凌晨打卡加 1440 修正）。"""
    if not sched_in or not sched_out:
        return punches[0], punches[-1]
    si = _punch_mins(sched_in)
    so = _punch_mins(sched_out)
    cross = so < si
    if cross:
        so += 1440   # 跨夜班：把下班時間加一天
    in_grp, out_grp = [], []
    for p in punches:
        pm = _punch_mins(p)
        # 跨夜班：凌晨打卡若距上班超過 6 小時，視為隔天
        if cross and pm < si - 360:
            pm += 1440
        (out_grp if abs(pm - so) <= abs(pm - si) else in_grp).append(p)
    return (in_grp[0] if in_grp else ""), (out_grp[-1] if out_grp else "")


def _punch_remark(clock_in: str, clock_out: str, sched_in: str, sched_out: str) -> str:
    """產生備注文字：遲到、早退、上班忘打卡、下班忘打卡（可複合）。"""
    remarks: list[str] = []
    if not clock_in:
        remarks.append("上班忘打卡")
    elif sched_in and _punch_mins(clock_in) > _punch_mins(sched_in):
        remarks.append("遲到")
    if not clock_out:
        remarks.append("下班忘打卡")
    elif sched_out and sched_in:
        si = _punch_mins(sched_in)
        so = _punch_mins(sched_out)
        co = _punch_mins(clock_out)
        if so < si:
            so += 1440
        if co < si - 120:
            co += 1440
        if co < so:
            remarks.append("早退")
    return " ".join(remarks)


def _calc_hours(sched_in: str, sched_out: str) -> float:
    """計算實際工時（扣除休息）。
    ≤ 4h：不扣；> 4h 且 < 9h：扣 0.5h；≥ 9h：扣 1h。"""
    si = _punch_mins(sched_in)
    so = _punch_mins(sched_out)
    if so < si:
        so += 1440
    total = (so - si) / 60.0
    if total <= 4.0:
        return total
    return total - (1.0 if total >= 9.0 else 0.5)


def _suggest_shifts(
    clock_in: str, clock_out: str, shift_codes: dict,
    preferred_hours: float = 8.0, top_n: int = 3
) -> list[tuple]:
    """找最適班別：實際打卡時間不會造成遲到/早退的班別 top_n 名。
    排序規則：工時差距最小優先（同工時 → 時間最接近）。
    回傳 [(code, sched_in, sched_out, delta_in, delta_out, hours), ...]
    delta_in > 0 = 比排班提早到；delta_out > 0 = 比排班晚下班（加班）。
    """
    ci_m = _punch_mins(clock_in) if clock_in else None
    co_m = _punch_mins(clock_out) if clock_out else None
    if ci_m is None and co_m is None:
        return []
    candidates = []
    for code in shift_codes:
        if _is_rest(code):
            continue
        si, so = _shift_times(code)
        if not si or not so:
            continue
        si_m = _punch_mins(si)
        so_m = _punch_mins(so)
        if so_m < si_m:
            so_m += 1440
        ci_a = (ci_m + 1440) if ci_m is not None and ci_m < si_m - 240 else ci_m
        co_a = (co_m + 1440) if co_m is not None and co_m < si_m - 240 else co_m
        d_in  = (si_m - ci_a) if ci_a is not None else 0   # + = 早到
        d_out = (co_a - so_m) if co_a is not None else 0   # + = 加班
        if ci_a is not None and (d_in < 0 or d_in > 120):
            continue
        if co_a is not None and d_out < -5:
            continue
        h = _shift_hours(code) or _calc_hours(si, so)
        # 主要：工時差距（相差 0.5h = +1000 分）；次要：時間契合度
        score = abs(h - preferred_hours) * 1000 + d_in * 0.7 + max(0, d_out) * 0.5
        candidates.append((code, si, so, d_in, d_out, h, score))
    candidates.sort(key=lambda x: x[6])
    return [(c, si, so, d_in, d_out, h)
            for c, si, so, d_in, d_out, h, _ in candidates[:top_n]]


def _assign_single_punch(punch: str, sched_in: str, sched_out: str) -> tuple[str, str]:
    """只有一筆打卡時，依距表定上班/下班的近遠判斷是實際上班還是實際下班。
    回傳 (實際上班, 實際下班)，其中一個為空字串。"""
    if not sched_in and not sched_out:
        return punch, ""
    p = _punch_mins(punch)
    if sched_in and sched_out:
        si = _punch_mins(sched_in)
        so = _punch_mins(sched_out)
        if so < si:
            so += 1440
        if p < si - 120:
            p += 1440
        return ("", punch) if abs(p - so) <= abs(p - si) else (punch, "")
    return ("", punch) if sched_out else (punch, "")


class _PunchTable(tk.Frame):
    """Canvas-based table 支援單格著色。"""

    COLS = [
        ("員工代號",  90), ("員工姓名",  90), ("出勤日期", 110),
        ("排班代號",  70), ("表定上班",  80), ("表定下班",  80),
        ("實際上班",  80), ("實際下班",  80), ("備注",     140),
    ]
    ROW_H  = 27
    HDR_H  = 32
    FONT   = ("TkDefaultFont", 12)
    BFONT  = ("TkDefaultFont", 12, "bold")
    HDR_BG = "#EEF2FB"
    BG     = ("#FFFFFF", "#F8F9FD")
    GRID   = "#E8EBF4"

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._data: list[tuple] = []
        self._cv = tk.Canvas(self, bg="#FFFFFF", highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical",   command=self._cv.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self._cv.xview)
        self._cv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._cv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self._cv.bind("<MouseWheel>",
                      lambda e: self._cv.yview_scroll(int(-1 * e.delta / 60), "units"))
        self._tw = sum(w for _, w in self.COLS)
        self._draw_header()

    def _draw_header(self):
        x = 0
        for name, w in self.COLS:
            self._cv.create_rectangle(x, 0, x+w, self.HDR_H,
                                      fill=self.HDR_BG, outline=self.GRID)
            self._cv.create_text(x + w//2, self.HDR_H//2,
                                 text=name, font=self.BFONT, fill="#2D3048",
                                 anchor="center")
            x += w

    def populate(self, rows: list[tuple]) -> None:
        """rows: list of (values_tuple, fgs_tuple) — fgs 為各格顏色，空字串=黑。"""
        self._cv.delete("row")
        self._data = list(rows)
        y = self.HDR_H
        for i, (vals, fgs) in enumerate(rows):
            bg = self.BG[i % 2]
            x = 0
            for j, (val, (_, w)) in enumerate(zip(vals, self.COLS)):
                fg = fgs[j] if j < len(fgs) and fgs[j] else "#000000"
                self._cv.create_rectangle(x, y, x+w, y+self.ROW_H,
                                          fill=bg, outline=self.GRID, tags="row")
                self._cv.create_text(x+6, y+self.ROW_H//2,
                                     text="" if val is None else str(val),
                                     font=self.FONT, fill=fg,
                                     anchor="w", width=w-10, tags="row")
                x += w
            y += self.ROW_H
        self._cv.configure(scrollregion=(0, 0, self._tw, y))

    def all_values(self) -> list[tuple]:
        return [v for v, _ in self._data]


class EhrsApp(ctk.CTk):
    def __init__(self, auto_login: bool = True) -> None:
        super().__init__()
        self.title(f"文中排班工具  v{__version__}")
        self.geometry("1180x720")
        self.configure(fg_color=_C["app_bg"])

        self.client: EhrsClient | None = None
        self.account: str = ""
        self.cfg = load_config()
        self.schedule: list = []
        self.row_emp: dict = {}
        self.shift_codes: dict[str, tuple[str, int]] = dict(SHIFT_CODES_REF)
        self._punch_cache: list[tuple] = []
        self._raw_punch_rows: list[dict] = []   # 原始刷卡紀錄（每刷一筆一列）

        self._build_ui()
        if auto_login:
            self.after(250, self._do_login)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        # ── 設定 ttk 樣式（排班格 + 打卡紀錄用的 Treeview）──────────
        style = ttk.Style()
        style.configure("App.Treeview.Heading",
                         background=_C["tbl_hdr"], foreground=_C["ink"],
                         font=("TkDefaultFont", 12, "bold"), relief="flat")
        style.configure("App.Treeview",
                         rowheight=28, background=_C["white"],
                         foreground=_C["ink"], fieldbackground=_C["white"],
                         borderwidth=0)
        style.map("App.Treeview",
                  background=[("selected", _C["accent"])],
                  foreground=[("selected", "#FFFFFF")])
        # 去掉 Heading 的 relief border
        style.layout("App.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.label", {"sticky": "nswe"})
                ]})
            ]}),
        ])

        # ── 頂部品牌列 ─────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color=_C["white"], corner_radius=0,
                            height=54, border_width=0)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        # 底線
        tk.Frame(hdr, bg=_C["line"], height=1).pack(side="bottom", fill="x")

        # Logo 方塊
        logo = ctk.CTkFrame(hdr, fg_color=_C["accent"], width=34, height=34,
                              corner_radius=9)
        logo.pack(side="left", padx=(16, 10), pady=10)
        logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="文", font=("", 18, "bold"),
                      text_color="#FFFFFF").pack(expand=True)

        ctk.CTkLabel(hdr, text="文中排班",
                      font=("", 16, "bold"), text_color=_C["ink"]).pack(side="left")
        ctk.CTkLabel(hdr, text=f"v{__version__}",
                      font=("", 11), text_color=_C["muted"]).pack(
            side="left", padx=(5, 0), pady=(7, 0))

        # ── 分頁 ───────────────────────────────────────────────────
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=_C["white"],
            segmented_button_fg_color=_C["tab_bg"],
            segmented_button_selected_color=_C["white"],
            segmented_button_selected_hover_color=_C["white"],
            segmented_button_unselected_color=_C["tab_bg"],
            segmented_button_unselected_hover_color="#DDDFE8",
            text_color=_C["accent"],
            text_color_disabled=_C["muted"],
            border_color=_C["line"],
            border_width=1,
        )
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(8, 4))
        self._build_schedule_tab(self.tabview.add("排班"))
        self._build_punch_tab(self.tabview.add("打卡"))
        self._build_raw_punch_tab(self.tabview.add("打卡紀錄"))
        self._build_suggest_tab(self.tabview.add("建議"))
        self._build_settings_tab(self.tabview.add("設定"))

        # ── 底部 status bar（含匯入進度條）─────────────────────────
        btm = ctk.CTkFrame(self, fg_color=_C["white"], height=36,
                            corner_radius=0, border_width=0)
        btm.pack(side="bottom", fill="x")
        btm.pack_propagate(False)
        # 頂線
        tk.Frame(btm, bg=_C["line"], height=1).pack(side="top", fill="x")

        # 狀態指示燈
        self._status_dot = ctk.CTkLabel(btm, text="●", font=("", 10),
                                          text_color=_C["muted"], width=16)
        self._status_dot.pack(side="left", padx=(14, 0))

        self.status = ctk.CTkLabel(btm, text="就緒", anchor="w",
                                     text_color=_C["muted"], font=("", 12))
        self.status.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # 進度條區塊（平時隱藏，匯入時顯示）
        self._prog_frm = ctk.CTkFrame(btm, fg_color="transparent")
        self._prog_bar = ctk.CTkProgressBar(self._prog_frm, width=180, height=12,
                                              progress_color=_C["accent"],
                                              fg_color=_C["line"])
        self._prog_bar.set(0)
        self._prog_bar.pack(side="left", padx=(0, 6))
        self._prog_lbl = ctk.CTkLabel(self._prog_frm, text="", width=70,
                                        anchor="e", font=("", 12),
                                        text_color=_C["muted"])
        self._prog_lbl.pack(side="left")

    def _build_schedule_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab, fg_color=_C["white"], corner_radius=8,
                            border_width=1, border_color=_C["line"])
        bar.pack(fill="x", padx=6, pady=6)

        now = dt.date.today()
        self.year_var  = ctk.StringVar(value=str(now.year))
        self.month_var = ctk.StringVar(value=str(now.month))
        self._period_start: dt.date | None = None
        self._period_end:   dt.date | None = None
        self._grid_dates:   list[dt.date]  = []

        # ── 民國年 ──────────────────────────────────────────
        ctk.CTkLabel(bar, text="民國", text_color=_C["ink"],
                      font=("", 13)).pack(side="left", padx=(12, 2), pady=6)
        self.roc_year_var = ctk.StringVar(value=str(now.year - 1911))
        roc_ent = ctk.CTkEntry(bar, textvariable=self.roc_year_var, width=54,
                                border_color=_C["line"], fg_color=_C["white"],
                                text_color=_C["ink"])
        roc_ent.pack(side="left", pady=6)
        ctk.CTkLabel(bar, text="年", text_color=_C["ink"],
                      font=("", 13)).pack(side="left", padx=(2, 8))

        # ── 排班期間下拉 ─────────────────────────────────────
        self.period_var  = ctk.StringVar()
        self.period_menu = ctk.CTkOptionMenu(
            bar, variable=self.period_var, width=200,
            fg_color=_C["white"], button_color=_C["accent"],
            button_hover_color=_C["acc_h"],
            dropdown_fg_color=_C["white"],
            dropdown_hover_color=_C["tab_bg"],
            text_color=_C["ink"], dropdown_text_color=_C["ink"],
            command=self._on_period_select,
        )
        self.period_menu.pack(side="left", pady=6)
        roc_ent.bind("<Return>",   lambda e: self._refresh_period_menu())
        roc_ent.bind("<FocusOut>", lambda e: self._refresh_period_menu())

        # ── 操作按鈕 ─────────────────────────────────────────
        ctk.CTkButton(
            bar, text="載入排班",
            fg_color=_C["accent"], hover_color=_C["acc_h"],
            corner_radius=8, font=("", 13, "bold"),
            command=self._load_schedule,
        ).pack(side="left", padx=10, pady=6)
        ctk.CTkLabel(bar, text="雙擊格子可改班 / 刪班",
                      text_color=_C["muted"], font=("", 12)).pack(side="left", padx=4)
        ctk.CTkButton(bar, text="匯出 Excel", width=100,
                       fg_color=_C["muted"], hover_color="#6B7280",
                       corner_radius=8, command=self._export_schedule
                       ).pack(side="right", padx=6, pady=6)
        ctk.CTkButton(bar, text="匯入 Excel", width=100,
                       fg_color=_C["muted"], hover_color="#6B7280",
                       corner_radius=8, command=self._import_schedule
                       ).pack(side="right", padx=(4, 0), pady=6)

        # ── 排班格 ───────────────────────────────────────────
        wrap = tk.Frame(tab, bg=_C["app_bg"])
        wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.grid_tv = ttk.Treeview(wrap, style="App.Treeview",
                                     show="headings", height=18)
        ysb = ttk.Scrollbar(wrap, orient="vertical",   command=self.grid_tv.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.grid_tv.xview)
        self.grid_tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.grid_tv.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.grid_tv.bind("<Double-1>", self._on_grid_double)

        # 初始化期間選單
        self._refresh_period_menu()

    # ── 排班期間選單輔助 ───────────────────────────────────────────────────────
    def _refresh_period_menu(self) -> None:
        """依目前民國年重建排班期間選單，並自動定位到當前期間。"""
        try:
            roc_year = int(self.roc_year_var.get().strip())
        except ValueError:
            return
        self.year_var.set(str(roc_year + 1911))
        periods = _get_schedule_periods(roc_year)
        labels  = self._period_labels(periods)
        self.period_menu.configure(values=labels)
        # 自動選中包含今天的期間（否則選最後一期）
        today   = dt.date.today()
        sel_idx = len(periods) - 1
        for i, (s, e) in enumerate(periods):
            if s <= today <= e:
                sel_idx = i
                break
        self.period_var.set(labels[sel_idx])
        self._on_period_select(labels[sel_idx])

    def _on_period_select(self, choice: str) -> None:
        """期間下拉變更：更新內部 _period_start / _period_end。"""
        try:
            roc_year = int(self.roc_year_var.get().strip())
        except ValueError:
            return
        periods = _get_schedule_periods(roc_year)
        labels  = self._period_labels(periods)
        try:
            idx = labels.index(choice)
        except ValueError:
            return
        s, e = periods[idx]
        self._period_start = s
        self._period_end   = e
        self.year_var.set(str(s.year))
        self.month_var.set(str(s.month))

    @staticmethod
    def _period_labels(periods: list) -> list[str]:
        return [
            f"第 {i+1:>2} 期　{s.month}/{s.day} ～ {e.month}/{e.day}"
            for i, (s, e) in enumerate(periods)
        ]

    def _build_punch_tab(self, tab) -> None:
        _btn_outline = dict(fg_color=_C["white"], text_color=_C["ink"],
                             hover_color=_C["tab_bg"], border_width=1,
                             border_color=_C["line"], corner_radius=8)
        bar = ctk.CTkFrame(tab, fg_color=_C["white"], corner_radius=8,
                            border_width=1, border_color=_C["line"])
        bar.pack(fill="x", padx=6, pady=6)
        today = dt.date.today()
        self.p_start = ctk.CTkEntry(bar, width=110, border_color=_C["line"],
                                      fg_color=_C["white"], text_color=_C["ink"])
        self.p_start.insert(0, today.replace(day=1).isoformat())
        self.p_end = ctk.CTkEntry(bar, width=110, border_color=_C["line"],
                                    fg_color=_C["white"], text_color=_C["ink"])
        self.p_end.insert(0, today.isoformat())
        self._selected_emps: set[str] = set()    # empty = 全員
        self._selected_depts: set[str] = set()   # empty = 全部門
        self._emp_dept: dict[str, str] = {}      # emp_id -> 部門名稱
        self.p_dept_btn = ctk.CTkButton(
            bar, text="部門：全部 ▼", width=130, **_btn_outline,
            command=self._open_dept_picker
        )
        self.p_emp_btn = ctk.CTkButton(
            bar, text="員工：全員 ▼", width=130, **_btn_outline,
            command=self._open_emp_picker
        )
        self.p_abnormal = ctk.CTkCheckBox(bar, text="只顯示異常",
                                            text_color=_C["ink"],
                                            command=self._apply_punch_filter)
        self.p_abnormal.select()
        ctk.CTkLabel(bar, text="起", text_color=_C["ink"],
                      font=("", 13)).pack(side="left", padx=(12, 2))
        self.p_start.pack(side="left", pady=6)
        ctk.CTkLabel(bar, text="迄", text_color=_C["ink"],
                      font=("", 13)).pack(side="left", padx=(10, 2))
        self.p_end.pack(side="left", pady=6)
        self.p_dept_btn.pack(side="left", padx=(10, 0), pady=6)
        self.p_emp_btn.pack(side="left", padx=(6, 0), pady=6)
        self.p_abnormal.pack(side="left", padx=10)
        ctk.CTkButton(bar, text="查詢打卡",
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=8, font=("", 13, "bold"),
                       command=self._query_punch).pack(side="left", padx=12, pady=6)
        ctk.CTkButton(bar, text="匯出 Excel", width=100,
                       fg_color=_C["muted"], hover_color="#6B7280",
                       corner_radius=8, command=self._export_punch
                       ).pack(side="right", padx=6, pady=6)

        wrap = tk.Frame(tab, bg=_C["app_bg"])
        wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.punch_tbl = _PunchTable(wrap)
        self.punch_tbl.grid(row=0, column=0, sticky="nsew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

    def _build_raw_punch_tab(self, tab) -> None:
        _btn_outline = dict(fg_color=_C["white"], text_color=_C["ink"],
                             hover_color=_C["tab_bg"], border_width=1,
                             border_color=_C["line"], corner_radius=8)
        bar = ctk.CTkFrame(tab, fg_color=_C["white"], corner_radius=8,
                            border_width=1, border_color=_C["line"])
        bar.pack(fill="x", padx=6, pady=6)

        self._raw_selected_depts: set[str] = set()
        self._raw_selected_emps:  set[str] = set()

        self.raw_dept_btn = ctk.CTkButton(
            bar, text="部門：全部 ▼", width=130, **_btn_outline,
            command=self._open_raw_dept_picker
        )
        self.raw_emp_btn = ctk.CTkButton(
            bar, text="員工：全員 ▼", width=130, **_btn_outline,
            command=self._open_raw_emp_picker
        )
        self.raw_count_lbl = ctk.CTkLabel(bar, text="共 0 筆",
                                            text_color=_C["muted"], font=("", 12))

        self.raw_dept_btn.pack(side="left", padx=(12, 0), pady=6)
        self.raw_emp_btn.pack(side="left", padx=(6, 0), pady=6)
        self.raw_count_lbl.pack(side="left", padx=10)
        ctk.CTkButton(bar, text="匯出 Excel", width=100,
                       fg_color=_C["muted"], hover_color="#6B7280",
                       corner_radius=8, command=self._export_raw_punch
                       ).pack(side="right", padx=6, pady=6)

        wrap = tk.Frame(tab, bg=_C["app_bg"])
        wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        cols = ("dept", "emp_id", "emp_name", "date", "time", "jiezhuan", "source")
        self.raw_tv = ttk.Treeview(wrap, style="App.Treeview",
                                    columns=cols, show="headings", height=20)
        for col, lbl, w in [
            ("dept",      "部門",   90),
            ("emp_id",    "員工代號", 80),
            ("emp_name",  "員工姓名", 80),
            ("date",      "出勤日期", 110),
            ("time",      "刷卡時間", 90),
            ("jiezhuan",  "結轉",    60),
            ("source",    "來源",    70),
        ]:
            self.raw_tv.heading(col, text=lbl)
            self.raw_tv.column(col, width=w, anchor="center", stretch=False)
        ysb = ttk.Scrollbar(wrap, orient="vertical",   command=self.raw_tv.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.raw_tv.xview)
        self.raw_tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.raw_tv.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

    def _populate_raw_punch(self) -> None:
        rows = self._raw_punch_rows
        # 套用篩選
        def _match(r: dict) -> bool:
            eid  = r.get("員工代號", "")
            dept = r.get("部門名稱") or r.get("pa51014FullName", "")
            if self._raw_selected_depts and dept not in self._raw_selected_depts:
                return False
            if self._raw_selected_emps and eid not in self._raw_selected_emps:
                return False
            return True

        filtered = [r for r in rows if _match(r)]
        self.raw_tv.delete(*self.raw_tv.get_children())
        for r in filtered:
            dept   = r.get("部門名稱") or r.get("pa51014FullName", "")
            eid    = r.get("員工代號", "")
            ename  = r.get("員工姓名", "")
            date   = str(r.get("出勤日期", ""))[:10]
            time_  = str(r.get("刷卡時間", ""))[11:16]
            jz     = r.get("結轉註記", "")
            src    = r.get("來源", "")
            self.raw_tv.insert("", "end", values=(dept, eid, ename, date, time_, jz, src))
        self.raw_count_lbl.configure(text=f"共 {len(filtered)} 筆")

    # ── 打卡紀錄篩選器 ────────────────────────────────────────
    def _open_raw_dept_picker(self) -> None:
        depts = sorted({
            r.get("部門名稱") or r.get("pa51014FullName", "")
            for r in self._raw_punch_rows
        } - {""})
        if not depts:
            messagebox.showinfo("提示", "請先在打卡分頁查詢資料。")
            return
        self._generic_picker(
            title="選擇部門", items=depts,
            selected=self._raw_selected_depts,
            on_confirm=lambda s: (
                setattr(self, "_raw_selected_depts", s),
                self.raw_dept_btn.configure(
                    text="部門：全部 ▼" if not s else f"部門：已選 {len(s)} ▼"
                ),
                self._populate_raw_punch(),
            )
        )

    def _open_raw_emp_picker(self) -> None:
        emps: dict[str, str] = {}
        for r in self._raw_punch_rows:
            eid = r.get("員工代號", "")
            name = r.get("員工姓名", "")
            # 只顯示符合部門篩選的員工
            dept = r.get("部門名稱") or r.get("pa51014FullName", "")
            if self._raw_selected_depts and dept not in self._raw_selected_depts:
                continue
            if eid:
                emps[eid] = name
        if not emps:
            messagebox.showinfo("提示", "請先在打卡分頁查詢資料。")
            return
        self._generic_picker(
            title="選擇員工",
            items=[f"{eid}  {name}" for eid, name in sorted(emps.items())],
            selected={f"{eid}  {emps[eid]}" for eid in self._raw_selected_emps if eid in emps},
            on_confirm=lambda s: (
                setattr(self, "_raw_selected_emps",
                        {item.split()[0] for item in s}),
                self.raw_emp_btn.configure(
                    text="員工：全員 ▼" if not s else f"員工：已選 {len(s)} 人 ▼"
                ),
                self._populate_raw_punch(),
            ),
            searchable=True,
        )

    def _generic_picker(
        self, title: str, items: list[str], selected: set[str],
        on_confirm, searchable: bool = False
    ) -> None:
        """通用多選清單彈窗。"""
        top = ctk.CTkToplevel(self)
        top.title(title)
        top.geometry("280x460")
        top.configure(fg_color=_C["app_bg"])
        top.resizable(False, True)
        top.transient(self)
        top.after(60, top.grab_set)

        search_var = ctk.StringVar()
        if searchable:
            ctk.CTkEntry(top, placeholder_text="搜尋…",
                          textvariable=search_var,
                          border_color=_C["line"], fg_color=_C["white"],
                          text_color=_C["ink"]).pack(
                fill="x", padx=12, pady=(12, 4))

        scroll = ctk.CTkScrollableFrame(top, fg_color=_C["white"],
                                          corner_radius=8,
                                          border_width=1, border_color=_C["line"])
        scroll.pack(fill="both", expand=True, padx=12,
                     pady=(8 if not searchable else 4, 0))

        chk_vars: dict[str, tk.BooleanVar] = {}
        chk_widgets: dict[str, ctk.CTkCheckBox] = {}
        for item in items:
            init = (not selected) or (item in selected)
            var = tk.BooleanVar(value=init)
            chk = ctk.CTkCheckBox(scroll, text=item, variable=var,
                                   text_color=_C["ink"])
            chk.pack(anchor="w", padx=4, pady=3)
            chk_vars[item] = var
            chk_widgets[item] = chk

        if searchable:
            def _filter(*_):
                q = search_var.get().lower()
                for it, w in chk_widgets.items():
                    if not q or q in it.lower():
                        w.pack(anchor="w", padx=4, pady=3)
                    else:
                        w.pack_forget()
            search_var.trace_add("write", _filter)

        btn_bar = ctk.CTkFrame(top, fg_color="transparent")
        btn_bar.pack(fill="x", padx=12, pady=(6, 0))
        ctk.CTkButton(btn_bar, text="全選", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(True)  for v in chk_vars.values()]
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_bar, text="清除", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(False) for v in chk_vars.values()]
                       ).pack(side="left")

        def apply():
            picked = {it for it, v in chk_vars.items() if v.get()}
            final  = set() if len(picked) == len(items) else picked
            on_confirm(final)
            top.destroy()

        ctk.CTkButton(top, text="確定",
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=9, font=("", 13, "bold"),
                       command=apply).pack(fill="x", padx=12, pady=10)

    def _export_raw_punch(self) -> None:
        rows = list(self.raw_tv.get_children())
        if not rows:
            messagebox.showwarning("無資料", "沒有打卡紀錄可匯出。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 活頁簿", "*.xlsx")],
            initialfile="打卡紀錄.xlsx",
        )
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "打卡紀錄"
        headers = ["部門", "員工代號", "員工姓名", "出勤日期", "刷卡時間", "結轉", "來源"]
        ws.append(headers)
        hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = hdr_fill
        for iid in rows:
            ws.append(list(self.raw_tv.item(iid)["values"]))
        for i, w in enumerate([12, 10, 10, 14, 10, 8, 8], start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        wb.save(path)
        self._set_status(f"已匯出打卡紀錄 → {path}")

    def _build_settings_tab(self, tab) -> None:
        # 垂直置中：上下各放一個彈性區塊
        ctk.CTkFrame(tab, fg_color="transparent").pack(fill="both", expand=True)

        card = ctk.CTkFrame(tab, fg_color=_C["white"], corner_radius=14,
                             border_width=1, border_color=_C["line"])
        card.pack(padx=60, anchor="center")

        # ── 卡片標題 ──────────────────────────────────────────
        hdr_row = ctk.CTkFrame(card, fg_color="transparent")
        hdr_row.pack(fill="x", padx=28, pady=(24, 0))

        icon_box = ctk.CTkFrame(hdr_row, fg_color=_C["tab_bg"],
                                  width=46, height=46, corner_radius=12)
        icon_box.pack(side="left", padx=(0, 14))
        icon_box.pack_propagate(False)
        ctk.CTkLabel(icon_box, text="⚙", font=("", 20),
                      text_color=_C["accent"]).pack(expand=True)

        title_col = ctk.CTkFrame(hdr_row, fg_color="transparent")
        title_col.pack(side="left")
        ctk.CTkLabel(title_col, text="連線設定",
                      font=("", 17, "bold"), text_color=_C["ink"],
                      anchor="w").pack(anchor="w")
        ctk.CTkLabel(title_col, text="帳密只存本機 ~/.ehrs_tool，不會上傳",
                      font=("", 12), text_color=_C["muted"],
                      anchor="w").pack(anchor="w")

        # 分隔線
        tk.Frame(card, bg=_C["line"], height=1).pack(fill="x", pady=(18, 0))

        # ── 表單 ──────────────────────────────────────────────
        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(fill="x", padx=28, pady=(18, 26))

        def _field(label: str, attr: str, default: str, **kw) -> None:
            ctk.CTkLabel(form, text=label,
                          font=("", 13, "bold"), text_color=_C["ink"],
                          anchor="w").pack(anchor="w", pady=(0, 4))
            ent = ctk.CTkEntry(form, width=400, height=42, corner_radius=9,
                                border_color=_C["line"], fg_color=_C["white"],
                                text_color=_C["ink"], font=("", 13), **kw)
            ent.insert(0, default)
            ent.pack(anchor="w", pady=(0, 14))
            setattr(self, attr, ent)

        _field("帳號",   "s_acc",  self.cfg.get("account",  ""))
        _field("密碼",   "s_pwd",  self.cfg.get("password", ""), show="*")
        _field("站台網址", "s_base", self.cfg.get("base_url", DEFAULT_BASE_URL))

        ctk.CTkButton(
            form, text="儲存並登入", width=400, height=44,
            fg_color=_C["accent"], hover_color=_C["acc_h"],
            corner_radius=10, font=("", 14, "bold"),
            command=self._save_and_login,
        ).pack(anchor="w")

        ctk.CTkFrame(tab, fg_color="transparent").pack(fill="both", expand=True)

    # -------------------------------------------------------------- helpers
    def _set_status(self, text: str) -> None:
        self.status.configure(text=text)
        if not hasattr(self, "_status_dot"):
            return
        # 依訊息內容切換狀態燈顏色
        if any(k in text for k in ("錯誤", "失敗")):
            dot_clr = _C["red"]
        elif any(k in text for k in ("完成", "已載入", "已登入", "已匯出",
                                      "已寫入", "已刪除", "已套用", "安裝完成")):
            dot_clr = _C["green"]
        elif any(k in text for k in ("中…", "查詢", "載入", "匯入", "下載")):
            dot_clr = _C["accent"]
        else:
            dot_clr = _C["muted"]
        self._status_dot.configure(text_color=dot_clr)

    def _show_progress(self, done: int, total: int) -> None:
        """顯示匯入進度條（在主執行緒呼叫）。"""
        if total == 0:
            return
        self._prog_bar.set(done / total)
        self._prog_lbl.configure(text=f"{done} / {total}")
        if not self._prog_frm.winfo_ismapped():
            self._prog_frm.pack(side="right")

    def _hide_progress(self) -> None:
        """隱藏進度條。"""
        self._prog_frm.pack_forget()
        self._prog_bar.set(0)

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

    # -------------------------------------------------------------- 自動更新
    def _check_update(self) -> None:
        try:
            from updater import start_update_check
            start_update_check(
                status_cb=lambda msg: self.after(0, lambda m=msg: self._set_status(m))
            )
        except Exception:
            pass  # updater 模組不存在或匯入失敗時靜默略過

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

        if self._period_start and self._period_end:
            # ── 四週期間模式 ─────────────────────────────────────
            start, end = self._period_start, self._period_end
            period_lbl = f"{start.month}/{start.day}～{end.month}/{end.day}"
            dates = [start + dt.timedelta(days=i)
                     for i in range((end - start).days + 1)]

            # 收集期間橫跨的所有年月
            months_needed: set[tuple[int, int]] = set()
            cur = start.replace(day=1)
            while cur <= end:
                months_needed.add((cur.year, cur.month))
                cur = (cur.replace(day=28) + dt.timedelta(days=4)).replace(day=1)

            self._set_status(f"載入 {period_lbl} 排班中…")

            def work_p():
                emp_map: dict[str, object] = {}
                for yr, mn in sorted(months_needed):
                    try:
                        for emp in self.client.get_schedule(yr, mn):
                            if emp.emp_id not in emp_map:
                                emp_map[emp.emp_id] = emp
                            else:
                                existing = {s.date for s in emp_map[emp.emp_id].shifts}
                                for s in emp.shifts:
                                    if s.date not in existing:
                                        emp_map[emp.emp_id].shifts.append(s)
                    except Exception:
                        pass
                return list(emp_map.values()), dates

            def done_p(result):
                sched, dates = result
                self.schedule = sched
                self._build_code_map(sched)
                self._populate_grid_period(dates, sched)
                self._set_status(f"已載入 {period_lbl} 排班（共 {len(sched)} 人）")

            self._run_async(work_p, done_p)
        else:
            # ── 單月備用模式 ─────────────────────────────────────
            try:
                year  = int(self.year_var.get())
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
                self._set_status(f"已載入 {year}-{month:02d} 排班（共 {len(sched)} 人）")

            self._run_async(work, done)

    def _build_code_map(self, sched) -> None:
        for emp in sched:
            for s in emp.shifts:
                if s.code:
                    existing = self.shift_codes.get(s.code, ("", 3, 0.0))
                    self.shift_codes[s.code] = (
                        s.name or existing[0],
                        s.kind if s.kind is not None else 3,
                        existing[2] if len(existing) > 2 else 0.0,
                    )

    def _populate_grid(self, year: int, month: int, sched) -> None:
        """單月排班格（備用，import 結束後呼叫）。"""
        self._grid_dates = []   # 清除期間模式
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

    def _populate_grid_period(self, dates: list[dt.date], sched) -> None:
        """四週期間排班格：以連續日期清單建立 28 欄。"""
        self._grid_dates = list(dates)
        cols = ["emp"] + [f"dd{i}" for i in range(len(dates))]
        tv = self.grid_tv
        tv.delete(*tv.get_children())
        tv["columns"] = cols
        tv.heading("emp", text="員工")
        tv.column("emp", width=150, anchor="w", stretch=False)
        for i, d in enumerate(dates):
            wd = _WEEK[d.weekday()]
            tv.heading(f"dd{i}", text=f"{d.month}/{d.day}\n{wd}")
            tv.column(f"dd{i}", width=52, anchor="center", stretch=False)
        self.row_emp = {}
        for emp in sched:
            by_date = {s.date: s.code for s in emp.shifts}
            values  = [f"{emp.emp_id} {emp.name}"] + [
                by_date.get(d.isoformat(), "") for d in dates
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
        emp = self.row_emp.get(row)
        if emp is None:
            return
        col_idx = int(col[1:]) - 2   # "#2" → 0, "#3" → 1, ...
        if col_idx < 0:
            return
        if self._grid_dates:
            # ── 四週期間模式 ─────────────────────────────
            if col_idx >= len(self._grid_dates):
                return
            date_str = self._grid_dates[col_idx].isoformat()
        else:
            # ── 單月備用模式 ─────────────────────────────
            day = col_idx + 1
            year  = int(self.year_var.get())
            month = int(self.month_var.get())
            date_str = f"{year:04d}-{month:02d}-{day:02d}"
        self._open_editor(emp, date_str)

    def _open_editor(self, emp, date_str: str) -> None:
        cur = emp.shift_on(date_str)

        top = ctk.CTkToplevel(self)
        top.title("編輯排班")
        top.geometry("400x480")
        top.configure(fg_color=_C["app_bg"])
        top.transient(self)
        top.after(60, top.grab_set)

        # 員工資訊列
        hdr_row = ctk.CTkFrame(top, fg_color=_C["white"], corner_radius=10,
                                 border_width=1, border_color=_C["line"])
        hdr_row.pack(fill="x", padx=16, pady=(16, 0))
        ctk.CTkLabel(hdr_row, text=f"{emp.emp_id} {emp.name}",
                      font=("", 15, "bold"), text_color=_C["ink"]).pack(
            side="left", padx=14, pady=10)
        ctk.CTkLabel(hdr_row, text=date_str,
                      font=("", 13), text_color=_C["muted"]).pack(
            side="left", padx=4)
        cur_txt = f"目前：{cur.code} {cur.name}" if cur else "目前：（空白）"
        ctk.CTkLabel(hdr_row, text=cur_txt,
                      font=("", 12), text_color=_C["muted"]).pack(
            side="right", padx=14)

        # 預先計算所有班別的上班/下班/時數
        opt_rows: list[tuple[str, str, str, str]] = []  # (code, start, end, hours)
        for c in sorted(self.shift_codes):
            si, so = _shift_times(c)
            h = _shift_hours(c) or (_calc_hours(si, so) if si and so else 0.0)
            h_str = f"{h:g}" if h else ""
            opt_rows.append((c, si, so, h_str))

        entry = ctk.CTkEntry(top, width=368, placeholder_text="輸入班別代號…",
                               border_color=_C["line"], fg_color=_C["white"],
                               text_color=_C["ink"])
        entry.pack(pady=(10, 4), padx=16)
        if cur:
            entry.insert(0, cur.code)

        tv_wrap = tk.Frame(top, bg=_C["app_bg"])
        tv_wrap.pack(fill="x", padx=16, pady=(0, 4))
        tv = ttk.Treeview(tv_wrap, style="App.Treeview",
                           columns=("code", "start", "end", "hours"),
                           show="headings", height=10, selectmode="browse")
        tv.heading("code",  text="班別代號"); tv.column("code",  width=90, anchor="center")
        tv.heading("start", text="上班時間"); tv.column("start", width=85, anchor="center")
        tv.heading("end",   text="下班時間"); tv.column("end",   width=85, anchor="center")
        tv.heading("hours", text="時數");     tv.column("hours", width=60, anchor="center")
        sb = ttk.Scrollbar(tv_wrap, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def _fill(typed: str = "") -> None:
            tv.delete(*tv.get_children())
            t = typed.lower()
            if not t or any(k in t for k in ("空白", "刪除")):
                tv.insert("", "end", iid="__del__", values=("(空白/刪除)", "", "", ""))
            for code, si, so, h in opt_rows:
                name = self.shift_codes.get(code, ("", 3))[0]
                if not t or t in code.lower() or t in name.lower() or t in si or t in so:
                    tv.insert("", "end", values=(code, si, so, h))

        _fill()

        def _pick() -> None:
            sel = tv.selection()
            if sel:
                code = str(tv.item(sel[0])["values"][0])
                entry.delete(0, "end")
                if code != "(空白/刪除)":
                    entry.insert(0, code)

        def _on_entry_key(event) -> None:
            if event.keysym == "Down":
                tv.focus_set()
                children = tv.get_children()
                if children:
                    tv.selection_set(children[0])
                    tv.focus(children[0])
                return
            if event.keysym == "Return":
                save(); return
            _fill(entry.get())

        def _on_tv_key(event) -> None:
            if event.keysym == "Return":
                _pick(); save()
            elif event.keysym == "Up":
                sel = tv.selection()
                children = tv.get_children()
                if sel and children and sel[0] == children[0]:
                    entry.focus_set()

        entry.bind("<KeyRelease>", _on_entry_key)
        tv.bind("<<TreeviewSelect>>", lambda e: _pick())
        tv.bind("<KeyRelease>", _on_tv_key)
        entry.focus_set()

        def save() -> None:
            raw = entry.get().strip()
            top.destroy()
            if not raw or raw.startswith("(空白"):
                self._apply_delete(emp, date_str, cur)
            else:
                self._apply_set(emp, date_str, raw.split()[0])

        btn_row = ctk.CTkFrame(top, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkButton(btn_row, text="取消", width=170,
                       fg_color=_C["line"], hover_color="#D0D4E4",
                       text_color=_C["ink"], corner_radius=9,
                       command=top.destroy).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row, text="儲存", width=170,
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=9, font=("", 13, "bold"),
                       command=save).pack(side="left")

    def _apply_set(self, emp, date_str: str, code: str) -> None:
        year, month, _ = (int(x) for x in date_str.split("-"))
        self._set_status("寫入中…")
        # 排班一律 kind=3；只有排休（991）才讓系統自動偵測 kind
        kind = None if _is_rest(code) else 3

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

    # --------------------------------------------------------- export / import
    def _export_schedule(self) -> None:
        if not self.schedule:
            messagebox.showwarning("尚未載入", "請先載入排班。")
            return
        # 決定要匯出的日期清單
        if self._grid_dates:
            dates_obj = self._grid_dates
            s, e = dates_obj[0], dates_obj[-1]
            default_name = f"排班_{s.year}-{s.month:02d}{s.day:02d}_{e.month:02d}{e.day:02d}.xlsx"
            sheet_title  = f"{s.month}/{s.day}～{e.month}/{e.day}"
        else:
            year, month = int(self.year_var.get()), int(self.month_var.get())
            ndays       = _cal.monthrange(year, month)[1]
            dates_obj   = [dt.date(year, month, d) for d in range(1, ndays + 1)]
            default_name = f"排班_{year}-{month:02d}.xlsx"
            sheet_title  = f"{year}-{month:02d}"

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 活頁簿", "*.xlsx")],
            initialfile=default_name,
        )
        if not path:
            return

        dates_iso = [d.isoformat() for d in dates_obj]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        ws.append(["員工代號", "員工姓名"] + dates_iso)
        hdr_fill   = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        wkend_fill = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
        for cell in ws[1]:
            cell.font      = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.fill      = hdr_fill
        ws.row_dimensions[1].height = 36
        for col_i, d in enumerate(dates_obj, start=3):
            if d.weekday() >= 5:
                ws.cell(1, col_i).fill = wkend_fill
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 10
        for i in range(3, len(dates_obj) + 3):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 7
        for emp in self.schedule:
            by_day = {s.date: s.code for s in emp.shifts}
            ws.append([emp.emp_id, emp.name] + [by_day.get(d, "") for d in dates_iso])
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
            row[0].alignment = Alignment(horizontal="left")
            row[1].alignment = Alignment(horizontal="left")
        wb.save(path)
        self._set_status(f"已匯出排班 → {path}")

    def _export_punch(self) -> None:
        data = self.punch_tbl.all_values()
        if not data:
            messagebox.showwarning("無資料", "請先查詢打卡。")
            return
        start = self.p_start.get().strip()
        end = self.p_end.get().strip()
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 活頁簿", "*.xlsx")],
            initialfile=f"打卡_{start}_{end}.xlsx",
        )
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "打卡明細"
        headers = ["員工代號", "員工姓名", "出勤日期", "排班代號", "表定上班", "表定下班", "實際上班", "實際下班", "備注"]
        ws.append(headers)
        hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = hdr_fill
        for row in data:
            ws.append(list(row))
        for i, w in enumerate([10, 10, 14, 10, 10, 10, 10, 10, 18], start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        wb.save(path)
        self._set_status(f"已匯出 {len(data)} 筆打卡 → {path}")

    def _import_schedule(self) -> None:
        if not self._need_client():
            return
        path = filedialog.askopenfilename(
            filetypes=[("Excel 活頁簿", "*.xlsx *.xls")],
            title="選擇排班 Excel",
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as exc:
            messagebox.showerror("讀取失敗", str(exc))
            return
        if len(rows) < 2:
            messagebox.showwarning("空檔", "Excel 沒有資料列。")
            return
        # 第一列找日期欄（YYYY-MM-DD 格式）
        header = rows[0]
        date_cols: dict[int, str] = {}
        for i, h in enumerate(header):
            h_str = str(h).strip() if h is not None else ""
            if len(h_str) == 10 and h_str[4] == "-" and h_str[7] == "-":
                try:
                    dt.date.fromisoformat(h_str)
                    date_cols[i] = h_str
                except ValueError:
                    pass
        if not date_cols:
            messagebox.showerror(
                "格式錯誤",
                "找不到日期欄（YYYY-MM-DD）。\n請使用本工具匯出的 Excel 再匯入。",
            )
            return
        first_date = next(iter(date_cols.values()))
        year, month = int(first_date[:4]), int(first_date[5:7])
        # 收集有班別的格子，同一 emp+date 取最後一筆
        seen: dict[tuple, dict] = {}
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            emp_id = str(row[0]).strip()
            for col_idx, date_str in date_cols.items():
                raw = row[col_idx] if col_idx < len(row) else None
                code = str(raw).strip() if raw is not None else ""
                if code:
                    seen[(emp_id, date_str)] = {
                        "emp_id": emp_id, "date": date_str, "shift_code": code
                    }
        changes = list(seen.values())
        if not changes:
            messagebox.showinfo("無班別", "沒有找到任何班別資料。")
            return
        if not messagebox.askyesno(
            "確認匯入",
            f"將匯入 {year}-{month:02d} 排班，共 {len(changes)} 筆。\n\n"
            "這會寫入正式班表，確定？",
        ):
            return
        self._set_status(f"匯入中…")
        self._show_progress(0, len(changes))

        def work():
            for ch in changes:
                if not _is_rest(ch["shift_code"]):
                    ch.setdefault("kind", 3)

            def on_progress(done_n: int, total: int) -> None:
                self.after(0, lambda d=done_n, t=total: (
                    self._show_progress(d, t),
                    self._set_status(f"匯入中  {d} / {t}"),
                ))

            return self.client.set_shifts_bulk(
                year, month, changes, dry_run=False, progress_cb=on_progress
            )

        def done(results):
            self._hide_progress()
            ok_n = sum(1 for r in results if r.get("ok"))
            fail_n = len(results) - ok_n
            self._set_status(f"匯入完成：{ok_n} 筆成功，{fail_n} 筆失敗")
            if fail_n:
                errs = [
                    f"{r['emp_id']} {r['date']}：{r.get('error', '未知')}"
                    for r in results if not r.get("ok")
                ]
                messagebox.showwarning(
                    "部分失敗",
                    f"{fail_n} 筆寫入失敗：\n" + "\n".join(errs[:10])
                    + ("\n…" if fail_n > 10 else ""),
                )
            self._load_schedule()

        self._run_async(work, done)

    # ---------------------------------------------------------------- punch
    def _query_punch(self) -> None:
        if not self._need_client():
            return
        start = self.p_start.get().strip()
        end = self.p_end.get().strip()
        self._set_status("查詢打卡中…")

        def work():
            punch_rows = self.client.get_punch_records(
                start, end, readable=True
            )
            sched_map: dict[tuple, tuple] = {}
            emp_names: dict[str, str] = {}
            start_d = dt.date.fromisoformat(start)
            end_d   = dt.date.fromisoformat(end)
            cur = start_d.replace(day=1)
            while cur <= end_d:
                try:
                    for es in self.client.get_schedule(cur.year, cur.month):
                        emp_names[es.emp_id] = es.name
                        for s in es.shifts:
                            sched_map[(es.emp_id, s.date)] = (s.code, s.name)
                except Exception:
                    pass
                cur = (cur.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
            return punch_rows, sched_map, emp_names, start, end

        def done(result):
            punch_rows, sched_map, emp_names, s, e = result
            self._raw_punch_rows = list(punch_rows)
            self._populate_punch(punch_rows, sched_map, emp_names, s, e)
            self._populate_raw_punch()
            self._set_status(f"打卡 {start}~{end} 共 {len(punch_rows)} 筆")

        self._run_async(work, done)

    def _populate_punch(
        self,
        rows,
        sched_map: dict | None = None,
        emp_names: dict | None = None,
        range_start: str | None = None,
        range_end: str | None = None,
    ) -> None:
        # 0. 建立 emp_id -> 部門名稱 對照表（從原始打卡資料）
        for r in rows:
            eid = r.get("員工代號", "")
            dept = r.get("部門名稱") or r.get("pa51014FullName", "")
            if eid and dept:
                self._emp_dept[eid] = dept

        # 1. 整理有打卡的日子
        sorted_rows = sorted(rows, key=lambda r: (
            r.get("員工代號", ""),
            str(r.get("出勤日期", ""))[:10],
            str(r.get("刷卡時間", "")),
        ))
        display: list[tuple] = []
        punched: set[tuple] = set()
        for (emp_id, date), grp in groupby(sorted_rows, key=lambda r: (
            r.get("員工代號", ""), str(r.get("出勤日期", ""))[:10]
        )):
            punched.add((emp_id, date))
            group = list(grp)
            name = group[0].get("員工姓名", "")
            code, _ = (sched_map or {}).get((emp_id, date), ("", ""))
            sched_in, sched_out = _shift_times(code) if code else ("", "")
            # 提取打卡時間，區分有無結轉標記
            all_punches = [(str(r.get("刷卡時間", ""))[11:16],
                            r.get("結轉註記", "")) for r in group]
            jiezhuan = [t for t, flag in all_punches if flag == "結轉" and t]
            all_times  = [t for t, _  in all_punches if t]
            # 有結轉打卡 → 優先使用，用寬鬆分群；否則用全部打卡與原分群邏輯
            if jiezhuan:
                use_punches = jiezhuan
                split_fn = _split_punches_jiezhuan
            else:
                use_punches = all_times
                split_fn = _split_punches
            if len(use_punches) == 1:
                clock_in, clock_out = _assign_single_punch(use_punches[0], sched_in, sched_out)
            else:
                clock_in, clock_out = split_fn(use_punches, sched_in, sched_out)
            remark = _punch_remark(clock_in, clock_out, sched_in, sched_out)
            display.append((emp_id, name, date, code, sched_in, sched_out, clock_in, clock_out, remark))

        # 2. 找出有排班但完全沒打卡的日子 → 曠職
        if sched_map and range_start and range_end:
            start_d = dt.date.fromisoformat(range_start)
            end_d   = dt.date.fromisoformat(range_end)
            for (emp_id, date), (code, _) in sched_map.items():
                if not code or _is_rest(code):
                    continue
                d = dt.date.fromisoformat(date)
                if start_d <= d <= end_d and (emp_id, date) not in punched:
                    name = (emp_names or {}).get(emp_id, "")
                    sched_in, sched_out = _shift_times(code) if code else ("", "")
                    display.append((emp_id, name, date, code, sched_in, sched_out, "", "", "曠職"))

        RED = "#CC0000"
        def _fgs(row: tuple) -> tuple:
            remark = row[8]
            ci_red = RED if remark and any(k in remark for k in ("遲到", "上班忘打卡", "曠職")) else ""
            co_red = RED if remark and any(k in remark for k in ("早退", "下班忘打卡", "曠職")) else ""
            rm_red = RED if remark else ""
            return ("", "", "", "", "", "", ci_red, co_red, rm_red)

        self._punch_cache = [
            (row, _fgs(row))
            for row in sorted(display, key=lambda x: (x[0], x[2]))
        ]
        self._apply_punch_filter()

    def _apply_punch_filter(self) -> None:
        only_abnormal = self.p_abnormal.get()
        filtered = [
            (vals, fgs) for vals, fgs in self._punch_cache
            if (not self._selected_depts or
                self._emp_dept.get(vals[0], "") in self._selected_depts)
            and (not self._selected_emps or vals[0] in self._selected_emps)
            and (not only_abnormal or any(fgs))
        ]
        self.punch_tbl.populate(filtered)
        self.after(50, self._compute_suggestions)

    def _update_dept_btn(self) -> None:
        n = len(self._selected_depts)
        self.p_dept_btn.configure(
            text="部門：全部 ▼" if n == 0 else f"部門：已選 {n} ▼"
        )

    def _update_emp_btn(self) -> None:
        n = len(self._selected_emps)
        self.p_emp_btn.configure(
            text="員工：全員 ▼" if n == 0 else f"員工：已選 {n} 人 ▼"
        )

    def _open_dept_picker(self) -> None:
        # 從 _emp_dept 收集所有不重複部門
        depts = sorted(set(self._emp_dept.values()))
        if not depts:
            messagebox.showinfo("提示", "請先查詢打卡資料。")
            return

        top = ctk.CTkToplevel(self)
        top.title("選擇部門")
        top.geometry("260x400")
        top.configure(fg_color=_C["app_bg"])
        top.resizable(False, True)
        top.transient(self)
        top.after(60, top.grab_set)

        scroll = ctk.CTkScrollableFrame(top, fg_color=_C["white"],
                                          corner_radius=8,
                                          border_width=1, border_color=_C["line"])
        scroll.pack(fill="both", expand=True, padx=12, pady=(12, 0))

        chk_vars: dict[str, tk.BooleanVar] = {}
        for dept in depts:
            init = (not self._selected_depts) or (dept in self._selected_depts)
            var = tk.BooleanVar(value=init)
            ctk.CTkCheckBox(scroll, text=dept, variable=var,
                             text_color=_C["ink"]).pack(anchor="w", padx=4, pady=3)
            chk_vars[dept] = var

        btn_bar = ctk.CTkFrame(top, fg_color="transparent")
        btn_bar.pack(fill="x", padx=12, pady=(6, 0))
        ctk.CTkButton(btn_bar, text="全選", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(True) for v in chk_vars.values()]
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_bar, text="清除", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(False) for v in chk_vars.values()]
                       ).pack(side="left")

        def apply():
            selected = {d for d, v in chk_vars.items() if v.get()}
            self._selected_depts = set() if len(selected) == len(depts) else selected
            self._update_dept_btn()
            self._apply_punch_filter()
            top.destroy()

        ctk.CTkButton(top, text="確定",
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=9, font=("", 13, "bold"),
                       command=apply).pack(fill="x", padx=12, pady=10)

    def _open_emp_picker(self) -> None:
        # 從打卡快取收集所有員工
        emps: dict[str, str] = {}
        for vals, _ in self._punch_cache:
            emps[vals[0]] = vals[1]

        if not emps:
            messagebox.showinfo("提示", "請先查詢打卡資料。")
            return

        top = ctk.CTkToplevel(self)
        top.title("選擇員工")
        top.geometry("280x490")
        top.configure(fg_color=_C["app_bg"])
        top.resizable(False, True)
        top.transient(self)
        top.after(60, top.grab_set)

        # 搜尋框
        search_var = ctk.StringVar()
        ctk.CTkEntry(top, placeholder_text="搜尋員工…",
                      textvariable=search_var,
                      border_color=_C["line"], fg_color=_C["white"],
                      text_color=_C["ink"]).pack(fill="x", padx=12, pady=(12, 4))

        # 捲動式 checkbox 清單
        scroll = ctk.CTkScrollableFrame(top, fg_color=_C["white"],
                                          corner_radius=8,
                                          border_width=1, border_color=_C["line"])
        scroll.pack(fill="both", expand=True, padx=12, pady=(4, 0))

        chk_vars: dict[str, tk.BooleanVar] = {}
        chk_widgets: dict[str, ctk.CTkCheckBox] = {}
        for eid in sorted(emps):
            init = (not self._selected_emps) or (eid in self._selected_emps)
            var = tk.BooleanVar(value=init)
            chk = ctk.CTkCheckBox(scroll, text=f"{eid}  {emps[eid]}",
                                   variable=var, text_color=_C["ink"])
            chk.pack(anchor="w", padx=4, pady=3)
            chk_vars[eid] = var
            chk_widgets[eid] = chk

        def _filter(*_):
            q = search_var.get().lower()
            for eid, w in chk_widgets.items():
                show = not q or q in eid.lower() or q in emps[eid].lower()
                if show:
                    w.pack(anchor="w", padx=4, pady=2)
                else:
                    w.pack_forget()

        search_var.trace_add("write", _filter)

        # 全選 / 清除
        btn_bar = ctk.CTkFrame(top, fg_color="transparent")
        btn_bar.pack(fill="x", padx=12, pady=(6, 0))
        ctk.CTkButton(btn_bar, text="全選", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(True) for v in chk_vars.values()]
                       ).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_bar, text="清除", width=80,
                       fg_color=_C["tab_bg"], text_color=_C["ink"],
                       hover_color=_C["line"], corner_radius=8,
                       command=lambda: [v.set(False) for v in chk_vars.values()]
                       ).pack(side="left")

        def apply():
            selected = {eid for eid, v in chk_vars.items() if v.get()}
            self._selected_emps = set() if len(selected) == len(emps) else selected
            self._update_emp_btn()
            self._apply_punch_filter()
            top.destroy()

        ctk.CTkButton(top, text="確定",
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=9, font=("", 13, "bold"),
                       command=apply).pack(fill="x", padx=12, pady=10)

    # --------------------------------------------------------- suggest tab
    def _build_suggest_tab(self, tab) -> None:
        bar = ctk.CTkFrame(tab, fg_color=_C["white"], corner_radius=8,
                            border_width=1, border_color=_C["line"])
        bar.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(bar, text="依打卡查詢結果建議最適班別（忘打卡不列入）",
                      text_color=_C["muted"], font=("", 12)).pack(
            side="left", padx=12, pady=8)
        self.p_overtime = ctk.CTkCheckBox(bar, text="加班建議",
                                            text_color=_C["ink"],
                                            command=self._compute_suggestions)
        self.p_overtime.pack(side="left", padx=12)
        ctk.CTkButton(bar, text="重新計算", width=90,
                       fg_color=_C["accent"], hover_color=_C["acc_h"],
                       corner_radius=8,
                       command=self._compute_suggestions).pack(
            side="left", padx=8, pady=6)
        self.suggest_scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self.suggest_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def _compute_suggestions(self) -> None:
        if not hasattr(self, "suggest_scroll"):
            return
        for w in self.suggest_scroll.winfo_children():
            w.destroy()

        # 遲到/早退
        late_early = [
            vals for vals, _ in self._punch_cache
            if vals[8]
            and "忘打卡" not in vals[8]
            and any(k in vals[8] for k in ("遲到", "早退"))
        ]
        # 加班（勾選才顯示）
        overtime: list[tuple] = []
        if self.p_overtime.get():
            for vals, _ in self._punch_cache:
                _, _, _, _, sched_in, sched_out, clock_in, clock_out, remark = vals
                if not clock_in or not clock_out or not sched_in or not sched_out:
                    continue
                if "忘打卡" in remark:
                    continue
                si_m = _punch_mins(sched_in)
                so_m = _punch_mins(sched_out)
                co_m = _punch_mins(clock_out)
                if so_m < si_m:
                    so_m += 1440
                if co_m < si_m - 120:
                    co_m += 1440
                if co_m - so_m >= 30:   # 加班 ≥ 30 分鐘才列入
                    overtime.append(vals)

        if not late_early and not overtime:
            ctk.CTkLabel(self.suggest_scroll,
                         text="沒有需要建議的異常記錄（請先到打卡分頁查詢）",
                         text_color=_C["muted"]).pack(pady=24)
            return

        if late_early:
            ctk.CTkLabel(self.suggest_scroll, text="遲到 / 早退",
                         font=("", 13, "bold"), text_color=_C["ink"]).pack(
                anchor="w", padx=6, pady=(4, 0))
            for vals in late_early:
                self._build_suggest_card(vals, overtime_mode=False)

        if overtime:
            ctk.CTkLabel(self.suggest_scroll, text="加班",
                         font=("", 13, "bold"), text_color=_C["ink"]).pack(
                anchor="w", padx=6, pady=(10, 0))
            for vals in overtime:
                self._build_suggest_card(vals, overtime_mode=True)

    def _build_suggest_card(self, vals: tuple, overtime_mode: bool = False) -> None:
        emp_id, name, date, shift_code, sched_in, sched_out, \
            clock_in, clock_out, remark = vals
        cur_h = _shift_hours(shift_code) or (_calc_hours(sched_in, sched_out) if sched_in and sched_out else 8.0)
        if overtime_mode:
            if clock_in and clock_out:
                ci_m = _punch_mins(clock_in)
                co_m = _punch_mins(clock_out)
                si_m = _punch_mins(sched_in) if sched_in else 0
                if co_m < si_m - 120:
                    co_m += 1440
                raw_h = (co_m - ci_m) / 60.0
                actual_h = raw_h - (1.0 if raw_h >= 9.0 else 0.5)
            else:
                actual_h = cur_h + 1.0
            pref_h = round(actual_h * 2) / 2   # 四捨五入到 0.5
            sugg = [s for s in _suggest_shifts(
                clock_in, clock_out, self.shift_codes, preferred_hours=pref_h, top_n=6)
                if s[5] > cur_h][:3]
            hdr_color = "#1E6B42"   # 深綠
            tag_text  = "加班建議"
        else:
            sugg = _suggest_shifts(clock_in, clock_out, self.shift_codes, preferred_hours=cur_h)
            hdr_color = "#8B4500"   # 深橙棕
            tag_text  = remark

        outer = ctk.CTkFrame(self.suggest_scroll, corner_radius=10,
                              fg_color=_C["white"],
                              border_width=1, border_color=_C["line"])
        outer.pack(fill="x", pady=5, padx=2)

        # ── 標題列 ──────────────────────────────────────────
        hdr = ctk.CTkFrame(outer, fg_color=hdr_color, corner_radius=10)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text=f"{name}  {emp_id}",
                      font=("", 13, "bold"), text_color="#FFFFFF").pack(
            side="left", padx=14, pady=8)
        ctk.CTkLabel(hdr, text=date,
                      text_color="#FFDDC0", font=("", 12)).pack(side="left", padx=4)
        ctk.CTkLabel(hdr, text=tag_text,
                      text_color="#FFE680", font=("", 12, "bold")).pack(
            side="right", padx=14)

        # ── 應排班 vs 實際打卡 ────────────────────────────────
        info = ctk.CTkFrame(outer, fg_color="transparent")
        info.pack(fill="x", padx=10, pady=(8, 4))

        cur_f = ctk.CTkFrame(info, fg_color="#FFF4EE", corner_radius=8,
                              border_width=1, border_color="#FDDDC9")
        cur_f.pack(side="left", fill="both", expand=True, padx=(0, 4))
        ctk.CTkLabel(cur_f, text="應排班",
                      text_color=_C["muted"], font=("", 11)).pack(
            anchor="w", padx=10, pady=(8, 0))
        ctk.CTkLabel(cur_f, text=shift_code or "—",
                      text_color=_C["accent"], font=("", 15, "bold")).pack(
            anchor="w", padx=10)
        code_name = self.shift_codes.get(shift_code, ("", 3))[0] if shift_code else ""
        ctk.CTkLabel(cur_f, text=code_name or f"{sched_in}–{sched_out}",
                      font=("", 11), text_color=_C["ink"]).pack(anchor="w", padx=10)
        if shift_code or (sched_in and sched_out):
            disp_h = _shift_hours(shift_code) or (
                _calc_hours(sched_in, sched_out) if sched_in and sched_out else 0.0)
            if disp_h:
                ctk.CTkLabel(cur_f, text=f"{disp_h:g}h",
                              text_color=_C["muted"]).pack(
                    anchor="w", padx=10, pady=(0, 8))

        act_f = ctk.CTkFrame(info, fg_color="#EEF4FF", corner_radius=8,
                              border_width=1, border_color="#CCDDFF")
        act_f.pack(side="left", fill="both", expand=True, padx=(4, 0))
        ctk.CTkLabel(act_f, text="實際打卡",
                      text_color=_C["muted"], font=("", 11)).pack(
            anchor="w", padx=10, pady=(8, 0))
        punch_txt = (f"{clock_in}  →  {clock_out}"
                     if clock_in and clock_out else clock_in or clock_out or "—")
        ctk.CTkLabel(act_f, text=punch_txt,
                      font=("", 14, "bold"), text_color=_C["ink"]).pack(
            anchor="w", padx=10, pady=(4, 8))

        # ── 建議班別 ─────────────────────────────────────────
        if not sugg:
            ctk.CTkLabel(outer, text="找不到合適的班別建議",
                          text_color=_C["muted"]).pack(pady=8)
        for i, (code, si, so, d_in, d_out, h) in enumerate(sugg):
            row_bg = _C["row_alt"] if i % 2 == 0 else _C["white"]
            row = ctk.CTkFrame(outer, fg_color=row_bg, corner_radius=6)
            row.pack(fill="x", padx=8, pady=2)

            rank_txt = "★" if i == 0 else str(i + 1)
            ctk.CTkLabel(row, text=rank_txt, width=24,
                          text_color="#D08000" if i == 0 else _C["muted"],
                          font=("", 13, "bold")).pack(side="left", padx=6, pady=5)

            s_name = self.shift_codes.get(code, ("", 3))[0]
            ctk.CTkLabel(row, text=f"{code}  {s_name or f'{si}–{so}'}",
                          width=210, anchor="w",
                          text_color=_C["ink"]).pack(side="left", padx=4)
            ctk.CTkLabel(row, text=f"{h:g}h",
                          width=36, text_color=_C["muted"]).pack(side="left")

            parts: list[str] = []
            if d_in > 0:
                parts.append(f"早{d_in}分")
            if d_out > 0:
                parts.append(f"晚{d_out}分")
            elif d_out < 0:
                parts.append(f"早退{-d_out}分")
            ctk.CTkLabel(row, text="·".join(parts),
                          text_color="#D07000", width=110).pack(side="left", padx=4)

            ctk.CTkButton(
                row, text="套用", width=60, height=28,
                fg_color=_C["accent"], hover_color=_C["acc_h"],
                corner_radius=6, font=("", 12),
                command=lambda eid=emp_id, dt=date, c=code: self._apply_suggestion(eid, dt, c)
            ).pack(side="right", padx=8, pady=5)

    def _apply_suggestion(self, emp_id: str, date_str: str, code: str) -> None:
        if not self._need_client():
            return
        year, month, _ = (int(x) for x in date_str.split("-"))
        self._set_status(f"套用中：{emp_id} {date_str} → {code}…")

        def work():
            return self.client.set_shift(
                year, month, emp_id, date_str, code,
                kind=None if _is_rest(code) else 3,
                dry_run=False,
            )

        def done(_):
            self._set_status(f"已套用：{emp_id} {date_str} → {code}")

        self._run_async(work, done)


def main() -> None:
    EhrsApp().mainloop()


if __name__ == "__main__":
    main()
