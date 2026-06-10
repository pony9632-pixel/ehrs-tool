"""自訂 UI 元件與滾輪綁定（從 app.py 抽出）。

包含：UI 配色字典 _C、跨平台滾輪/觸控板事件處理、
排班雙排表頭 _ScheduleHeader、排班格 _ScheduleGrid、
打卡表 _PunchTable、建議列表 _SuggestionTable。
"""
import datetime as dt
import tkinter as tk
from tkinter import ttk

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

_WHEEL_ACTIVE_TARGET = None
_WHEEL_ACTIVE_WIDGET = None
_WHEEL_BOUND = False


def _wheel_units(event) -> int:
    """把滾輪 / 觸控板事件換算成 yview_scroll 的 units（負值=往上）。
    跨平台：
      - macOS：delta 為小整數（±1、±2…），直接當單位數，不可除以 120/60。
      - Windows：delta 為 ±120 的倍數。
      - Linux/X11：無 delta，改用 Button-4（上）/ Button-5（下）。"""
    d = getattr(event, "delta", 0)
    if d:
        if abs(d) >= 120:                       # Windows / 部分 X11
            return int(-d / 120) or (-1 if d > 0 else 1)
        return -d                                # macOS：delta 已是單位數
    num = getattr(event, "num", 0)
    if num == 4:
        return -1
    if num == 5:
        return 1
    if num == 6:
        return -1
    if num == 7:
        return 1
    return 0


def _wheel_axis(event) -> str:
    """判斷滾輪事件要導向垂直或水平捲動。"""
    num = getattr(event, "num", 0)
    if num in (6, 7):
        return "x"
    # Tk 在多數平台把水平滾輪/觸控板事件映射成 Shift-MouseWheel。
    if getattr(event, "state", 0) & 0x0001:
        return "x"
    return "y"


def _scroll_target(target, event) -> bool:
    n = _wheel_units(event)
    if not n:
        return False
    axis = _wheel_axis(event)
    try:
        if axis == "x" and hasattr(target, "xview_scroll"):
            target.xview_scroll(n, "units")
        elif axis == "x" and hasattr(target, "xview"):
            target.xview("scroll", n, "units")
        elif hasattr(target, "yview_scroll"):
            target.yview_scroll(n, "units")
        else:
            target.yview("scroll", n, "units")
    except Exception:
        return False
    return True


def _bind_mousewheel(widget, target=None) -> None:
    """讓 widget 區域內的上下滾輪/觸控板捲動 target（預設為 widget 本身）。

    為何用全域 active target 而非每個 widget 都 bind_all / unbind_all：
    在 macOS + CustomTkinter 的巢狀版面中，滾輪事件不一定會送達被捲動的
    那個 widget（可能落到外層 CTk frame），導致只有少數頁面能捲。改採
    「全域只綁一次，指標進入時切換目前 target」的做法，避免離開某個表格時
    unbind_all 清掉 CustomTkinter 自己的觸控板捲動綁定。
    跨平台：macOS delta 為小整數、Windows 為 ±120 倍數、Linux 用 Button-4/5。
    也支援 Shift/水平滾輪導向 xview_scroll。"""
    global _WHEEL_BOUND
    tgt = target if target is not None else widget

    def _handler(event):
        if _scroll_target(tgt, event):
            return "break"
        return None

    def _global_handler(event):
        if _WHEEL_ACTIVE_TARGET is None:
            return None
        if _scroll_target(_WHEEL_ACTIVE_TARGET, event):
            return "break"
        return None

    def _on_enter(_e):
        global _WHEEL_ACTIVE_TARGET, _WHEEL_ACTIVE_WIDGET
        _WHEEL_ACTIVE_TARGET = tgt
        _WHEEL_ACTIVE_WIDGET = widget

    def _on_leave(_e):
        global _WHEEL_ACTIVE_TARGET, _WHEEL_ACTIVE_WIDGET
        if _WHEEL_ACTIVE_WIDGET is widget:
            _WHEEL_ACTIVE_TARGET = None
            _WHEEL_ACTIVE_WIDGET = None

    def _safe_bind(seq: str, callback, *, bind_all: bool = False) -> None:
        try:
            if bind_all:
                widget.bind_all(seq, callback, add="+")
            else:
                widget.bind(seq, callback)
        except tk.TclError:
            # 部分 Tk build 不支援 Button-6/7；略過即可，垂直捲動仍可用。
            pass

    if not _WHEEL_BOUND:
        for seq in (
            "<MouseWheel>", "<Shift-MouseWheel>",
            "<Button-4>", "<Button-5>", "<Button-6>", "<Button-7>",
        ):
            _safe_bind(seq, _global_handler, bind_all=True)
        _WHEEL_BOUND = True

    # 直接綁在 widget 上作為後備（事件確實送達時即可用）
    for seq in (
        "<MouseWheel>", "<Shift-MouseWheel>",
        "<Button-4>", "<Button-5>", "<Button-6>", "<Button-7>",
    ):
        _safe_bind(seq, _handler)
    # 指標進入/離開時切換 app 層級綁定，涵蓋事件未直接送達 widget 的情況
    widget.bind("<Enter>", _on_enter)
    widget.bind("<Leave>", _on_leave)


def _bind_scrollable_frame_mousewheel(scrollable) -> None:
    """補強 CTkScrollableFrame 的觸控板/滾輪事件，不覆蓋 CTk 內建綁定。"""
    target = (
        getattr(scrollable, "_parent_canvas", None)
        or getattr(scrollable, "_canvas", None)
        or scrollable
    )
    _bind_mousewheel(scrollable, target=target)


class _ScheduleHeader(tk.Canvas):
    """自訂雙排表頭：上列顯示月/日，下列顯示星期，六日紅字底色。
    與排班 Treeview 水平捲動同步（由外部呼叫 sync_x）。"""
    DATE_H  = 22
    WD_H    = 18
    TOTAL_H = DATE_H + WD_H   # 40 px
    EMP_W   = 150
    HOURS_W = 60    # 總時數欄
    COL_W   = 52
    _BG     = "#EEF2FB"
    _WKBG   = "#FFF0F0"   # 六日背景
    _GRID   = "#E8EBF4"
    _INK    = "#2D3048"
    _RED    = "#CC2200"   # 六日文字

    def __init__(self, parent, **kw):
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("bd", 0)
        kw.setdefault("bg", self._BG)
        kw.setdefault("height", self.TOTAL_H)
        super().__init__(parent, **kw)
        self._dates: list[dt.date] = []

    def set_dates(self, dates: list[dt.date]) -> None:
        self._dates = list(dates)
        self._redraw()

    def sync_x(self, first, *_) -> None:
        """由 Treeview xscrollcommand 呼叫，同步水平捲動位置。"""
        self.xview_moveto(first)

    def _redraw(self) -> None:
        self.delete("all")
        dates   = self._dates
        total_w = self.EMP_W + self.HOURS_W + self.COL_W * len(dates)
        self.configure(scrollregion=(0, 0, total_w, self.TOTAL_H))

        # 員工欄（跨兩排高度）
        self.create_rectangle(0, 0, self.EMP_W, self.TOTAL_H,
                               fill=self._BG, outline=self._GRID)
        self.create_text(self.EMP_W // 2, self.TOTAL_H // 2,
                         text="員工",
                         font=("TkDefaultFont", 12, "bold"),
                         fill=self._INK, anchor="center")

        # 時數欄（跨兩排高度）
        hx = self.EMP_W
        self.create_rectangle(hx, 0, hx + self.HOURS_W, self.TOTAL_H,
                               fill=self._BG, outline=self._GRID)
        self.create_text(hx + self.HOURS_W // 2, self.TOTAL_H // 2,
                         text="時數",
                         font=("TkDefaultFont", 12, "bold"),
                         fill=self._INK, anchor="center")

        for i, d in enumerate(dates):
            x0 = self.EMP_W + self.HOURS_W + i * self.COL_W
            x1 = x0 + self.COL_W
            is_wkend = d.weekday() >= 5   # 5=Sat, 6=Sun
            bg = self._WKBG if is_wkend else self._BG
            fg = self._RED  if is_wkend else self._INK

            # 上列：月/日
            self.create_rectangle(x0, 0, x1, self.DATE_H,
                                   fill=bg, outline=self._GRID)
            self.create_text((x0 + x1) // 2, self.DATE_H // 2,
                             text=f"{d.month}/{d.day}",
                             font=("TkDefaultFont", 11, "bold"),
                             fill=fg, anchor="center")

            # 下列：星期
            self.create_rectangle(x0, self.DATE_H, x1, self.TOTAL_H,
                                   fill=bg, outline=self._GRID)
            self.create_text((x0 + x1) // 2,
                             self.DATE_H + self.WD_H // 2,
                             text=_WEEK[d.weekday()],
                             font=("TkDefaultFont", 11),
                             fill=fg, anchor="center")


class _ScheduleGrid(tk.Frame):
    """Canvas-based schedule grid；kind=3（排休）顯示紅字。"""
    ROW_H   = 26
    EMP_W   = 150
    HOURS_W = 60
    COL_W   = 52
    FONT    = ("TkDefaultFont", 11)
    BG      = ("#FFFFFF", "#F8F9FD")
    GRID_C  = "#E8EBF4"
    INK     = "#2D3048"
    REST_FG = "#CC2200"   # kind=3 排休紅字

    def __init__(self, parent, on_double=None, **kw):
        super().__init__(parent, **kw)
        self._on_double = on_double          # callback(emp_obj, col_idx)
        self._rows: list = []                # [(emp, values_list, kinds_list)]
        self._dates: list = []

        self._cv = tk.Canvas(self, bg="#FFFFFF", highlightthickness=0)
        # 設定捲動單位=列高,讓滾輪/觸控板一格捲一列(預設 0 會導致幾乎不動)
        self._cv.configure(yscrollincrement=self.ROW_H)
        self._cv.grid(row=0, column=0, sticky="nsew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        _bind_mousewheel(self._cv)
        self._cv.bind("<Double-Button-1>", self._handle_double)

    # ── 對外 scroll API（接替 ttk.Treeview） ──────────────────────────────
    def configure(self, **kw):
        ys = kw.pop("yscrollcommand", None)
        xs = kw.pop("xscrollcommand", None)
        if ys: self._cv.configure(yscrollcommand=ys)
        if xs: self._cv.configure(xscrollcommand=xs)
        if kw: super().configure(**kw)

    def yview(self, *a): self._cv.yview(*a)
    def xview(self, *a): self._cv.xview(*a)
    def xview_moveto(self, f): self._cv.xview_moveto(f)

    # ── 填充資料 ──────────────────────────────────────────────────────────
    def populate(self, dates: list, emp_rows: list) -> None:
        """
        dates    : list[dt.date]
        emp_rows : [(emp_obj, values_list, kinds_list)]
          values_list[0] = 員工標籤, [1] = 時數, [2+] = 班別代號
          kinds_list[j]  = 第 j 個日期的 pb29004 (int|None)
        """
        self._cv.delete("all")
        self._rows  = emp_rows
        self._dates = dates
        total_w = self.EMP_W + self.HOURS_W + self.COL_W * len(dates)

        for i, (emp, vals, kinds) in enumerate(emp_rows):
            y0 = i * self.ROW_H
            y1 = y0 + self.ROW_H
            bg = self.BG[i % 2]

            # 員工欄（靠左文字）
            x = 0
            self._cv.create_rectangle(x, y0, x + self.EMP_W, y1,
                                      fill=bg, outline=self.GRID_C)
            self._cv.create_text(x + 8, (y0 + y1) // 2, text=vals[0],
                                 font=self.FONT, fill=self.INK,
                                 anchor="w", width=self.EMP_W - 12)
            x += self.EMP_W

            # 時數欄（置中）
            self._cv.create_rectangle(x, y0, x + self.HOURS_W, y1,
                                      fill=bg, outline=self.GRID_C)
            self._cv.create_text(x + self.HOURS_W // 2, (y0 + y1) // 2,
                                 text=vals[1], font=self.FONT, fill=self.INK,
                                 anchor="center")
            x += self.HOURS_W

            # 日期欄
            for j in range(len(dates)):
                code = vals[2 + j] if 2 + j < len(vals) else ""
                kind = kinds[j] if j < len(kinds) else None
                fg   = self.REST_FG if self._is_rest_cell(kind, code) else self.INK
                self._cv.create_rectangle(x, y0, x + self.COL_W, y1,
                                          fill=bg, outline=self.GRID_C)
                if code:
                    self._cv.create_text(x + self.COL_W // 2, (y0 + y1) // 2,
                                         text=code, font=self.FONT, fill=fg,
                                         anchor="center", width=self.COL_W - 4)
                x += self.COL_W

        total_h = len(emp_rows) * self.ROW_H
        self._cv.configure(scrollregion=(0, 0, total_w, total_h))

    @staticmethod
    def _is_rest_cell(kind, code: str) -> bool:
        """kind=3 → 排休紅字；kind=None 時以 991* 代碼判斷。"""
        if kind == 3:
            return True
        if kind is None and code and code.startswith("991"):
            return True
        return False

    def _handle_double(self, event) -> None:
        if not self._on_double or not self._rows:
            return
        cx = self._cv.canvasx(event.x)
        cy = self._cv.canvasy(event.y)
        if cx < self.EMP_W + self.HOURS_W:
            return   # 點到員工欄或時數欄
        col_idx = int((cx - self.EMP_W - self.HOURS_W) // self.COL_W)
        row_idx = int(cy // self.ROW_H)
        if row_idx < 0 or row_idx >= len(self._rows):
            return
        if col_idx < 0 or col_idx >= len(self._dates):
            return
        self._on_double(self._rows[row_idx][0], col_idx)


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
        # 設定捲動單位=列高,讓滾輪/觸控板一格捲一列(預設 0 會導致幾乎不動)
        self._cv.configure(yscrollincrement=self.ROW_H)
        vsb = ttk.Scrollbar(self, orient="vertical",   command=self._cv.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self._cv.xview)
        self._cv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._cv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        _bind_mousewheel(self._cv)
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


class _SuggestionTable(tk.Frame):
    """High-density suggestion list with row selection and quick apply."""

    COLS = [
        ("status", "狀態", 58, "center"),
        ("type", "類型", 70, "center"),
        ("date", "日期", 92, "center"),
        ("emp", "員工", 120, "w"),
        ("current", "原班別", 72, "center"),
        ("scheduled", "表定", 92, "center"),
        ("actual", "實際", 92, "center"),
        ("delta", "差異", 92, "center"),
        ("best", "最佳建議", 130, "w"),
        ("hours", "工時", 52, "center"),
        ("action", "套用", 74, "center"),
    ]

    def __init__(self, parent, on_select=None, on_apply=None, **kw):
        super().__init__(parent, **kw)
        self._rows: dict[str, dict] = {}
        self._on_select = on_select
        self._on_apply = on_apply
        cols = [c[0] for c in self.COLS]
        self.tv = ttk.Treeview(
            self, style="App.Treeview", columns=cols, show="headings",
            selectmode="browse", height=18
        )
        for key, label, width, anchor in self.COLS:
            self.tv.heading(key, text=label)
            self.tv.column(key, width=width, anchor=anchor, stretch=False)
        ysb = ttk.Scrollbar(self, orient="vertical", command=self.tv.yview)
        xsb = ttk.Scrollbar(self, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tv.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        _bind_mousewheel(self.tv)
        self.tv.tag_configure("applied", foreground=_C["green"])
        self.tv.tag_configure("failed", foreground=_C["red"])
        self.tv.tag_configure("empty", foreground=_C["muted"])
        self.tv.bind("<<TreeviewSelect>>", self._handle_select)
        self.tv.bind("<ButtonRelease-1>", self._handle_click)

    def populate(self, rows: list[dict]) -> None:
        self._rows = {r["id"]: r for r in rows}
        self.tv.delete(*self.tv.get_children())
        for row in rows:
            tags = []
            if row.get("status") == "已套用":
                tags.append("applied")
            elif row.get("status") == "失敗":
                tags.append("failed")
            elif not row.get("best_candidate"):
                tags.append("empty")
            self.tv.insert("", "end", iid=row["id"], values=self._values(row),
                           tags=tuple(tags))

    def select(self, row_id: str | None) -> None:
        if row_id and row_id in self._rows:
            self.tv.selection_set(row_id)
            self.tv.focus(row_id)
            self.tv.see(row_id)
            if self._on_select:
                self._on_select(self._rows[row_id])

    def first_id(self) -> str | None:
        children = self.tv.get_children()
        return children[0] if children else None

    def _values(self, row: dict) -> tuple:
        best = row.get("best_candidate")
        best_txt = "無合適建議"
        hours = ""
        action = ""
        if best:
            code, si, so, _d_in, _d_out, h = best
            name = row.get("shift_names", {}).get(code, "")
            best_txt = f"{code} {name or f'{si}-{so}'}".strip()
            hours = f"{h:g}h"
            action = "" if row.get("status") in ("已套用", "套用中") else "套用最佳"
        return (
            row.get("status", "待處理"),
            row.get("type", ""),
            row.get("date", ""),
            f"{row.get('emp_id', '')} {row.get('name', '')}".strip(),
            row.get("current_shift", ""),
            row.get("scheduled", ""),
            row.get("actual", ""),
            row.get("delta", ""),
            best_txt,
            hours,
            action,
        )

    def _handle_select(self, _event=None) -> None:
        sel = self.tv.selection()
        if sel and self._on_select:
            self._on_select(self._rows.get(sel[0]))

    def _handle_click(self, event) -> None:
        iid = self.tv.identify_row(event.y)
        col = self.tv.identify_column(event.x)
        if not iid or col != f"#{len(self.COLS)}":
            return
        row = self._rows.get(iid)
        if (row and row.get("best_candidate")
                and row.get("status") not in ("已套用", "套用中")):
            if self._on_apply:
                self._on_apply(iid)
