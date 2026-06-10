"""打卡 / 排班純邏輯（不依賴 Tk，可獨立單元測試）。

從 app.py 抽出：備注產生、打卡分群、工時計算、班別建議、
四週期間推算與欄位取值工具。test_punch_logic.py 直接測這裡的函式。
"""
import datetime as dt
import re
from functools import lru_cache

from shift_codes_ref import SHIFT_CODES_REF

# 各 API 回傳列的欄位 key 候選（readable 中文欄名在前、原始代碼在後）
_DEPT_KEYS = ("部門名稱", "部門", "部門全名", "pa51014FullName", "pa51014Name")
_DEPT_CODE_KEYS = ("部門代號", "pa51014")
_DEPT_FULL_KEYS = ("部門全名", "pa51014FullName")
_DEPT_NAME_KEYS = ("部門名稱", "部門", "pa51014Name")
_EMP_ID_KEYS = ("員工代號", "pa51002")
_EMP_NAME_KEYS = ("員工姓名", "pa51004")


def _first_field(row: dict, keys: tuple[str, ...]) -> str:
    for key in keys:
        value = row.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


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


@lru_cache(maxsize=None)
def _shift_times(code: str) -> tuple[str, str]:
    """從 SHIFT_CODES_REF 解析班別的表定上班/下班時間（HH:MM）。
    SHIFT_CODES_REF 為唯讀常數,輸入班別代碼有限,故以 lru_cache 快取(熱路徑)。"""
    name = SHIFT_CODES_REF.get(code, ("", 3, 0.0))[0]
    m = re.search(r"(\d{2}:\d{2})-(\d{2}:\d{2})", name)
    return (m.group(1), m.group(2)) if m else ("", "")


@lru_cache(maxsize=None)
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


def _overtime_hours(clock_in: str, clock_out: str, sched_in: str, sched_out: str,
                    kind: int | None = None) -> float:
    """加班 / 休息日加班時數(四捨五入至 0.5h);無加班回 0.0。
    為「備注字串」與「出勤統計」共用的單一計算來源(避免從字串反解)。"""
    # 休息日(kind=2):工作全程即加班
    if kind == 2:
        if not (clock_in and clock_out and sched_in and sched_out):
            return 0.0
        ci = _punch_mins(clock_in)
        co = _punch_mins(clock_out)
        if co < ci:
            co += 1440   # 跨夜
        worked = co - ci
        return round(worked / 60 * 2) / 2 if worked > 0 else 0.0
    # 一般工作日:下班超過表定 30 分才算加班(以打卡下班為準)
    if not (clock_out and sched_in and sched_out):
        return 0.0
    si = _punch_mins(sched_in)
    so = _punch_mins(sched_out)
    co = _punch_mins(clock_out)
    if so < si:
        so += 1440
    if co < si - 120:
        co += 1440
    if co > so + 30:
        return round((co - so) / 60 * 2) / 2
    return 0.0


def _punch_remark(clock_in: str, clock_out: str, sched_in: str, sched_out: str,
                  kind: int | None = None) -> str:
    """產生備注文字：遲到、早退、上班忘打卡、下班忘打卡、加班（可複合）。
    kind=2 休息日：有打卡即算「休息日加班 Xh」，不另做遲到/早退判斷。
    """
    remarks: list[str] = []

    # ── 休息日（kind=2）：只要有打卡就是加班 ──────────────────────
    if kind == 2:
        oh = _overtime_hours(clock_in, clock_out, sched_in, sched_out, kind=2)
        if oh > 0:
            remarks.append(f"休息日加班{oh:g}h")
        elif clock_in:
            remarks.append("休息日加班")
        return " ".join(remarks)

    # ── 一般工作日 ─────────────────────────────────────────────────
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
        elif co > so + 30:                           # 超過 30 分鐘算加班
            oh = _overtime_hours(clock_in, clock_out, sched_in, sched_out, kind=kind)
            remarks.append(f"加班{oh:g}h")
    return " ".join(remarks)


# 備注分類所用關鍵字(集中一處,供 _classify_remark 與各分頁共用)
_REMARK_ERR_KEYWORDS = ("遲到", "上班忘打卡", "早退", "下班忘打卡", "曠職")
_REMARK_OT_KEYWORD = "加班"   # 含「加班」「休息日加班」


def _classify_remark(remark: str) -> str:
    """把備注歸成單一類別:'err'(異常) / 'ot'(加班) / 'leave'(假別) / ''(正常)。
    優先序 err > ot > leave。打卡分頁著色、篩選、統計、建議共用此判斷。"""
    remark = remark or ""
    if not remark:
        return ""
    if any(k in remark for k in _REMARK_ERR_KEYWORDS):
        return "err"
    if _REMARK_OT_KEYWORD in remark:
        return "ot"
    return "leave"


def _deduct_break(total: float) -> float:
    """依工作總時數扣除休息：≤ 4h 不扣；> 4h 且 < 9h 扣 0.5h；≥ 9h 扣 1h。
    排班工時、出勤統計、加班建議共用的唯一扣除規則。"""
    if total <= 4.0:
        return total
    return total - (1.0 if total >= 9.0 else 0.5)


def _calc_hours(sched_in: str, sched_out: str) -> float:
    """計算實際工時（扣除休息）。"""
    si = _punch_mins(sched_in)
    so = _punch_mins(sched_out)
    if so < si:
        so += 1440
    return _deduct_break((so - si) / 60.0)


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
