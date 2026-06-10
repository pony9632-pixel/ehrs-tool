"""純函式單元測試 —— 打卡/排班邏輯(不需登入、不開 GUI)。

執行:
    python3 test_punch_logic.py        # 內建 runner
    python3 -m pytest test_punch_logic.py -q

涵蓋 app.py 內的純函式:_punch_remark / _split_punches /
_split_punches_jiezhuan / _assign_single_punch / _calc_hours /
_suggest_shifts / _get_schedule_periods / _first_field。
匯入 app 僅載入模組層級程式(customtkinter 主題設定),不建立 Tk 視窗。
"""
import datetime as dt

import app


# 給 _suggest_shifts 用的最小假班別表:code -> (時間字串, kind, 時數)
FAKE_CODES = {
    "A001": ("09:00-18:00", 1, 8.0),
    "B001": ("12:00-20:00", 1, 7.0),
    "C001": ("10:00-20:00", 1, 9.0),
    "991":  ("排休", 3, 0.0),          # 排休應被略過
    "N001": ("22:00-07:00", 1, 8.0),   # 跨夜班
}


# --------------------------------------------------------------------------- #
# _punch_remark
# --------------------------------------------------------------------------- #
def test_remark_normal_on_time():
    assert app._punch_remark("09:00", "18:00", "09:00", "18:00") == ""


def test_remark_late():
    assert app._punch_remark("09:05", "18:00", "09:00", "18:00") == "遲到"


def test_remark_early_leave():
    assert app._punch_remark("09:00", "17:30", "09:00", "18:00") == "早退"


def test_remark_forgot_in():
    assert app._punch_remark("", "18:00", "09:00", "18:00") == "上班忘打卡"


def test_remark_forgot_out():
    assert app._punch_remark("09:00", "", "09:00", "18:00") == "下班忘打卡"


def test_remark_forgot_both():
    assert app._punch_remark("", "", "09:00", "18:00") == "上班忘打卡 下班忘打卡"


def test_remark_overtime_threshold():
    # 超過表定 30 分內不算加班
    assert app._punch_remark("09:00", "18:25", "09:00", "18:00") == ""
    # 超過 30 分 → 加班(四捨五入到 0.5h)
    assert app._punch_remark("09:00", "19:00", "09:00", "18:00") == "加班1h"
    assert app._punch_remark("09:00", "18:40", "09:00", "18:00") == "加班0.5h"


def test_remark_late_and_overtime_combined():
    r = app._punch_remark("09:10", "19:00", "09:00", "18:00")
    assert "遲到" in r and "加班1h" in r


def test_remark_restday_kind2():
    assert app._punch_remark("10:00", "16:00", "10:00", "18:00", kind=2) == "休息日加班6h"
    # 只有上班卡:仍標休息日加班(無時數)
    assert app._punch_remark("10:00", "", "10:00", "18:00", kind=2) == "休息日加班"


def test_remark_overnight_no_false_early():
    # 跨夜班 22:00-07:00,準時上下班不應誤判早退
    assert app._punch_remark("22:00", "07:00", "22:00", "07:00") == ""


# --------------------------------------------------------------------------- #
# _overtime_hours / _classify_remark
# --------------------------------------------------------------------------- #
def test_overtime_hours_weekday():
    assert app._overtime_hours("09:00", "19:00", "09:00", "18:00") == 1.0
    assert app._overtime_hours("09:00", "18:40", "09:00", "18:00") == 0.5
    assert app._overtime_hours("09:00", "18:20", "09:00", "18:00") == 0.0   # 30 分內
    assert app._overtime_hours("09:00", "17:30", "09:00", "18:00") == 0.0   # 早退非加班


def test_overtime_hours_restday():
    assert app._overtime_hours("10:00", "16:00", "10:00", "18:00", kind=2) == 6.0
    assert app._overtime_hours("10:00", "", "10:00", "18:00", kind=2) == 0.0


def test_overtime_hours_matches_remark_string():
    # _punch_remark 的字串時數應等於 _overtime_hours(單一來源,不再 regex 反解)
    import re as _re
    for ci, co, si, so, kind in [
        ("09:00", "19:00", "09:00", "18:00", None),
        ("09:00", "18:40", "09:00", "18:00", None),
        ("10:00", "16:00", "10:00", "18:00", 2),
    ]:
        remark = app._punch_remark(ci, co, si, so, kind=kind)
        m = _re.search(r"加班([\d.]+)h", remark)
        parsed = float(m.group(1)) if m else 0.0
        assert parsed == app._overtime_hours(ci, co, si, so, kind=kind)


def test_classify_remark_priority():
    assert app._classify_remark("") == ""
    assert app._classify_remark("遲到") == "err"
    assert app._classify_remark("曠職") == "err"
    assert app._classify_remark("加班1h") == "ot"
    assert app._classify_remark("休息日加班6h") == "ot"
    assert app._classify_remark("特休") == "leave"
    # err 優先於 ot:遲到 + 加班 → err
    assert app._classify_remark("遲到 加班1h") == "err"


# --------------------------------------------------------------------------- #
# _split_punches / _split_punches_jiezhuan
# --------------------------------------------------------------------------- #
def test_split_punches_basic():
    assert app._split_punches(["08:55", "12:00", "18:05"], "09:00", "18:00") == ("08:55", "18:05")


def test_split_punches_no_sched():
    assert app._split_punches(["08:55", "18:05"], "", "") == ("08:55", "18:05")


def test_split_jiezhuan_overnight():
    # 跨夜班:凌晨下班卡應歸到下班群
    ci, co = app._split_punches_jiezhuan(["21:55", "06:10"], "22:00", "07:00")
    assert ci == "21:55" and co == "06:10"


# --------------------------------------------------------------------------- #
# _assign_single_punch
# --------------------------------------------------------------------------- #
def test_assign_single_morning_is_clockin():
    assert app._assign_single_punch("09:02", "09:00", "18:00") == ("09:02", "")


def test_assign_single_evening_is_clockout():
    assert app._assign_single_punch("17:55", "09:00", "18:00") == ("", "17:55")


# --------------------------------------------------------------------------- #
# _calc_hours / _deduct_break
# --------------------------------------------------------------------------- #
def test_calc_hours_deductions():
    assert app._calc_hours("09:00", "13:00") == 4.0       # ≤4h 不扣
    assert app._calc_hours("09:00", "14:00") == 4.5       # >4h <9h 扣 0.5
    assert app._calc_hours("09:00", "18:00") == 8.0       # 9h - 1h
    assert app._calc_hours("09:00", "19:00") == 9.0       # 10h - 1h


def test_deduct_break_boundaries():
    # 回歸:出勤統計過去在剛好 9h 時用「> 9」只扣 0.5h,與 _calc_hours 不一致
    assert app._deduct_break(4.0) == 4.0      # ≤4h 不扣
    assert app._deduct_break(4.5) == 4.0      # >4h 扣 0.5
    assert app._deduct_break(8.99) == 8.49
    assert app._deduct_break(9.0) == 8.0      # 剛好 9h 也要扣 1h
    assert app._deduct_break(10.0) == 9.0


# --------------------------------------------------------------------------- #
# _suggest_shifts
# --------------------------------------------------------------------------- #
def test_suggest_picks_matching_hours():
    res = app._suggest_shifts("09:00", "18:00", FAKE_CODES, preferred_hours=8.0)
    assert res, "should return candidates"
    assert res[0][0] == "A001"   # 09:00-18:00 = 8h,最貼合


def test_suggest_skips_rest_code():
    res = app._suggest_shifts("09:00", "18:00", FAKE_CODES, preferred_hours=8.0)
    assert all(code != "991" for code, *_ in res)


def test_suggest_no_punch_returns_empty():
    assert app._suggest_shifts("", "", FAKE_CODES) == []


# --------------------------------------------------------------------------- #
# _get_schedule_periods
# --------------------------------------------------------------------------- #
def test_periods_count_and_anchor():
    periods = app._get_schedule_periods(115)
    assert len(periods) == 14
    assert periods[0] == (dt.date(2025, 12, 22), dt.date(2026, 1, 18))
    # 每期 28 天(start..end 含端點 27 天差)
    for s, e in periods:
        assert (e - s).days == 27
        assert s.weekday() == 0   # 週一起始


def test_periods_adjacent_year_shift():
    # ±52 週(364 天)推算
    p115 = app._get_schedule_periods(115)[0][0]
    p116 = app._get_schedule_periods(116)[0][0]
    assert (p116 - p115).days == 364


# --------------------------------------------------------------------------- #
# _first_field
# --------------------------------------------------------------------------- #
def test_first_field_precedence_and_blank():
    row = {"員工代號": "  ", "pa51002": "E123"}
    assert app._first_field(row, ("員工代號", "pa51002")) == "E123"  # 跳過空白
    assert app._first_field({}, ("x", "y")) == ""


# --------------------------------------------------------------------------- #
# _wheel_units（滾輪/觸控板捲動，跨平台）
# --------------------------------------------------------------------------- #
class _FakeWheelEvent:
    def __init__(self, delta=0, num=0):
        self.delta = delta
        self.num = num


def test_wheel_macos_small_delta_nonzero():
    # 回歸:macOS 觸控板小 delta 過去被 int(-delta/60) 變 0 → 不捲動
    assert app._wheel_units(_FakeWheelEvent(delta=1)) == -1
    assert app._wheel_units(_FakeWheelEvent(delta=3)) == -3
    assert app._wheel_units(_FakeWheelEvent(delta=-2)) == 2


def test_wheel_windows_120_multiples():
    assert app._wheel_units(_FakeWheelEvent(delta=120)) == -1
    assert app._wheel_units(_FakeWheelEvent(delta=240)) == -2
    assert app._wheel_units(_FakeWheelEvent(delta=-120)) == 1


def test_wheel_linux_buttons():
    assert app._wheel_units(_FakeWheelEvent(num=4)) == -1
    assert app._wheel_units(_FakeWheelEvent(num=5)) == 1


def test_wheel_noop():
    assert app._wheel_units(_FakeWheelEvent()) == 0


if __name__ == "__main__":
    import sys
    funcs = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for fn in funcs:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"FAIL {fn.__name__}: {e}")
        except Exception as e:  # noqa
            failed += 1
            print(f"ERROR {fn.__name__}: {type(e).__name__}: {e}")
    print(f"\n{len(funcs) - failed}/{len(funcs)} passed")
    sys.exit(1 if failed else 0)
