"""文中資訊 eHRS(線上人資)API client.

對應站台: http://ehrs.studioa.com.tw/Ehrsnet  (ASP.NET MVC, 版本 V26.x)

此模組以 requests.Session 維持登入後的 Forms-Authentication cookie
(.ASPXAUTHEHRSNET2),封裝三大功能:

  1. 登入                login()
  2. 排班讀取            get_schedule() / get_schedule_raw()
  3. 刷卡(打卡)讀取      get_punch_records()
  4. 排班建立/修改/刪除  set_shift() / delete_shift()   ← 寫入,預設 dry-run

文中系統的欄位代碼說明(reverse-engineered):
  員工 (PA51 員工主檔)
    pa51002  員工代號        pa51004  員工姓名
    pa51014  部門代號        pa51014Name 部門名稱
    pa51135  職稱代號        pa51135Name 職稱名稱
    pa51020  預設班別代號    pa51011  員工現況(1=在職,2=離職)
    pa51024  到職日期(ISO,如 2017-06-01T00:00:00;get_schedule 的 emp 物件即帶)
    pa51018  刷卡類別(1=免刷卡,2=刷卡人員)
  排班 (PB29/WPB29 排班檔)
    pb29001  公司別          pb29002  排班日期
    pb29003  員工代號        pb29004  班別種類(數字)
    pb29005  班別代號        pb29005Name 班別名稱
    pb29006  備註            pb29008  排班檢核警告訊息
    pb29995  記錄主鍵(PK)    -> 有值=既有班(用 Update);無=新增(用 Create)
  刷卡 (PA65 刷卡明細 / PB12 出勤)
    pb12003  出勤日期
    pa65003  刷卡時間        pa65002  卡鐘代號
    pa65006  識別卡號        pa65008Name 來源
    pa65007Name 結轉註記
"""
from __future__ import annotations

import datetime as _dt
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Any, Iterable, Optional, Union

import requests

DEFAULT_BASE_URL = "http://ehrs.studioa.com.tw/Ehrsnet"
AUTH_COOKIE = ".ASPXAUTHEHRSNET2"

DateLike = Union[str, _dt.date, _dt.datetime]


class EhrsError(Exception):
    """eHRS API 操作失敗。"""


# --------------------------------------------------------------------------- #
# 日期工具
# --------------------------------------------------------------------------- #
def _as_date(value: DateLike) -> _dt.date:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, str):
        return _dt.date.fromisoformat(value[:10])
    raise TypeError(f"無法解析日期: {value!r}")


def _iso_ms(value: DateLike, end_of_day: bool = False) -> str:
    """轉成文中 API 用的時間字串,例如 2026-06-01T00:00:00.000。"""
    d = _as_date(value)
    t = "23:59:59.000" if end_of_day else "00:00:00.000"
    return f"{d.isoformat()}T{t}"


# 四週排班期間:民國115年第1期起始 2025-12-22(週一),每期 28 天連續推算。
_PERIOD_ANCHOR = _dt.date(2025, 12, 22)
_PERIOD_DAYS = 28


def _period_start_of(value: DateLike) -> _dt.date:
    """回傳包含該日期的四週期間之起始日(週一)。
    例:2026-07-01 → 2026-06-08(第7期 6/8～7/5 的起點)。"""
    d = _as_date(value)
    idx = (d - _PERIOD_ANCHOR).days // _PERIOD_DAYS
    return _PERIOD_ANCHOR + _dt.timedelta(days=idx * _PERIOD_DAYS)


# 寫入時伺服器會逐一回傳這些「跳過/確認某項驗證」的旗標(每回合一個,
# 一格要來回近 10 趟)。confirm=True 本就是「全部按確定」,故第一趟就把
# 整串旗標 + selected 一次帶上,讓伺服器一次過關 → 把 ~10 趟壓成 1 趟。
# 這只改 round-trip 次數,不改寫入結果(原本逐回合也是全部確認)。
# 清單外若出現新旗標,_post_shift_write 的迴圈仍會逐回合補上(後援)。
_WRITE_CONFIRM_FLAGS = (
    "isCheckOverTime",
    "isCheckCanNotWorkingTime",
    "isCheckScheduleIntervals",
    "isCheckAllowConsecutive",
    "isCheckTotalWorkingHoursOfWeek",
    "isCheckPeriodicOverTime",
    "isCheckPeriodicHolidays",
    "isCheckPeriodic2WeekHolidays",
    "isCheckPeriodic1WeekHolidays",
)


# --------------------------------------------------------------------------- #
# 排班資料結構
# --------------------------------------------------------------------------- #
@dataclass
class Shift:
    """單一格排班(某員工某天)。"""

    date: str            # YYYY-MM-DD
    code: str            # pb29005 班別代號 (例 'B005'、'991')
    name: str            # pb29005Name 班別名稱 (例 'HQ0900-1800')
    kind: Any            # pb29004 班別種類
    pk: Optional[int]    # pb29995 記錄主鍵,None 表示尚未建檔
    raw: dict = field(repr=False, default_factory=dict)


@dataclass
class EmployeeSchedule:
    """一位員工整月的排班。"""

    emp_id: str          # pa51002
    name: str            # pa51004
    title: str           # pa51135Name 職稱
    working_minutes: int
    rest_days: int
    shifts: list[Shift]
    raw: dict = field(repr=False, default_factory=dict)
    hire_date: str = ""  # pa51024 到職日期 (ISO 字串; 可能為空)

    def shift_on(self, date: DateLike) -> Optional[Shift]:
        target = _as_date(date).isoformat()
        return next((s for s in self.shifts if s.date == target), None)


# --------------------------------------------------------------------------- #
# Client
# --------------------------------------------------------------------------- #
class EhrsClient:
    def __init__(
        self,
        base_url: str = DEFAULT_BASE_URL,
        timeout: int = 30,
        user_agent: str = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    ) -> None:
        self.base = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": user_agent,
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/json, text/plain, */*",
            }
        )
        self._logged_in = False

    # ----- 低階 ----- #
    def _url(self, path: str) -> str:
        return f"{self.base}/{path.lstrip('/')}"

    def _post_json(self, path: str, payload: dict) -> Any:
        if not self._logged_in:
            raise EhrsError("尚未登入,請先呼叫 login()")
        resp = self.session.post(
            self._url(path), json=payload, timeout=self.timeout
        )
        if resp.status_code != 200:
            raise EhrsError(f"{path} 回應 HTTP {resp.status_code}")
        try:
            return resp.json()
        except ValueError as exc:  # 多半是 session 過期被導回登入頁
            raise EhrsError(
                f"{path} 回應非 JSON(session 可能已過期):{resp.text[:200]}"
            ) from exc

    # ----- 登入 ----- #
    def login(
        self, account: str, password: str, language: str = "ZH-TW"
    ) -> "EhrsClient":
        """以帳密登入並建立 session。成功後同一 client 即可呼叫其餘方法。"""
        host = requests.utils.urlparse(self.base).hostname or ""
        # 先取得 ASP.NET_SessionId
        self.session.get(self._url("Account/Index"), timeout=self.timeout)
        resp = self.session.post(
            self._url("Account/Create"),
            data={
                "account": account,
                "password": password,
                "currentLanguage": language,
                "returnUrl": "",
                "localHostUrl": host,
            },
            allow_redirects=False,
            timeout=self.timeout,
        )
        ok = resp.status_code in (301, 302) and AUTH_COOKIE in self.session.cookies
        if not ok:
            raise EhrsError(
                "登入失敗,請確認帳號密碼。"
                f"(HTTP {resp.status_code}, 已取得 cookie: "
                f"{list(self.session.cookies.keys())})"
            )
        self._logged_in = True
        return self

    # ----- 排班讀取 ----- #
    def _schedule_filter(
        self, year: int, month: int, **overrides: Any
    ) -> dict:
        """組出 Webm1031/Fetch 的 filter。查詢與寫入共用,
        寫入時這整包 filter 必須一起送(否則伺服器只處理不寫檔)。"""
        first = _dt.date(year, month, 1)
        last = (
            _dt.date(year + (month == 12), (month % 12) + 1, 1)
            - _dt.timedelta(days=1)
        )
        payload = {
            "type": "shift",
            "displayMode": 0,
            "year": year,
            "month": month,
            "flexType": 2,
            "shiftType": 0,
            "calendarType": 0,
            "startDayType": 0,
            "start": _iso_ms(first),
            "end": _iso_ms(last, end_of_day=True),
            "flexTimeCycles": [],
        }
        payload.update(overrides)
        return payload

    def get_schedule_raw(
        self, year: int, month: int, **overrides: Any
    ) -> dict:
        """回傳 Webm1031/Fetch 的原始 calendar 物件(完整、未整理)。"""
        payload = self._schedule_filter(year, month, **overrides)
        data = self._post_json("Webm/Webm1031/Fetch", payload)
        return data["data"]["calendar"]

    def get_schedule(self, year: int, month: int) -> list[EmployeeSchedule]:
        """回傳整理後的整月排班(每位員工 + 每日班別)。"""
        calendar = self.get_schedule_raw(year, month)
        result: list[EmployeeSchedule] = []
        for emp in calendar.get("shiftEmployees", []):
            shifts: list[Shift] = []
            for cell in emp.get("cells", []):
                for sch in cell.get("schedules", []):
                    code = sch.get("pb29005") or ""
                    if not code and sch.get("pb29995") is None:
                        continue  # 空白格
                    shifts.append(
                        Shift(
                            date=_as_date(cell["calendarDate"]).isoformat(),
                            code=code,
                            name=sch.get("pb29005Name") or "",
                            kind=sch.get("pb29004"),
                            pk=sch.get("pb29995"),
                            raw=sch,
                        )
                    )
            result.append(
                EmployeeSchedule(
                    emp_id=emp.get("pa51002", ""),
                    name=emp.get("pa51004", ""),
                    title=emp.get("pa51135Name", ""),
                    working_minutes=emp.get("workingMinutes", 0),
                    rest_days=emp.get("restDays", 0),
                    shifts=shifts,
                    raw=emp,
                    hire_date=emp.get("pa51024", "") or "",
                )
            )
        return result

    # ----- 請假記錄讀取 ----- #
    def get_leave_records(
        self,
        start: DateLike,
        end: DateLike,
        emp_status: Iterable[str] = ("1", "2"),
    ) -> list[dict]:
        """查詢請假記錄 (Webr3070)。

        每筆記錄關鍵欄位：
          pa60002      員工代號
          pa60002Name  員工姓名
          pa60004Name  假別（特休/病假/事假…）
          pa60006      出勤日期  YYYY-MM-DDTHH:MM:SS
          pa60007      請假開始
          pa60008      請假結束
          pa60011      時數（分鐘）
        """
        s = _as_date(start)
        e = _as_date(end)
        payload: dict[str, Any] = {
            "filterMethod": 1,
            "hourFilter": "",
            "isDesignPreview": False,
            "leaveName": 1,
            "pa51011s": list(emp_status),
            "pa60006End": _iso_ms(e, end_of_day=True),
            "pa60006Start": _iso_ms(s),
            "pa6002021End": _iso_ms(e, end_of_day=True),
            "pa6002021Start": _iso_ms(s),
            "periodFilter": 3,
            "periodSettingEnd": _iso_ms(e, end_of_day=True),
            "periodSettingStart": _iso_ms(s),
            "queryMode": 1,
            "queryModeName": "依請假日期區間",
            "reportType": "WEBR3070RA",
            "reportType2": "WEBR3070RA",
            "templateId": "56dab7dd-1956-46cf-91ef-86b8ab0feafe",
        }
        data = self._post_json("Webr/Webr3070/BrowseResource", payload)["data"]
        return data.get("data", [])

    # ----- 刷卡(打卡)讀取 ----- #
    def get_punch_records(
        self,
        start: DateLike,
        end: DateLike,
        emp_start: Optional[str] = None,
        emp_end: Optional[str] = None,
        dept_start: Optional[str] = None,
        dept_end: Optional[str] = None,
        punch_types: Iterable[str] = ("2",),       # pa51018s:2=刷卡人員
        emp_status: Iterable[str] = ("1", "2"),    # pa51011s:1=在職,2=離職
        report_setting: int = 1,                   # 1=實際刷卡明細表
        readable: bool = False,
        **overrides: Any,
    ) -> list[dict]:
        """查詢刷卡(打卡)明細。

        readable=True 時會把 paXXXXX 欄位代碼換成中文欄名
        (依 API 回傳的 modelDef 對照表)。
        """
        payload: dict[str, Any] = {
            "pb12003Start": _iso_ms(start),
            "pb12003End": _iso_ms(end),
            "pa51018s": list(punch_types),
            "pa51011s": list(emp_status),
            "reportSetting": report_setting,
            "reportType": "WEBR3060R",
            "isDesignPreview": False,
        }
        if emp_start is not None:
            payload["pa51002Start"] = emp_start
        if emp_end is not None:
            payload["pa51002End"] = emp_end
        if dept_start is not None:
            payload["pa51014Start"] = dept_start
        if dept_end is not None:
            payload["pa51014End"] = dept_end
        payload.update(overrides)

        data = self._post_json("Webr/Webr3060/BrowseResource", payload)["data"]
        rows = data.get("data", [])
        if not readable:
            return rows
        cols = {
            f: c.get("name", f)
            for f, c in data.get("modelDef", {}).get("columns", {}).items()
        }
        return [
            {cols.get(k, k): v for k, v in row.items()} for row in rows
        ]

    # ----- 特休餘額讀取 ----- #
    #
    # 文中沒有「特休額度」報表(WEBR3100 補休 / WEBR3120 彈性假 / WEBR3121 法定假
    # 三張額度報表都不含特休,店長選單也沒有特休報表)。唯一拿得到「各員工特休
    # 剩餘 + 到期日」的地方,是『請假單簽呈』(WEBF1010)選假別=特休後,按抵用
    # 時數旁的「選擇」跳出的「特休假抵扣項目」清單。其端點:
    #     Common/Webf1010AskForLeave/ResourceDeductionList
    # 依 payload.wfl60.fl60002(員工代號)回傳「該員工」目前可用的各批特休額度,
    # 換掉 fl60002 即可查任一員工(實測店長帳號可查同部門其他人,非只限本人)。
    # 這是唯讀查詢,只是列出可抵扣額度,不會建立或送出任何請假單。
    #
    # 回傳結構:data.deductionData.deductions[] 每筆=一批特休額度,關鍵欄位
    #   deductionLeaveName 假別(特休)   year 年度          seniority 年資
    #   startDate 啟用日期               endDate 停用日期(=到期日)
    #   deductionMinutes 給假時數(分)    usedMinutes 已用    leftMinutes 剩餘
    # 時數單位為分鐘;天數 = 分鐘 / 每日時數(預設 480 = 8 小時)。fl60004=9 即特休。
    _ANNUAL_LEAVE_ENDPOINT = "Common/Webf1010AskForLeave/ResourceDeductionList"

    def get_annual_leave_balance(
        self,
        emp_id: str,
        dept: str,
        ref_date: Optional[DateLike] = None,
        company: Optional[str] = None,
        leave_type: int = 9,
        day_minutes: int = 480,
    ) -> list[dict]:
        """查詢單一員工的特休(特別休假)各批額度:剩餘時數/天數 + 到期日。

        ref_date 預設今天,決定「目前可用」的快照(已過停用日的批次不回傳)。
        company 預設由 emp_id 的英文字首推得(如 'SA1588'→'SA')。
        回傳 list,每批一個 dict:emp_id、year(年度)、seniority(年資)、
        start_date(啟用)、end_date(停用=到期)、granted/used/left_minutes、
        對應的 granted/used/left_days(= 分鐘 / day_minutes,四捨五入兩位)。
        """
        ref = _as_date(ref_date) if ref_date else _dt.date.today()
        if company:
            comp = company
        else:
            m = re.match(r"[A-Za-z]+", emp_id or "")
            comp = m.group(0) if m else ""
        day_iso = f"{ref.isoformat()}T00:00:00.000"
        wfl60 = {
            "pa51014": str(dept),
            "fl60001": comp,
            "fl60002": emp_id,
            "fl60004": leave_type,
            "fl60028Range": 1,
            "fl60005": day_iso,
            "fl60006": day_iso,
            "fl60007Date": day_iso,
            "fl60008Date": day_iso,
            "fl60007Time": f"{ref.isoformat()}T09:00:00.000",
            "fl60008Time": f"{ref.isoformat()}T18:00:00.000",
        }
        envelope = self._post_json(
            self._ANNUAL_LEAVE_ENDPOINT, {"wfl60": wfl60, "sameEmpWfl60s": []}
        )
        data = (envelope or {}).get("data") or {}
        deductions = (data.get("deductionData") or {}).get("deductions") or []
        out: list[dict] = []
        for row in deductions:
            gm = row.get("deductionMinutes") or 0
            um = row.get("usedMinutes") or 0
            lm = row.get("leftMinutes") or 0
            out.append(
                {
                    "emp_id": emp_id,
                    "leave_name": row.get("deductionLeaveName") or "特休",
                    "year": row.get("year"),
                    "seniority": row.get("seniority"),
                    "start_date": row.get("startDate") or "",
                    "end_date": row.get("endDate") or "",
                    "granted_minutes": gm,
                    "used_minutes": um,
                    "left_minutes": lm,
                    "granted_days": round(gm / day_minutes, 2),
                    "used_days": round(um / day_minutes, 2),
                    "left_days": round(lm / day_minutes, 2),
                    "raw": row,
                }
            )
        return out

    def get_annual_leave_balance_bulk(
        self,
        employees: list[dict],
        ref_date: Optional[DateLike] = None,
        leave_type: int = 9,
        day_minutes: int = 480,
        progress_cb=None,        # Callable[[done: int, total: int], None]
        max_workers: int = 6,
    ) -> list[dict]:
        """批次查詢多位員工特休餘額。employees 每筆需含 emp_id、dept,可選 name。

        回傳攤平的清單,且「每位員工至少一列」:有額度者每批一列;查無額度者一列
        (note='（無特休額度）');查詢失敗者一列(error=訊息)。順序同 employees 輸入,
        同一員工的多批依 API 回傳序。progress_cb(done, total) 每處理完一位呼叫一次。
        注意:eHRS 同一 session 多半被 ASP.NET session 鎖序列化,調高 workers 幫助有限。
        """
        results: list[Optional[list[dict]]] = [None] * len(employees)
        done = [0]
        lock = threading.Lock()

        def _blank(eid: str, name: str, **extra) -> dict:
            base = {
                "emp_id": eid, "name": name, "leave_name": "特休",
                "year": None, "seniority": None, "start_date": "", "end_date": "",
                "granted_minutes": 0, "used_minutes": 0, "left_minutes": 0,
                "granted_days": 0, "used_days": 0, "left_days": 0,
            }
            base.update(extra)
            return base

        def _tick() -> None:
            if progress_cb:
                with lock:
                    done[0] += 1
                    n = done[0]
                progress_cb(n, len(employees))

        def _one(idx: int, emp: dict) -> None:
            eid = emp.get("emp_id") or ""
            name = emp.get("name") or ""
            dept = emp.get("dept") or ""
            try:
                batches = self.get_annual_leave_balance(
                    eid, dept, ref_date=ref_date,
                    leave_type=leave_type, day_minutes=day_minutes,
                )
                if batches:
                    for b in batches:
                        b["name"] = name
                    results[idx] = batches
                else:
                    results[idx] = [_blank(eid, name, note="（無特休額度）")]
            except EhrsError as exc:
                results[idx] = [_blank(eid, name, error=str(exc))]
            _tick()

        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = [pool.submit(_one, i, e) for i, e in enumerate(employees)]
            for f in as_completed(futures):
                f.result()

        flat: list[dict] = []
        for sub in results:
            if sub:
                flat.extend(sub)
        return flat

    # ----- 排班寫入(建立 / 修改 / 刪除) ----- #
    #
    # ⚠️  此站台為正式環境(Studio A 人資)。寫入會動到真實班表。
    #     為安全起見,set_shift / delete_shift 預設 dry_run=True,只回傳
    #     將送出的 payload 供檢視,不會真的送出。確認無誤後再傳 dry_run=False。
    #
    #     寫入端點(由模組 JS 還原):
    #       Webm/Webm1031Wpb29/Create   新增一筆班 (wpb29 無 pb29995)
    #       Webm/Webm1031Wpb29/Update   修改既有班 (wpb29 有 pb29995)
    #       Webm/Webm1031Wpb29/Destroy  刪除一筆班
    #
    def _build_shift_payload(
        self,
        calendar: dict,
        filter_dict: dict,
        emp_id: str,
        date: DateLike,
        shift_code: str,
        kind: int = 3,
    ) -> dict:
        """組出 Create/Update 用的 payload。沿用 Fetch 回來的 employee 物件,
        並把整包 filter 攤平在最外層(對應前端 _.assign(data, filter))。"""
        emp = next(
            (e for e in calendar.get("shiftEmployees", [])
             if e.get("pa51002") == emp_id),
            None,
        )
        if emp is None:
            raise EhrsError(f"找不到員工 {emp_id}(請確認該員工在當月班表內)")

        target = _as_date(date)
        cell = next(
            (c for c in emp.get("cells", [])
             if _as_date(c["calendarDate"]) == target),
            None,
        )
        date_iso = f"{target.isoformat()}T00:00:00"
        if cell is None:
            # 該月尚未在 eHRS 開啟過/無格子，直接建立最小 wpb29（走 Create）
            wpb29 = {
                "pb29002": date_iso,
                "pb29003": emp_id,
                "pb29004": kind,
                "pb29005": shift_code,
                "pb29001": emp.get("pb29001") or emp.get("pa51001") or "",
                "pb29006": "",
                "pb29010": 0,
            }
        else:
            existing = (cell.get("schedules") or [None])[0]
            wpb29 = dict(existing) if existing else {}
            wpb29.update(
                {
                    "pb29002": date_iso,
                    "pb29003": emp_id,
                    "pb29004": kind,
                    "pb29005": shift_code,
                }
            )
            wpb29.setdefault("pb29001", existing.get("pb29001") if existing else "")
            wpb29.setdefault("pb29006", "")
            wpb29.setdefault("pb29010", 0)

        data = {
            "isCheckOverTime": False,
            "isCheckCanNotWorkingTime": False,
            "isCheckScheduleIntervals": False,
            "isCheckAllowConsecutive": False,
            "isCheckTotalWorkingHoursOfWeek": False,
            "isCheckPeriodicOverTime": False,
            "isCheckPeriodicHolidays": False,
            "isCheckPeriodic2WeekHolidays": False,
            "isCheckPeriodic1WeekHolidays": False,
            "wpb29": wpb29,
            "employee": emp,
            "movementWpb29": None,
            "isDestroyMovement": False,
        }
        data.update(filter_dict)  # type/start/end/year/month... 攤平在最外層
        return data

    def _post_shift_write(
        self,
        path: str,
        payload: dict,
        confirm: bool = True,
        max_rounds: int = 12,
    ) -> tuple[dict, list[dict]]:
        """送出排班寫入,並處理伺服器的確認回合(userResponse)。

        文中的寫入不是一次就成:伺服器遇到需提醒的規則(例:週期例假數)
        會回 userResponse={code, responseFlag, message},前端逐一確認後
        把 data[responseFlag]=True 再送一次,直到沒有 userResponse 才真正寫檔。
        code 90/91/99 都是「確認後繼續」,差別在前端 UI 呈現。

        confirm=False 時遇到需確認就中止丟例外(不自動代為確認)。
        回傳 (最後一次的 data 內層, 期間自動確認過的提醒清單)。
        """
        data = dict(payload)
        # confirm=True 時預先帶齊所有已知確認旗標,讓伺服器一趟就過(原本
        # 一回合只補一個旗標,要來回近 10 趟)。confirm=False 不預設,維持
        # 「遇到第一個需確認就停下丟例外」的行為。
        if confirm:
            for f in _WRITE_CONFIRM_FLAGS:
                data[f] = True
            data["selected"] = True
        confirmations: list[dict] = []
        for _ in range(max_rounds):
            envelope = self._post_json(path, data)
            inner = envelope.get("data", envelope) if isinstance(
                envelope, dict
            ) else envelope
            ur = inner.get("userResponse") if isinstance(inner, dict) else None
            if not ur:
                return inner, confirmations
            if not confirm:
                raise EhrsError(
                    f"寫入需確認(confirm=False 故中止):code={ur.get('code')} "
                    f"訊息={ur.get('message')}"
                )
            flag = ur.get("responseFlag")
            code = ur.get("code")
            confirmations.append(
                {"code": code, "flag": flag, "message": ur.get("message")}
            )
            if flag:
                data[flag] = True
            if code is not None:
                data["code"] = code
            if code == 91:
                data["selected"] = True
        raise EhrsError(
            f"確認回合超過上限 {max_rounds},疑似異常循環(已確認:{confirmations})"
        )

    def set_shift(
        self,
        year: int,
        month: int,
        emp_id: str,
        date: DateLike,
        shift_code: str,
        kind: Optional[int] = None,
        dry_run: bool = True,
        confirm: bool = True,
    ) -> dict:
        """建立或修改某員工某天的班別。

        會先抓當月班表,若該日已有班(有 pb29995)走 Update,否則走 Create。
        kind=None(預設)時自動從班表偵測該班別代號使用的種類,找不到預設 3。
        dry_run=True(預設)只回傳 payload 不送出。
        confirm=True 時自動通過伺服器的規則提醒(例:例假數),等同 UI 點確認;
        想看到提醒就停下時設 confirm=False。
        """
        filt = self._schedule_filter(year, month)
        calendar = self.get_schedule_raw(year, month)
        if kind is None:
            kind = self._kind_from_calendar(calendar, shift_code, default=3)
        payload = self._build_shift_payload(
            calendar, filt, emp_id, date, shift_code, kind
        )
        is_update = payload["wpb29"].get("pb29995") is not None
        path = (
            "Webm/Webm1031Wpb29/Update" if is_update
            else "Webm/Webm1031Wpb29/Create"
        )
        if dry_run:
            return {"dry_run": True, "endpoint": path, "payload": payload}
        result, confirmations = self._post_shift_write(
            path, payload, confirm=confirm
        )
        return {
            "ok": True,
            "endpoint": path,
            "confirmations": confirmations,
            "wpb29": result.get("wpb29") if isinstance(result, dict) else None,
        }

    def _kind_from_calendar(self, calendar: dict, shift_code: str, default: int = 3) -> int:
        for emp in calendar.get("shiftEmployees", []):
            for cell in emp.get("cells", []):
                for s in (cell.get("schedules") or []):
                    if s.get("pb29005") == shift_code and s.get("pb29004") is not None:
                        return s["pb29004"]
        return default

    def set_shifts_bulk(
        self,
        year: int,
        month: int,
        changes: list[dict],
        dry_run: bool = True,
        confirm: bool = True,
        progress_cb=None,        # Callable[[done: int, total: int], None]
        max_workers: int = 6,
    ) -> list[dict]:
        """批次建立/修改班別,並行送出請求。
        changes 每筆需含 emp_id、date、shift_code,可選 kind。
        eHRS 以「四週期間(期)」管理排班,跨月期間(如第7期 6/8～7/5)的
        7/1～7/5 其實屬於「六月」那次查詢回來的 calendar。因此這裡會把
        相關月份(各 date 的當月 + 前一個月)全部抓回,再針對每筆資料找出
        同時含「該員工 + 該日格子」的 calendar 來寫入。
        progress_cb(done, total) 每處理完一筆即呼叫一次。
        max_workers 控制並行 HTTP 連線數,預設 6。注意:eHRS 同一 session 的
        寫入會被伺服器序列化(ASP.NET session 鎖),調高 workers 幾乎不會更快,
        匯入加速主要來自每格寫入的確認回合從 ~10 趟壓成 1 趟(見 _post_shift_write)。"""
        # 要抓的「月曆月份」與「四週期間」。
        # ① 月曆檢視(shiftType=0)只回傳當月 1 號～月底,維持既有當月寫入行為。
        # ② 但跨月期間(如第7期 6/8～7/5)的 7/1～7/5 在任何月曆檢視都查不到
        #    (七月月曆尚未開,回傳 0 人),必須改用「彈性週期檢視」(shiftType=1)
        #    並把 start/end 設成該期間邊界,才會回傳 7/1～7/5 的格子。
        #    員工依彈性工時別分散在不同 flexType,故每個期間逐一查 flexType 再合併。
        month_keys: set[tuple[int, int]] = set()
        period_keys: set[tuple[_dt.date, _dt.date]] = set()
        for ch in changes:
            d = _as_date(ch["date"])
            month_keys.add((d.year, d.month))
            ps = _period_start_of(d)
            pe = ps + _dt.timedelta(days=_PERIOD_DAYS - 1)
            period_keys.add((ps, pe))

        # cals 順序很重要:月曆在前、期間在後。_pick_cal 取第一個「同時含
        # 該員工＋該日格子」者;當月日期(如 6/15)兩邊都有,會優先用月曆 filter
        # (維持既有可運作行為);溢出日期(如 7/1)只有期間檢視有,才落到期間。
        cals: list[tuple[dict, dict]] = []   # (filter, calendar)

        # ── ① 月曆檢視 ──────────────────────────────────────
        for (y, m) in sorted(month_keys):
            try:
                cal = self.get_schedule_raw(y, m)
            except EhrsError:
                continue
            cell_dates = [
                _as_date(c["calendarDate"])
                for e in cal.get("shiftEmployees", [])
                for c in e.get("cells", [])
            ]
            if not cell_dates:
                continue
            filt = dict(
                self._schedule_filter(y, m),
                start=_iso_ms(min(cell_dates)),
                end=_iso_ms(max(cell_dates), end_of_day=True),
            )
            cals.append((filt, cal))

        # ── ② 四週期間(彈性週期)檢視 ──────────────────────
        for (ps, pe) in sorted(period_keys):
            for ft in (2, 1, 0):   # 4週=2、雙週=1、單週=0;逐一查並合併
                ov = dict(
                    shiftType=1, flexType=ft,
                    start=_iso_ms(ps), end=_iso_ms(pe, end_of_day=True),
                )
                try:
                    cal = self.get_schedule_raw(ps.year, ps.month, **ov)
                except EhrsError:
                    continue
                if not cal.get("shiftEmployees"):
                    continue
                filt = dict(self._schedule_filter(ps.year, ps.month), **ov)
                cals.append((filt, cal))

        def _pick_cal(emp_id: str, date: str) -> tuple[dict | None, dict | None]:
            """找出最適合寫入此員工此日的 (filter, calendar)。
            優先:同時含該員工與該日格子;其次:含該員工(格子由 Create 補)。"""
            target = _as_date(date)
            for filt, cal in cals:
                emp = next(
                    (e for e in cal.get("shiftEmployees", [])
                     if e.get("pa51002") == emp_id),
                    None,
                )
                if emp and any(
                    _as_date(c["calendarDate"]) == target
                    for c in emp.get("cells", [])
                ):
                    return filt, cal
            for filt, cal in cals:
                if any(e.get("pa51002") == emp_id
                       for e in cal.get("shiftEmployees", [])):
                    return filt, cal
            return None, None

        total = len(changes)
        results: list[dict | None] = [None] * total
        done_count = [0]
        lock = threading.Lock()

        def _process(idx: int, ch: dict) -> None:
            emp_id     = ch["emp_id"]
            date       = ch["date"]
            shift_code = ch["shift_code"]
            filt, calendar = _pick_cal(emp_id, date)
            if calendar is None:
                results[idx] = {"ok": False, "emp_id": emp_id, "date": date,
                                "error": f"找不到員工 {emp_id}(請確認該員工在此期間班表內)"}
                _tick()
                return
            kind = ch.get("kind") or self._kind_from_calendar(calendar, shift_code)
            try:
                payload = self._build_shift_payload(
                    calendar, filt, emp_id, date, shift_code, kind
                )
            except EhrsError as exc:
                results[idx] = {"ok": False, "emp_id": emp_id, "date": date, "error": str(exc)}
                _tick()
                return
            is_update = payload["wpb29"].get("pb29995") is not None
            path = (
                "Webm/Webm1031Wpb29/Update" if is_update
                else "Webm/Webm1031Wpb29/Create"
            )
            if dry_run:
                results[idx] = {
                    "dry_run": True, "endpoint": path,
                    "emp_id": emp_id, "date": date, "shift_code": shift_code,
                }
                _tick()
                return
            try:
                result, confirmations = self._post_shift_write(path, payload, confirm=confirm)
                results[idx] = {
                    "ok": True, "endpoint": path,
                    "emp_id": emp_id, "date": date, "shift_code": shift_code,
                    "confirmations": confirmations,
                }
            except EhrsError as exc:
                results[idx] = {"ok": False, "emp_id": emp_id, "date": date, "error": str(exc)}
            _tick()

        def _tick() -> None:
            if progress_cb:
                with lock:
                    done_count[0] += 1
                    n = done_count[0]
                progress_cb(n, total)

        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = [pool.submit(_process, i, ch) for i, ch in enumerate(changes)]
            for f in as_completed(futures):
                f.result()   # re-raise any unexpected exception

        return [r for r in results if r is not None]

    def delete_shift(
        self,
        year: int,
        month: int,
        emp_id: str,
        date: DateLike,
        dry_run: bool = True,
    ) -> dict:
        """刪除某員工某天的班。dry_run=True(預設)只回傳 payload 不送出。"""
        filt = self._schedule_filter(year, month)
        calendar = self.get_schedule_raw(year, month)
        emp = next(
            (e for e in calendar.get("shiftEmployees", [])
             if e.get("pa51002") == emp_id),
            None,
        )
        if emp is None:
            raise EhrsError(f"找不到員工 {emp_id}")
        target = _as_date(date)
        cell = next(
            (c for c in emp.get("cells", [])
             if _as_date(c["calendarDate"]) == target),
            None,
        )
        existing = (cell.get("schedules") or [None])[0] if cell else None
        if not existing or existing.get("pb29995") is None:
            raise EhrsError(f"{emp_id} 在 {target.isoformat()} 沒有可刪除的班")
        payload = {
            "filter": filt,        # Destroy 的 filter 是巢狀 key(非攤平)
            "wpb29": existing,
            "employee": emp,
            "isDestroyMovement": False,
        }
        if dry_run:
            return {
                "dry_run": True,
                "endpoint": "Webm/Webm1031Wpb29/Destroy",
                "payload": payload,
            }
        return self._post_json("Webm/Webm1031Wpb29/Destroy", payload)
