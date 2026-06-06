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
                )
            )
        return result

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
        if cell is None:
            raise EhrsError(f"班表中找不到日期 {target.isoformat()}")

        existing = (cell.get("schedules") or [None])[0]
        date_iso = f"{target.isoformat()}T00:00:00"
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
    ) -> list[dict]:
        """批次建立/修改班別,只抓一次班表。
        changes 每筆需含 emp_id、date、shift_code,可選 kind。"""
        filt = self._schedule_filter(year, month)
        calendar = self.get_schedule_raw(year, month)
        results: list[dict] = []
        for ch in changes:
            emp_id = ch["emp_id"]
            date = ch["date"]
            shift_code = ch["shift_code"]
            kind = ch.get("kind") or self._kind_from_calendar(calendar, shift_code)
            try:
                payload = self._build_shift_payload(
                    calendar, filt, emp_id, date, shift_code, kind
                )
            except EhrsError as exc:
                results.append({"ok": False, "emp_id": emp_id, "date": date, "error": str(exc)})
                continue
            is_update = payload["wpb29"].get("pb29995") is not None
            path = (
                "Webm/Webm1031Wpb29/Update" if is_update
                else "Webm/Webm1031Wpb29/Create"
            )
            if dry_run:
                results.append({
                    "dry_run": True, "endpoint": path,
                    "emp_id": emp_id, "date": date, "shift_code": shift_code,
                })
                continue
            try:
                result, confirmations = self._post_shift_write(path, payload, confirm=confirm)
                results.append({
                    "ok": True, "endpoint": path,
                    "emp_id": emp_id, "date": date, "shift_code": shift_code,
                    "confirmations": confirmations,
                })
            except EhrsError as exc:
                results.append({"ok": False, "emp_id": emp_id, "date": date, "error": str(exc)})
        return results

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
