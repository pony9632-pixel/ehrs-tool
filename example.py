"""ehrs_client 使用範例。

帳號密碼從環境變數讀取,不寫死在程式碼裡:

    export EHRS_ACCOUNT="你的帳號"
    export EHRS_PASSWORD="你的密碼"
    python3 example.py

(也可放進 .env 後 `source .env`;.env 已被 .gitignore 排除)
"""
import datetime as dt
import os
import sys

from ehrs_client import EhrsClient, EhrsError


def get_credentials() -> tuple[str, str]:
    acc = os.environ.get("EHRS_ACCOUNT")
    pwd = os.environ.get("EHRS_PASSWORD")
    if not acc or not pwd:
        sys.exit(
            "請先設定環境變數:\n"
            '  export EHRS_ACCOUNT="你的帳號"\n'
            '  export EHRS_PASSWORD="你的密碼"'
        )
    return acc, pwd


def main() -> None:
    acc, pwd = get_credentials()
    client = EhrsClient().login(acc, pwd)
    print(f"登入成功:{acc}")

    today = dt.date.today()

    # --- 排班讀取 ---
    schedule = client.get_schedule(today.year, today.month)
    print(f"\n== {today.year}-{today.month:02d} 排班(共 {len(schedule)} 人)==")
    for emp in schedule:
        first = emp.shifts[0] if emp.shifts else None
        sample = f"{first.date} {first.name}" if first else "(無班)"
        print(f"  {emp.emp_id} {emp.name:<6} [{emp.title}] 班數={len(emp.shifts)}  例:{sample}")

    # --- 打卡讀取(最近 7 天,限自己這個帳號)---
    start = today - dt.timedelta(days=7)
    punches = client.get_punch_records(
        start, today, emp_start=acc, emp_end=acc, readable=True
    )
    print(f"\n== 打卡 {start}~{today}(共 {len(punches)} 筆)==")
    for row in punches:
        print(f"  {row.get('出勤日期', '')[:10]}  {row.get('刷卡時間', '')[11:16]}  {row.get('來源', '')}")

    # --- 排班寫入(預設 dry_run,只印 payload 不送出)---
    # 真的要寫入時把 dry_run 改 False;會自動通過伺服器的規則確認。
    if schedule:
        emp = schedule[0]
        preview = client.set_shift(
            today.year, today.month, emp.emp_id,
            today.replace(day=min(28, today.day)), "991",
            dry_run=True,
        )
        print("\n== set_shift dry-run 預覽 ==")
        print("  endpoint:", preview["endpoint"])
        print("  wpb29:", {k: preview["payload"]["wpb29"].get(k)
                           for k in ("pb29002", "pb29003", "pb29005", "pb29995")})

    # 實際寫入範例(預設註解掉,避免誤動正式班表):
    # client.set_shift(2026, 6, "SA1588", "2026-06-08", "991", dry_run=False)
    # client.delete_shift(2026, 6, "SA1588", "2026-06-08", dry_run=False)


if __name__ == "__main__":
    try:
        main()
    except EhrsError as exc:
        sys.exit(f"操作失敗:{exc}")
