"""文中 eHRS 命令列工具(在本資料夾執行)。

帳密走環境變數:
    export EHRS_ACCOUNT="帳號"
    export EHRS_PASSWORD="密碼"

指令:
    python3 -m ehrs schedule 2026 6               # 整月排班(摘要)
    python3 -m ehrs schedule 2026 6 --emp SA1588  # 某人整月每日班別
    python3 -m ehrs punch 2026-06-01 2026-06-05   # 打卡(預設查登入者本人)
    python3 -m ehrs punch 2026-06-01 2026-06-05 --emp SA2671
    python3 -m ehrs punch 2026-06-01 2026-06-05 --all
    python3 -m ehrs set SA1588 2026-06-08 991     # 排班(dry-run;加 --yes 才送出)
    python3 -m ehrs del SA1588 2026-06-08         # 刪班(dry-run;加 --yes 才送出)

寫入(set/del)預設只顯示將送出的內容(dry-run),確認無誤後加 --yes 才真的寫入。
"""
import argparse
import datetime as dt
import os
import sys

from ehrs_client import EhrsClient, EhrsError


def _login() -> tuple[EhrsClient, str]:
    acc = os.environ.get("EHRS_ACCOUNT")
    pwd = os.environ.get("EHRS_PASSWORD")
    if not acc or not pwd:
        sys.exit("請先設定環境變數 EHRS_ACCOUNT 與 EHRS_PASSWORD")
    try:
        return EhrsClient().login(acc, pwd), acc
    except EhrsError as exc:
        sys.exit(f"登入失敗:{exc}")


def _detect_kind(client: EhrsClient, year: int, month: int,
                 shift_code: str, default: int = 3) -> int:
    """從當月班表裡找這個班別代號用的種類(pb29004),找不到就用 default。"""
    try:
        cal = client.get_schedule_raw(year, month)
    except EhrsError:
        return default
    for emp in cal.get("shiftEmployees", []):
        for cell in emp.get("cells", []):
            for s in (cell.get("schedules") or []):
                if s.get("pb29005") == shift_code and s.get("pb29004") is not None:
                    return s["pb29004"]
    return default


def cmd_schedule(client: EhrsClient, acc: str, args) -> None:
    sched = client.get_schedule(args.year, args.month)
    if args.emp:
        emp = next((e for e in sched if e.emp_id == args.emp), None)
        if not emp:
            sys.exit(f"當月班表找不到員工 {args.emp}")
        print(f"{emp.emp_id} {emp.name} [{emp.title}] — {args.year}-{args.month:02d}")
        for s in sorted(emp.shifts, key=lambda x: x.date):
            print(f"  {s.date}  {s.code:<6} {s.name}")
    else:
        print(f"{args.year}-{args.month:02d} 排班(共 {len(sched)} 人)")
        for e in sched:
            print(f"  {e.emp_id}  {e.name:<6} [{e.title}]  班數={len(e.shifts)} 休={e.rest_days}")


def cmd_punch(client: EhrsClient, acc: str, args) -> None:
    emp = None if args.all else (args.emp or acc)
    rows = client.get_punch_records(
        args.start, args.end, emp_start=emp, emp_end=emp, readable=True
    )
    scope = "全員" if args.all else emp
    print(f"打卡 {args.start}~{args.end}({scope},共 {len(rows)} 筆)")
    for r in rows:
        print(
            f"  {r.get('員工代號', ''):<8}{r.get('員工姓名', ''):<6} "
            f"{str(r.get('出勤日期', ''))[:10]} {str(r.get('刷卡時間', ''))[11:16]} "
            f"{r.get('來源', '')}"
        )


def cmd_set(client: EhrsClient, acc: str, args) -> None:
    d = dt.date.fromisoformat(args.date)
    kind = args.kind if args.kind is not None else _detect_kind(
        client, d.year, d.month, args.shift
    )
    if args.yes:
        res = client.set_shift(
            d.year, d.month, args.emp, args.date, args.shift,
            kind=kind, dry_run=False,
        )
        wp = res.get("wpb29") or {}
        action = "修改" if res["endpoint"].endswith("Update") else "新增"
        print(f"已{action}:{args.emp} {args.date} → {wp.get('pb29005')} "
              f"(pk={wp.get('pb29995')})")
        msgs = [c["message"] for c in res["confirmations"] if c.get("message")]
        if msgs:
            print("  已通過的規則確認:")
            for msg in msgs:
                print(f"    - {msg}")
    else:
        res = client.set_shift(
            d.year, d.month, args.emp, args.date, args.shift,
            kind=kind, dry_run=True,
        )
        action = "修改" if res["endpoint"].endswith("Update") else "新增"
        print(f"[dry-run] 將{action}:{args.emp} {args.date} → 班別 {args.shift} "
              f"(kind={kind})")
        print(f"  endpoint: {res['endpoint']}")
        print("  確認無誤後,加上 --yes 才會真的送出。")


def cmd_del(client: EhrsClient, acc: str, args) -> None:
    d = dt.date.fromisoformat(args.date)
    if args.yes:
        client.delete_shift(d.year, d.month, args.emp, args.date, dry_run=False)
        print(f"已刪除:{args.emp} {args.date} 的班")
    else:
        res = client.delete_shift(
            d.year, d.month, args.emp, args.date, dry_run=True
        )
        wp = res["payload"]["wpb29"]
        print(f"[dry-run] 將刪除:{args.emp} {args.date} 目前的班 "
              f"{wp.get('pb29005')} (pk={wp.get('pb29995')})")
        print("  確認無誤後,加上 --yes 才會真的送出。")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="ehrs", description="文中 eHRS 排班/打卡工具")
    sub = p.add_subparsers(dest="command")

    ps = sub.add_parser("schedule", help="查排班")
    ps.add_argument("year", type=int)
    ps.add_argument("month", type=int)
    ps.add_argument("--emp", help="只看某員工,顯示每日班別")

    pp = sub.add_parser("punch", help="查打卡")
    pp.add_argument("start", help="起日 YYYY-MM-DD")
    pp.add_argument("end", help="迄日 YYYY-MM-DD")
    pp.add_argument("--emp", help="指定員工代號(預設查登入者本人)")
    pp.add_argument("--all", action="store_true", help="查全部員工")

    pset = sub.add_parser("set", help="建立/修改排班")
    pset.add_argument("emp", help="員工代號")
    pset.add_argument("date", help="日期 YYYY-MM-DD")
    pset.add_argument("shift", help="班別代號,例如 991")
    pset.add_argument("--kind", type=int, default=None, help="班別種類(預設自動判斷)")
    pset.add_argument("--yes", action="store_true", help="真的送出(否則只 dry-run)")

    pdel = sub.add_parser("del", help="刪除排班")
    pdel.add_argument("emp", help="員工代號")
    pdel.add_argument("date", help="日期 YYYY-MM-DD")
    pdel.add_argument("--yes", action="store_true", help="真的送出(否則只 dry-run)")
    return p


def main(argv=None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.command:
        parser.print_help()
        return
    client, acc = _login()
    handlers = {
        "schedule": cmd_schedule,
        "punch": cmd_punch,
        "set": cmd_set,
        "del": cmd_del,
    }
    handlers[args.command](client, acc, args)


if __name__ == "__main__":
    try:
        main()
    except EhrsError as exc:
        sys.exit(f"操作失敗:{exc}")
