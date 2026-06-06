"""自動更新模組 — 從 GitHub Releases 下載最新版本。

兩種使用方式：
  1. 命令列（啟動腳本呼叫）：python3 updater.py
     → 同步執行，有進度輸出，完成後正常退出讓 .command 繼續啟動 app。
  2. App 內部（背景執行緒）：from updater import start_update_check
     → 非同步執行，用 status_cb 更新 UI 狀態列，找到新版本後重啟 app。
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import threading
import zipfile
from pathlib import Path
from typing import Callable, Optional

import requests

GITHUB_REPO = "pony9632-pixel/ehrs-tool"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
_APP_DIR = Path(__file__).parent.resolve()


# --------------------------------------------------------------------------- #
# 版本比較
# --------------------------------------------------------------------------- #
def _parse_ver(v: str) -> tuple[int, ...]:
    try:
        return tuple(int(x) for x in v.lstrip("v").split("."))
    except ValueError:
        return (0,)


def _is_newer(remote: str, local: str) -> bool:
    return _parse_ver(remote) > _parse_ver(local)


def _get_latest_release() -> tuple[str, str] | tuple[None, None]:
    """回傳 (latest_version, zip_url)，失敗回傳 (None, None)。"""
    resp = requests.get(API_URL, timeout=10,
                        headers={"Accept": "application/vnd.github+json"})
    if resp.status_code != 200:
        return None, None
    data = resp.json()
    tag: str = data.get("tag_name", "")
    latest = tag.lstrip("v")
    if not latest:
        return None, None
    zip_url: str = data.get("zipball_url", "")
    for asset in data.get("assets", []):
        if asset["name"].endswith(".zip"):
            zip_url = asset["browser_download_url"]
            break
    return latest, zip_url


def _install_zip(zip_url: str, cb: Callable[[str], None]) -> bool:
    """下載 zip → 解壓 → 覆蓋 .py 與 .command。成功回傳 True。"""
    resp = requests.get(zip_url, timeout=120, stream=True)
    resp.raise_for_status()

    with tempfile.TemporaryDirectory() as tmp:
        zip_path = os.path.join(tmp, "update.zip")
        with open(zip_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)

        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(tmp)

        src_dirs = [p for p in Path(tmp).iterdir()
                    if p.is_dir() and p.name != "__MACOSX"]
        if not src_dirs:
            cb("更新失敗：解壓後找不到目錄")
            return False
        src = src_dirs[0]

        for pattern in ("*.py", "*.command"):
            for f in src.glob(pattern):
                dest = _APP_DIR / f.name
                shutil.copy2(f, dest)
                if pattern == "*.command":
                    dest.chmod(0o755)

    return True


# --------------------------------------------------------------------------- #
# 1. CLI 同步模式（.command 啟動腳本呼叫）
# --------------------------------------------------------------------------- #
def check_and_update_cli() -> None:
    """同步版本檢查，供 .command 腳本在 app 啟動前呼叫。
    有更新 → 下載安裝後繼續（不重啟，直接讓 .command 跑新版 app.py）。
    無更新 / 失敗 → 靜默繼續。
    """
    from version import __version__

    print(f"[更新] 目前版本 v{__version__}，檢查中…", flush=True)
    try:
        latest, zip_url = _get_latest_release()
        if latest is None:
            print("[更新] 無法連線 GitHub，略過", flush=True)
            return
        if not _is_newer(latest, __version__):
            print(f"[更新] 已是最新版本 ✓", flush=True)
            return

        print(f"[更新] 發現新版本 v{latest}，下載中…", flush=True)
        ok = _install_zip(zip_url, lambda m: print(f"[更新] {m}", flush=True))
        if ok:
            print(f"[更新] v{latest} 安裝完成，啟動新版程式 ✓", flush=True)
    except Exception as exc:
        print(f"[更新] 檢查失敗：{exc}，略過", flush=True)


# --------------------------------------------------------------------------- #
# 2. App 內背景模式（app.py 呼叫，找到更新則重啟）
# --------------------------------------------------------------------------- #
def start_update_check(status_cb: Optional[Callable[[str], None]] = None) -> None:
    """背景執行緒檢查更新；找到新版本時下載後重啟 app。"""
    from version import __version__

    def _run() -> None:
        try:
            latest, zip_url = _get_latest_release()
            if latest is None or not _is_newer(latest, __version__):
                return
            if not zip_url:
                return

            def _cb(msg: str) -> None:
                if status_cb:
                    status_cb(msg)

            _cb(f"發現新版本 v{latest}，下載中…")
            ok = _install_zip(zip_url, _cb)
            if ok:
                _cb(f"v{latest} 安裝完成，重新啟動中…")
                os.execv(sys.executable,
                         [sys.executable, str(_APP_DIR / "app.py")])
        except Exception:
            pass

    threading.Thread(target=_run, daemon=True).start()


# --------------------------------------------------------------------------- #
# 命令列入口
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    check_and_update_cli()
