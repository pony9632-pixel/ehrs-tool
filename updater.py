"""自動更新模組 — 從 GitHub Releases 下載最新版本並重新啟動。

使用方式（在 app.py 啟動後呼叫一次）：
    from updater import start_update_check
    start_update_check(status_cb=self._set_status)
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


# --------------------------------------------------------------------------- #
# 公開 API
# --------------------------------------------------------------------------- #
def start_update_check(status_cb: Optional[Callable[[str], None]] = None) -> None:
    """啟動背景執行緒檢查 GitHub 最新 Release。
    找到新版本時自動下載並重新啟動；失敗時靜默（不影響正常使用）。
    status_cb 用於在 UI 狀態列顯示進度文字。
    """
    from version import __version__

    def _run() -> None:
        try:
            resp = requests.get(API_URL, timeout=10,
                                headers={"Accept": "application/vnd.github+json"})
            if resp.status_code != 200:
                return
            data = resp.json()
            tag: str = data.get("tag_name", "")
            latest = tag.lstrip("v")
            if not latest or not _is_newer(latest, __version__):
                return  # 已是最新版

            # 找 zip asset；沒有則用 GitHub 自動產的 zipball
            zip_url: str = data.get("zipball_url", "")
            for asset in data.get("assets", []):
                if asset["name"].endswith(".zip"):
                    zip_url = asset["browser_download_url"]
                    break

            if not zip_url:
                return

            _download_and_restart(latest, zip_url, status_cb)

        except Exception:
            pass  # 網路離線或 API 失敗 → 不影響 app 使用

    threading.Thread(target=_run, daemon=True).start()


# --------------------------------------------------------------------------- #
# 下載 + 解壓 + 重啟（在背景執行緒內執行）
# --------------------------------------------------------------------------- #
def _download_and_restart(
    version: str,
    zip_url: str,
    status_cb: Optional[Callable[[str], None]],
) -> None:
    def _cb(msg: str) -> None:
        if status_cb:
            status_cb(msg)

    try:
        _cb(f"發現新版本 v{version}，下載中…")
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

            # GitHub zipball 最外層是 owner-repo-<hash>/ 目錄
            src_dirs = [
                p for p in Path(tmp).iterdir()
                if p.is_dir() and p.name != "__MACOSX"
            ]
            if not src_dirs:
                _cb("更新失敗：解壓後找不到目錄")
                return
            src = src_dirs[0]

            # 複製所有 .py 檔與 .command 啟動檔
            for pattern in ("*.py", "*.command"):
                for f in src.glob(pattern):
                    dest = _APP_DIR / f.name
                    shutil.copy2(f, dest)
                    if pattern == "*.command":
                        dest.chmod(0o755)

        _cb(f"v{version} 安裝完成，重新啟動中…")

        # 重新啟動 app.py
        os.execv(sys.executable,
                 [sys.executable, str(_APP_DIR / "app.py")])

    except Exception as exc:
        _cb(f"更新失敗：{exc}")
