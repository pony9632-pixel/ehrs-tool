"""桌面 App 的本機設定(帳密)讀寫。

存在使用者家目錄 ~/.ehrs_tool/config.json,權限設 0600(只有自己可讀),
不放在專案資料夾、不進 git。也支援用環境變數覆蓋。
"""
import json
import os
from pathlib import Path

CONFIG_DIR = Path.home() / ".ehrs_tool"
CONFIG_PATH = CONFIG_DIR / "config.json"


def load_config() -> dict:
    data: dict = {}
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            data = {}
    # 環境變數優先(方便臨時覆蓋,不寫檔)
    if os.environ.get("EHRS_ACCOUNT"):
        data["account"] = os.environ["EHRS_ACCOUNT"]
    if os.environ.get("EHRS_PASSWORD"):
        data["password"] = os.environ["EHRS_PASSWORD"]
    if os.environ.get("EHRS_BASE_URL"):
        data["base_url"] = os.environ["EHRS_BASE_URL"]
    return data


def save_config(data: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    try:
        os.chmod(CONFIG_PATH, 0o600)  # 內含密碼,限本人讀寫
    except OSError:
        pass
