#!/bin/bash
# 切到腳本所在目錄，確保從任何地方雙擊都能正確執行
cd "$(dirname "$0")"

# 缺少相依套件時自動安裝
python3 -c "import customtkinter, openpyxl, requests" 2>/dev/null || {
    echo "正在安裝相依套件…"
    python3 -m pip install customtkinter openpyxl requests --quiet
}

# 啟動前先檢查並下載最新版本
python3 updater.py

# 啟動主程式
python3 app.py

osascript -e 'tell application "Terminal" to close front window' 2>/dev/null
