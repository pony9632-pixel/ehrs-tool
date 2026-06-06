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

# bash 結束後自動關閉終端機視窗，不顯示「有執行中程序」確認對話框
# 做法：把 AppleScript 放到背景、disown 讓它在 bash 結束後繼續存活，
#       bash 一 exit 就不再有 shell 程序，Terminal 關窗時不會再詢問。
( sleep 0.4 && osascript -e 'tell application "Terminal" to close front window' ) \
    >/dev/null 2>&1 &
disown $!
exit 0
