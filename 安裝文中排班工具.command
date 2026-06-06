#!/bin/bash
# 文中排班工具 — 一鍵安裝腳本
# 雙擊此檔案即可自動下載並安裝最新版本。

cd "$(dirname "$0")"

# ── 顏色 ──────────────────────────────────────────
R='\033[0;31m'; G='\033[0;32m'; Y='\033[1;33m'
B='\033[0;34m'; C='\033[0;36m'; NC='\033[0m'

echo ""
echo -e "${B}╔════════════════════════════════════════╗${NC}"
echo -e "${B}║     文中排班工具  一鍵安裝程式         ║${NC}"
echo -e "${B}╚════════════════════════════════════════╝${NC}"
echo ""

INSTALL_DIR="$HOME/Desktop/文中排班工具"
REPO="pony9632-pixel/ehrs-tool"
API="https://api.github.com/repos/$REPO/releases/latest"

# ── 1. 確認 Python3 ───────────────────────────────
echo -e "${C}[1/4]${NC} 確認 Python3…"
if ! command -v python3 &>/dev/null; then
    echo -e "${R}✗ 找不到 Python3。請先安裝：${NC}"
    echo "    https://www.python.org/downloads/macos/"
    open "https://www.python.org/downloads/macos/"
    read -rp "安裝完成後按 Enter 重試，或關閉此視窗…"
    if ! command -v python3 &>/dev/null; then
        echo -e "${R}仍然找不到 Python3，安裝中止。${NC}"; exit 1
    fi
fi
PY_VER=$(python3 --version 2>&1 | cut -d' ' -f2)
echo -e "    ${G}✓ Python3 $PY_VER${NC}"

# ── 2. 取得最新版下載連結 ─────────────────────────
echo -e "${C}[2/4]${NC} 取得最新版本資訊…"
ZIP_URL=$(curl -fsSL "$API" 2>/dev/null \
    | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('zipball_url',''))" 2>/dev/null)
TAG=$(curl -fsSL "$API" 2>/dev/null \
    | python3 -c "import sys,json; d=json.load(sys.stdin); print(d.get('tag_name',''))" 2>/dev/null)

if [ -z "$ZIP_URL" ]; then
    echo -e "${R}✗ 無法連線至 GitHub（請確認網路連線後重試）${NC}"
    read -rp "按 Enter 關閉…"; exit 1
fi
echo -e "    ${G}✓ 最新版本：$TAG${NC}"

# ── 3. 下載並解壓 ─────────────────────────────────
echo -e "${C}[3/4]${NC} 下載中（$TAG）…"
TMP=$(mktemp -d)
if ! curl -fsSL -o "$TMP/app.zip" "$ZIP_URL"; then
    echo -e "${R}✗ 下載失敗${NC}"
    rm -rf "$TMP"; read -rp "按 Enter 關閉…"; exit 1
fi

unzip -q "$TMP/app.zip" -d "$TMP/src"
SRC=$(find "$TMP/src" -mindepth 1 -maxdepth 1 -type d | head -1)
if [ -z "$SRC" ]; then
    echo -e "${R}✗ 解壓失敗${NC}"
    rm -rf "$TMP"; read -rp "按 Enter 關閉…"; exit 1
fi

mkdir -p "$INSTALL_DIR"
# 複製所有 .py 檔
find "$SRC" -maxdepth 1 -name "*.py" -exec cp {} "$INSTALL_DIR/" \;
# 複製 .command 啟動檔
find "$SRC" -maxdepth 1 -name "*.command" -exec cp {} "$INSTALL_DIR/" \;
chmod +x "$INSTALL_DIR/"*.command 2>/dev/null
rm -rf "$TMP"
echo -e "    ${G}✓ 檔案已安裝至桌面「文中排班工具」資料夾${NC}"

# ── 4. 安裝 Python 相依套件 ───────────────────────
echo -e "${C}[4/4]${NC} 安裝相依套件（首次約需 30 秒）…"
python3 -m pip install --upgrade customtkinter openpyxl requests --quiet
echo -e "    ${G}✓ 套件安裝完成${NC}"

# ── 完成 ──────────────────────────────────────────
echo ""
echo -e "${G}╔════════════════════════════════════════╗${NC}"
echo -e "${G}║  安裝完成！                            ║${NC}"
echo -e "${G}║                                        ║${NC}"
echo -e "${G}║  之後直接雙擊桌面資料夾裡的            ║${NC}"
echo -e "${G}║  「啟動文中排班工具.command」即可。    ║${NC}"
echo -e "${G}║                                        ║${NC}"
echo -e "${G}║  程式會自動檢查並更新到最新版本。      ║${NC}"
echo -e "${G}╚════════════════════════════════════════╝${NC}"
echo ""

open "$INSTALL_DIR"

read -rp "按 Enter 關閉此視窗…"
osascript -e 'tell application "Terminal" to close front window' 2>/dev/null
