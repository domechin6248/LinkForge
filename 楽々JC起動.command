#!/bin/bash
# ─────────────────────────────────────────
#  楽々JC v2.0.0 起動スクリプト
#  このファイルをダブルクリックするとアプリが起動します
# ─────────────────────────────────────────

# スクリプトのあるフォルダに移動
cd "$(dirname "$0")"

# 必要ライブラリがなければ自動インストール
python3 -c "import docx" 2>/dev/null || {
    echo "python-docx をインストール中..."
    python3 -m pip install python-docx --quiet
}
python3 -c "import openpyxl" 2>/dev/null || {
    echo "openpyxl をインストール中..."
    python3 -m pip install openpyxl --quiet
}
python3 -c "from pptx import Presentation" 2>/dev/null || {
    echo "python-pptx をインストール中..."
    python3 -m pip install python-pptx --quiet
}
python3 -c "import pdfplumber" 2>/dev/null || {
    echo "pdfplumber をインストール中..."
    python3 -m pip install pdfplumber --quiet
}
python3 -c "import tkinterdnd2" 2>/dev/null || {
    echo "tkinterdnd2 をインストール中..."
    python3 -m pip install tkinterdnd2 --quiet
}

# バックグラウンドで起動（ターミナルがブロックされないようにする）
python3 linkforge.py &
disown $!

# ターミナルウィンドウを閉じる
sleep 0.3
osascript -e 'tell application "Terminal" to close first window' 2>/dev/null || true
