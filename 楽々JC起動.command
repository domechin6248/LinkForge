#!/bin/bash
# ─────────────────────────────────────────
#  楽々JC v2.0.0 起動スクリプト
#  このファイルをダブルクリックするとアプリが起動します
# ─────────────────────────────────────────

# スクリプトのあるフォルダに移動
cd "$(dirname "$0")"

# python-docx がなければインストール
python3 -c "import docx" 2>/dev/null || {
    echo "python-docx をインストール中..."
    python3 -m pip install python-docx --quiet
}

# アプリ起動
python3 linkforge.py
