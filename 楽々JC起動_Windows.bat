@echo off
chcp 65001 > nul
REM ════════════════════════════════════════════════════════════════
REM  楽々JC v2.0.0  ─  Windows 起動スクリプト（ビルド不要版）
REM  このファイルをダブルクリックするとアプリが起動します
REM  ※ Python 3.8 以上が必要です
REM ════════════════════════════════════════════════════════════════

cd /d "%~dp0"

REM python-docx がなければインストール
python -c "import docx" 2>nul || (
    echo python-docx をインストール中...
    pip install python-docx -q
)

REM アプリ起動（コンソールウィンドウを非表示にして起動）
start "" pythonw linkforge.py
