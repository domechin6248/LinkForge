@echo off
chcp 65001 >nul
rem ─────────────────────────────────────────
rem  楽々JC v2.0.0 起動スクリプト（Windows）
rem  このファイルをダブルクリックするとアプリが起動します
rem ─────────────────────────────────────────

cd /d "%~dp0"

rem ── Python のインストール確認 ───────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo [エラー] Python がインストールされていません。
    echo https://www.python.org/ からインストールしてください。
    echo インストール時に「Add Python to PATH」にチェックを入れてください。
    echo.
    pause
    exit /b 1
)

rem ── 必要ライブラリの確認・自動インストール ─────────────────────
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo python-docx をインストール中...
    python -m pip install python-docx --quiet
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo openpyxl をインストール中...
    python -m pip install openpyxl --quiet
)

python -c "from pptx import Presentation" >nul 2>&1
if errorlevel 1 (
    echo python-pptx をインストール中...
    python -m pip install python-pptx --quiet
)

python -c "import pdfplumber" >nul 2>&1
if errorlevel 1 (
    echo pdfplumber をインストール中...
    python -m pip install pdfplumber --quiet
)

python -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    echo tkinterdnd2 をインストール中...
    python -m pip install tkinterdnd2 --quiet
)

rem ── アプリ起動（コンソールウィンドウなしで起動） ───────────────
where pythonw >nul 2>&1
if errorlevel 1 (
    start "" python linkforge.py
) else (
    start "" pythonw linkforge.py
)

exit
