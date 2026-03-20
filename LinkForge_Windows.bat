@echo off
chcp 65001 >nul
title LinkForge

cd /d "%~dp0"

REM Python の確認
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ================================================
    echo   [エラー] Python がインストールされていません。
    echo   https://www.python.org からインストールしてください。
    echo   （インストール時に Add Python to PATH にチェック）
    echo ================================================
    pause
    exit /b 1
)

REM python-docx の確認・自動インストール
python -c "import docx" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo python-docx をインストールしています...
    python -m pip install python-docx
)

REM アプリ起動
python linkforge.py
