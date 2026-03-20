@echo off
chcp 65001 >nul
REM ================================================================
REM  Rakuraku JC v2.0.0 - Startup Script (Windows)
REM  Double-click this file to launch the app
REM ================================================================

cd /d "%~dp0"

REM ── Check Python installation ───────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Python is not installed.
    echo Please install it from https://www.python.org/
    echo IMPORTANT: Check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM ── Check and install required libraries ────────────────────────
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installing python-docx...
    python -m pip install python-docx --quiet
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    python -m pip install openpyxl --quiet
)

python -c "from pptx import Presentation" >nul 2>&1
if errorlevel 1 (
    echo Installing python-pptx...
    python -m pip install python-pptx --quiet
)

python -c "import pdfplumber" >nul 2>&1
if errorlevel 1 (
    echo Installing pdfplumber...
    python -m pip install pdfplumber --quiet
)

python -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    echo Installing tkinterdnd2...
    python -m pip install tkinterdnd2 --quiet
)

REM ── Launch app (no console window) ──────────────────────────────
where pythonw >nul 2>&1
if errorlevel 1 (
    start "" python linkforge.py
) else (
    start "" pythonw linkforge.py
)

exit
