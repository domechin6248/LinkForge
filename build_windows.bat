@echo off
chcp 65001 > nul
REM ════════════════════════════════════════════════════════════════
REM  楽々JC v2.0.0  ─  Windows ビルドスクリプト
REM  このファイルをダブルクリックして実行してください
REM ════════════════════════════════════════════════════════════════

echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo   楽々JC v2.0.0  Windows ビルド
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.

REM ── 依存チェック ────────────────────────────────────────────────
echo [ 1/5 ] 依存パッケージを確認中...

python -c "import docx" 2>nul || (
    echo   python-docx をインストール中...
    pip install python-docx -q
)
python -c "import openpyxl" 2>nul || (
    echo   openpyxl をインストール中...
    pip install openpyxl -q
)
python -c "from pptx import Presentation" 2>nul || (
    echo   python-pptx をインストール中...
    pip install python-pptx -q
)
python -c "import pdfplumber" 2>nul || (
    echo   pdfplumber をインストール中...
    pip install pdfplumber -q
)
python -c "import tkinterdnd2" 2>nul || (
    echo   tkinterdnd2 をインストール中...
    pip install tkinterdnd2 -q
)
python -c "import PyInstaller" 2>nul || (
    echo   PyInstaller をインストール中...
    pip install pyinstaller -q
)

echo   依存パッケージ OK
echo.

REM ── アイコン確認 ─────────────────────────────────────────────────
echo [ 2/5 ] アイコンを確認中...

if not exist "icon.ico" (
    if exist "icon_src.png" (
        echo   icon_src.png からアイコンを生成中...
        python build_icon.py icon_src.png
    ) else if exist "icon.png" (
        echo   icon.png からアイコンを生成中...
        python build_icon.py icon.png
    ) else (
        echo   警告: icon.ico が見つかりません。
        echo   icon_src.png をこのフォルダに置いてから再実行することを推奨します。
        echo   アイコンなしで続行します...
    )
) else (
    echo   icon.ico 確認済み
)
echo.

REM ── 旧ビルド削除 ─────────────────────────────────────────────────
echo [ 3/5 ] 旧ビルドをクリア中...
if exist "dist\楽々JC" rmdir /s /q "dist\楽々JC"
if exist "dist\楽々JC.exe" del /q "dist\楽々JC.exe"
if exist "build\楽々JC" rmdir /s /q "build\楽々JC"
echo   クリア完了
echo.

REM ── PyInstaller 実行 ─────────────────────────────────────────────
echo [ 4/5 ] PyInstaller でビルド中...
echo   （数分かかります）
echo.

if exist "icon.ico" (
    python -m PyInstaller "楽々JC_Windows.spec" --noconfirm
) else (
    python -m PyInstaller linkforge.py ^
        --onefile --windowed --noconfirm ^
        --name "楽々JC" ^
        --hidden-import=docx ^
        --hidden-import=docx.oxml ^
        --hidden-import=docx.oxml.ns ^
        --hidden-import=lxml ^
        --hidden-import=lxml._elementpath
)

echo.
echo [ 5/5 ] 完了チェック中...
echo.

echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if exist "dist\楽々JC.exe" (
    echo   dist\楽々JC.exe が作成されました！
    echo   このファイルをダブルクリックして起動できます。
) else (
    echo   ビルドに失敗しました。上記のエラーを確認してください。
)
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
pause
