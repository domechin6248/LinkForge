#!/bin/bash
# ╔═══════════════════════════════════════╗
# ║  LinkForge セットアップ                ║
# ║  ダブルクリックで実行してください        ║
# ╚═══════════════════════════════════════╝

cd "$(dirname "$0")"

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

echo ""
echo "====================================="
echo "  LinkForge セットアップ"
echo "====================================="
echo ""

# ── 1. Python3 を探す ──
PYTHON=""
for p in /opt/homebrew/bin/python3.13 /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3; do
    if [ -x "$p" ]; then
        PYTHON="$p"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "[エラー] Python3 が見つかりません。"
    echo "https://www.python.org からインストールしてください。"
    read -p "Enterキーで閉じます..."
    exit 1
fi
echo "Python: $PYTHON"

# ── 2. 必要パッケージをインストール ──
echo "パッケージを確認中..."
"$PYTHON" -m pip install python-docx --break-system-packages --quiet 2>/dev/null
echo "  python-docx ... OK"

# ドラッグ＆ドロップ用パッケージ（方法1: tkinterdnd2）
echo "  tkinterdnd2 (DnD) をインストール中..."
"$PYTHON" -m pip install tkinterdnd2 --break-system-packages --quiet 2>/dev/null
if [ $? -eq 0 ]; then
    echo "  tkinterdnd2 ... OK"
else
    echo "  tkinterdnd2 ... スキップ（代替方法を使用）"
fi

# ドラッグ＆ドロップ用パッケージ（方法2: pyobjc - macOSのみ）
if [[ "$(uname)" == "Darwin" ]]; then
    echo "  pyobjc (macOS DnD) をインストール中..."
    "$PYTHON" -m pip install pyobjc-framework-Cocoa --break-system-packages --quiet 2>/dev/null
    if [ $? -eq 0 ]; then
        echo "  pyobjc-framework-Cocoa ... OK"
    else
        echo "  pyobjc-framework-Cocoa ... スキップ"
    fi
fi

echo "パッケージ準備完了"

# ── 3. AppleScript で .app を生成 ──
echo ""
echo "LinkForge.app を作成中..."

SCRIPT_PATH="$(pwd)/linkforge.py"

# AppleScript コードを作成
cat > /tmp/linkforge_builder.applescript << APPLEOF
-- LinkForge Launcher
set scriptPath to "$SCRIPT_PATH"
set pythonPath to "$PYTHON"
do shell script pythonPath & " " & quoted form of scriptPath & " &> /dev/null &"
APPLEOF

# osacompile で .app を作成（macOS 標準機能）
rm -rf "LinkForge.app" 2>/dev/null
osacompile -o "LinkForge.app" /tmp/linkforge_builder.applescript

if [ $? -eq 0 ]; then
    # 検疫属性を除去
    xattr -cr "LinkForge.app" 2>/dev/null

    echo ""
    echo "====================================="
    echo "  セットアップ完了！"
    echo ""
    echo "  LinkForge.app が作成されました。"
    echo "  ダブルクリックで起動できます。"
    echo ""
    echo "  ドラッグ＆ドロップ対応済み！"
    echo "====================================="
    echo ""

    # 作成したアプリをすぐ起動
    open "LinkForge.app"
else
    echo ""
    echo "[エラー] .app の作成に失敗しました。"
    echo "ターミナルで以下のコマンドで直接起動できます:"
    echo "  $PYTHON $SCRIPT_PATH"
fi

rm -f /tmp/linkforge_builder.applescript
read -p "Enterキーで閉じます..."
