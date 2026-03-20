#!/bin/bash
# ════════════════════════════════════════════════════════════════
#  楽々JC v2.0.0  ─  Mac ビルド・デプロイ スクリプト
#  このファイルをダブルクリックして実行してください
# ════════════════════════════════════════════════════════════════

set -e
cd "$(dirname "$0")"
SCRIPT_DIR="$(pwd)"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  🍺 楽々JC v2.0.0  Mac ビルド & デプロイ"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# ── [1/6] 依存パッケージ確認 ─────────────────────────────────────
echo "[ 1/6 ] 依存パッケージを確認中..."

_pip() { pip3 install "$1" --break-system-packages -q 2>/dev/null || pip3 install "$1" -q; }

python3 -c "import docx"        2>/dev/null || { echo "  → python-docx をインストール中...";  _pip python-docx;   }
python3 -c "import openpyxl"    2>/dev/null || { echo "  → openpyxl をインストール中...";      _pip openpyxl;      }
python3 -c "from pptx import Presentation" 2>/dev/null || { echo "  → python-pptx をインストール中..."; _pip python-pptx; }
python3 -c "import pdfplumber"  2>/dev/null || { echo "  → pdfplumber をインストール中...";    _pip pdfplumber;    }
python3 -c "import tkinterdnd2" 2>/dev/null || { echo "  → tkinterdnd2 をインストール中...";   _pip tkinterdnd2;   }
python3 -c "import PyInstaller" 2>/dev/null || { echo "  → PyInstaller をインストール中...";   _pip pyinstaller;   }
echo "  ✓ 依存パッケージ OK"
echo ""

# ── [2/6] アイコン変換（iconutil で高品質 icns 生成） ────────────
echo "[ 2/6 ] アイコンを生成中..."

if [ -f "icon.iconset/icon_512x512@2x.png" ]; then
    iconutil -c icns icon.iconset -o icon.icns 2>/dev/null && \
        echo "  ✓ icon.icns（iconutil 高品質版）を生成" || \
        echo "  ✓ icon.icns（既存ファイルを使用）"
elif [ -f "icon_src.png" ] && [ ! -f "icon.icns" ]; then
    python3 build_icon.py icon_src.png
fi
echo ""

# ── [3/6] 旧ビルドをクリア ───────────────────────────────────────
echo "[ 3/6 ] 旧ビルドをクリア中..."
rm -rf dist/楽々JC dist/楽々JC.app build/楽々JC 2>/dev/null || true
echo "  ✓ クリア完了"
echo ""

# ── [4/6] PyInstaller でビルド ───────────────────────────────────
echo "[ 4/6 ] PyInstaller でビルド中（数分かかります）..."
echo ""
python3 -m PyInstaller "楽々JC.spec" --noconfirm 2>&1 | grep -E "INFO|WARNING|ERROR|完了" | head -30
echo ""

if [ ! -d "dist/楽々JC.app" ]; then
    echo "❌ ビルド失敗。ログを確認してください。"
    exit 1
fi
echo "  ✓ dist/楽々JC.app 作成完了"
echo ""

# ── [5/6] デスクトップに配置 ────────────────────────────────────
echo "[ 5/6 ] デスクトップに配置中..."
DESKTOP="$HOME/Desktop"
if [ -d "$DESKTOP" ]; then
    rm -rf "$DESKTOP/楽々JC.app" 2>/dev/null || true
    cp -r "dist/楽々JC.app" "$DESKTOP/"
    # アイコンキャッシュをリフレッシュ（真っ白アイコン対策）
    touch "$DESKTOP/楽々JC.app"
    killall Dock 2>/dev/null || true
    echo "  ✓ デスクトップ: ~/Desktop/楽々JC.app"
else
    echo "  ⚠️  デスクトップが見つかりません。dist/楽々JC.app を使用してください。"
fi
echo ""

# ── [6/6] GitHub にプッシュ ──────────────────────────────────────
echo "[ 6/6 ] GitHub にコードをプッシュ中..."
if git rev-parse --is-inside-work-tree > /dev/null 2>&1; then
    git add -A
    # 変更があるときのみコミット
    if ! git diff --cached --quiet; then
        git commit -m "v2.0.0: アイコン追加・Mac/Windows ビルド対応・コードリファクタリング

- icon_src.png / icon.icns / icon.ico を追加
- 楽々JC.spec: linkforge.py エントリーポイント修正・アイコン・メタデータ設定
- 楽々JC_Windows.spec: Windows .exe ビルド spec 新規追加
- build_mac.command / build_windows.bat: ビルド全自動化スクリプト追加
- linkforge.py: LoggedFrame・make_btn() 追加・_check_update() 集約
- DropZone: count_extensions パラメータ追加
- version.txt: 2.0.0 に更新

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
        echo "  ✓ コミット完了"
    else
        echo "  ℹ️  変更なし（コミット不要）"
    fi
    git push origin main 2>&1 && echo "  ✓ GitHub プッシュ完了" || {
        echo "  ⚠️  プッシュ失敗。GitHub Desktop から手動でプッシュしてください。"
    }
else
    echo "  ⚠️  Git リポジトリが見つかりません。GitHub Desktop から手動でプッシュしてください。"
fi
echo ""

# ── 完了 ─────────────────────────────────────────────────────────
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  ✅  すべての処理が完了しました！"
echo ""
echo "  📱 デスクトップの 楽々JC.app をダブルクリックで起動"
echo "  📦 dist/楽々JC.app をアプリケーションフォルダに移動することも可能"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "今すぐアプリを起動しますか？ (y/n)"
read -r OPEN_NOW
if [ "$OPEN_NOW" = "y" ] || [ "$OPEN_NOW" = "Y" ]; then
    open "$DESKTOP/楽々JC.app" 2>/dev/null || open "dist/楽々JC.app"
fi
