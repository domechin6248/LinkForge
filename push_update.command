#!/bin/bash
# =====================================================
#  楽々JC アップデート配信スクリプト
#  使い方: このファイルをダブルクリックするだけ
# =====================================================

cd "$(dirname "$0")"

echo "======================================"
echo "  楽々JC アップデート配信ツール"
echo "======================================"
echo ""

# 現在のバージョンを取得
CURRENT=$(grep 'APP_VERSION' linkforge.py | head -1 | sed "s/.*= *['\"]//;s/['\"].*//")
echo "現在のバージョン: v$CURRENT"
echo ""

# 新バージョンを入力
echo "新しいバージョン番号を入力してください (例: 2.0.1)"
echo "そのまま Enter でキャンセル:"
read NEW_VERSION

if [ -z "$NEW_VERSION" ]; then
    echo ""
    echo "キャンセルしました。"
    read -p "Enterキーで閉じる..."
    exit 0
fi

if [ "$NEW_VERSION" = "$CURRENT" ]; then
    echo ""
    echo "バージョン番号が変わっていません。中断します。"
    read -p "Enterキーで閉じる..."
    exit 1
fi

echo ""
echo "v$CURRENT → v$NEW_VERSION に更新して配信します。"
echo "よろしいですか？ (y/N)"
read CONFIRM

if [[ "$CONFIRM" != "y" && "$CONFIRM" != "Y" ]]; then
    echo "キャンセルしました。"
    read -p "Enterキーで閉じる..."
    exit 0
fi

echo ""
echo "--- linkforge.py のバージョンを更新中..."

# linkforge.py の APP_VERSION を書き換え
sed -i '' "s/APP_VERSION  = \"$CURRENT\"/APP_VERSION  = \"$NEW_VERSION\"/" linkforge.py

if grep -q "APP_VERSION  = \"$NEW_VERSION\"" linkforge.py; then
    echo "    OK: APP_VERSION を $NEW_VERSION に変更しました"
else
    echo "    ERROR: バージョン変更に失敗しました。中断します。"
    read -p "Enterキーで閉じる..."
    exit 1
fi

# version.txt を更新
echo "$NEW_VERSION" > version.txt
echo "    OK: version.txt を更新しました"

echo ""
echo "--- GitHub にプッシュ中..."

git add linkforge.py version.txt rules.csv
git commit -m "Release v$NEW_VERSION"

if git push origin main; then
    echo ""
    echo "======================================"
    echo "  配信完了！ v$NEW_VERSION"
    echo "======================================"
    echo ""
    echo "全員のアプリが次回起動時に自動でアップデートを検知します。"
    echo ""
    echo "【Mac ユーザー】"
    echo "  アプリ起動時に「アップデートあり」と表示され、"
    echo "  「はい」を押すと自動で更新されます。"
    echo ""
    echo "【Windows ユーザー】"
    echo "  アプリ起動時に通知が出るので、"
    echo "  新しい 楽々JC.exe をビルドして LINE で送ってください。"
else
    echo ""
    echo "ERROR: GitHub へのプッシュに失敗しました。"
    echo "GitHub にログインしているか確認してください。"
fi

echo ""
read -p "Enterキーで閉じる..."
