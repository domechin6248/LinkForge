#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
楽々JC アイコン変換スクリプト
PNG → icon.icns (Mac) + icon.ico (Windows)

使い方:
  python3 build_icon.py <icon画像.png>

例:
  python3 build_icon.py icon_src.png
"""

import sys
import os
import shutil
import subprocess
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("Pillowをインストールします...")
    os.system("pip3 install pillow --break-system-packages -q")
    from PIL import Image


def make_ico(src: Path, dst: Path):
    """PNG → .ico (Windows 用: 多サイズ埋め込み)"""
    img = Image.open(src).convert("RGBA")
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48),
             (64, 64), (128, 128), (256, 256)]
    imgs = []
    for s in sizes:
        resized = img.resize(s, Image.LANCZOS)
        imgs.append(resized)
    imgs[0].save(dst, format="ICO", sizes=[(i.width, i.height) for i in imgs],
                 append_images=imgs[1:])
    print(f"  ✓ {dst.name} 作成完了")


def make_iconset(src: Path, iconset_dir: Path):
    """PNG → .iconset ディレクトリ（Mac用: iconutil の入力）"""
    img = Image.open(src).convert("RGBA")
    iconset_dir.mkdir(exist_ok=True)
    specs = [
        ("icon_16x16.png",      16),
        ("icon_16x16@2x.png",   32),
        ("icon_32x32.png",      32),
        ("icon_32x32@2x.png",   64),
        ("icon_64x64.png",      64),
        ("icon_64x64@2x.png",  128),
        ("icon_128x128.png",   128),
        ("icon_128x128@2x.png",256),
        ("icon_256x256.png",   256),
        ("icon_256x256@2x.png",512),
        ("icon_512x512.png",   512),
        ("icon_512x512@2x.png",1024),
    ]
    for name, size in specs:
        resized = img.resize((size, size), Image.LANCZOS)
        resized.save(iconset_dir / name)


def make_icns_mac(src: Path, dst: Path):
    """Mac 上でのみ動作: iconutil を使って .icns 生成"""
    iconset_dir = dst.parent / "icon.iconset"
    make_iconset(src, iconset_dir)
    result = subprocess.run(
        ["iconutil", "-c", "icns", str(iconset_dir), "-o", str(dst)],
        capture_output=True, text=True
    )
    shutil.rmtree(str(iconset_dir))
    if result.returncode == 0:
        print(f"  ✓ {dst.name} 作成完了")
    else:
        print(f"  ✗ icns 変換失敗: {result.stderr}")
        raise RuntimeError(result.stderr)


def make_icns_pillow(src: Path, dst: Path):
    """Linux/クロスプラットフォーム用: Pillow で .icns を直接生成"""
    img = Image.open(src).convert("RGBA")
    # Pillow の ICNS 対応サイズ
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    images = []
    for s in sizes:
        images.append(img.resize((s, s), Image.LANCZOS))
    images[0].save(dst, format="ICNS", append_images=images[1:])
    print(f"  ✓ {dst.name} 作成完了 (Pillow)")


def convert_icon(src_path: str):
    src = Path(src_path)
    if not src.exists():
        print(f"エラー: ファイルが見つかりません: {src}")
        sys.exit(1)

    out_dir = src.parent
    ico_path  = out_dir / "icon.ico"
    icns_path = out_dir / "icon.icns"

    print(f"入力: {src}")
    print("変換中...")

    # --- Windows 用 .ico ---
    make_ico(src, ico_path)

    # --- Mac 用 .icns ---
    import platform
    if platform.system() == "Darwin":
        try:
            make_icns_mac(src, icns_path)
        except Exception:
            print("  iconutil 失敗。Pillow で再試行...")
            make_icns_pillow(src, icns_path)
    else:
        make_icns_pillow(src, icns_path)

    print(f"\n出力先: {out_dir}")
    print(f"  icon.ico  → Windows ビルドで使用")
    print(f"  icon.icns → Mac ビルドで使用")
    print("\n次のステップ:")
    print("  【Mac】  ./build_mac.command")
    print("  【Win】  build_windows.bat")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        # 同じフォルダに icon_src.png があれば自動検出
        for name in ["icon_src.png", "icon.png", "app_icon.png"]:
            if Path(name).exists():
                print(f"→ {name} を自動検出。変換を開始します。")
                convert_icon(name)
                sys.exit(0)
        print("使い方: python3 build_icon.py <画像ファイル.png>")
        sys.exit(1)
    convert_icon(sys.argv[1])
