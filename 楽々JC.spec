# -*- mode: python ; coding: utf-8 -*-
# ════════════════════════════════════════
#  楽々JC v2.0.0  ─  Mac ビルド用 spec
#  実行: pyinstaller 楽々JC.spec
# ════════════════════════════════════════

from PyInstaller.utils.hooks import collect_all

# tkinterdnd2 はバイナリが必要なため collect_all で丸ごと取り込む
dnd_datas, dnd_binaries, dnd_hiddenimports = collect_all('tkinterdnd2')

a = Analysis(
    ['linkforge.py'],
    pathex=[],
    binaries=dnd_binaries,
    datas=[
        ('rules.csv', '.'),   # デフォルトルールを同梱
        *dnd_datas,
    ],
    hiddenimports=[
        # python-docx
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'lxml',
        'lxml._elementpath',
        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fonts',
        # python-pptx
        'pptx',
        'pptx.dml.color',
        'pptx.util',
        # pdfplumber
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        # tkinterdnd2
        *dnd_hiddenimports,
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='楽々JC',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.icns',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='楽々JC',
)
app = BUNDLE(
    coll,
    name='楽々JC.app',
    icon='icon.icns',
    bundle_identifier='com.rakuraku.jc',
    info_plist={
        'CFBundleName':               '楽々JC',
        'CFBundleDisplayName':        '楽々JC',
        'CFBundleVersion':            '2.0.0',
        'CFBundleShortVersionString': '2.0.0',
        'NSHighResolutionCapable':    True,
        'LSMinimumSystemVersion':     '10.13.0',
        'NSHumanReadableCopyright':   '© 2025 楽々JC',
    },
)
