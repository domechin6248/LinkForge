# -*- mode: python ; coding: utf-8 -*-
# ════════════════════════════════════════
#  楽々JC - GitHub Actions CI ビルド用 spec
#  出力名: rakuraku_jc.exe (ASCII名 - CI環境対応)
# ════════════════════════════════════════

from PyInstaller.utils.hooks import collect_all

dnd_datas, dnd_binaries, dnd_hiddenimports = collect_all('tkinterdnd2')

a = Analysis(
    ['linkforge.py'],
    pathex=[],
    binaries=dnd_binaries,
    datas=[
        ('rules.csv', '.'),
        *dnd_datas,
    ],
    hiddenimports=[
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'lxml',
        'lxml._elementpath',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fonts',
        'pptx',
        'pptx.dml.color',
        'pptx.util',
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
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
    a.binaries,
    a.datas,
    [],
    name='rakuraku_jc',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
