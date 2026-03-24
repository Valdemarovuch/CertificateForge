# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas_webview, binaries_webview, hiddenimports_webview = collect_all('webview')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries_webview,
    datas=[
        ('index.html', '.'),
        *datas_webview,
    ],
    hiddenimports=[
        *hiddenimports_webview,
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'webview.platforms.mshtml',
        'clr',
        'clr_loader',
        'pandas',
        'fitz',
        'pymupdf',
        'pypdf',
        'pypdf._reader',
        'pypdf._writer',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.lib',
        'reportlab.lib.colors',
        'reportlab.lib.pagesizes',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CertificateForge',
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
    icon=None,
)

