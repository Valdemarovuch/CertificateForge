# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_all

block_cipher = None

openpyxl_all = collect_all('openpyxl')
webview_datas = collect_data_files('webview')
webview_binaries = collect_data_files('webview', includes=['*.dll'])

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=openpyxl_all[1],
    datas=[
        ('index.html', '.'),
        ('assets', 'assets'),
    ] + openpyxl_all[0] + webview_datas,
    hiddenimports=openpyxl_all[2] + collect_submodules('webview') + [
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.styles.stylesheet',
        'openpyxl.drawing.image',
        'et_xmlfile',
        'xlrd',
        'reportlab.graphics',
        'reportlab.pdfbase._fontdata_enc_winansi',
        'reportlab.pdfbase._fontdata_enc_macroman',
        'reportlab.pdfbase._fontdata_enc_standard',
        'reportlab.pdfbase._fontdata_enc_symbol',
        'reportlab.pdfbase._fontdata_enc_zapfdingbats',
        'reportlab.pdfbase._fontdata_widths_courier',
        'reportlab.pdfbase._fontdata_widths_courierbold',
        'reportlab.pdfbase._fontdata_widths_courieroblique',
        'reportlab.pdfbase._fontdata_widths_courierboldoblique',
        'reportlab.pdfbase._fontdata_widths_helvetica',
        'reportlab.pdfbase._fontdata_widths_helveticabold',
        'reportlab.pdfbase._fontdata_widths_helveticaoblique',
        'reportlab.pdfbase._fontdata_widths_helveticaboldoblique',
        'reportlab.pdfbase._fontdata_widths_timesroman',
        'reportlab.pdfbase._fontdata_widths_timesbold',
        'reportlab.pdfbase._fontdata_widths_timesitalic',
        'reportlab.pdfbase._fontdata_widths_timesbolditalic',
        'reportlab.pdfbase._fontdata_widths_symbol',
        'reportlab.pdfbase._fontdata_widths_zapfdingbats',
    ],
    excludes=[
        'tkinter',
        'tkinter.ttk',
        '_tkinter',
        'matplotlib',
        'scipy',
        'numpy.testing',
        'unittest',
        'test',
        'tests',
        'pydoc',
        'doctest',
        'difflib',
        'argparse',
        'optparse',
        'bdb',
        'pdb',
        'profile',
        'pstats',
        'timeit',
        'trace',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CertificateForge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'Microsoft.Web.WebView2.Core.dll',
        'WebView2Loader.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
