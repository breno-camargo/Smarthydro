# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app_icon.ico', '.'),
        ('logo_final.png', '.'),
        ('signature_logo.png', '.'),
        ('gui_logo.png', '.'),
        ('modelo_relatorio.xlsx', '.'),
        ('modelo_email.html', '.')
    ],
    hiddenimports=['babel.numbers', 'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'openpyxl.chart'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['unittest', 'test', 'pydoc', 'tkinter.test', 'xml.test', 'distutils'],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RelatorioHidrometros',
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
    icon='app_icon.ico',
)
