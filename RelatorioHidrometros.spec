# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.json', '.'),
        ('app_icon.ico', '.'),
        ('header_logo.png', '.'),
        ('logo.jpeg', '.'),
        ('logo_final.png', '.'),
        ('signature_logo.png', '.'),
        ('gui_logo.png', '.'),
        ('modelo_relatorio.xlsx', '.'),
        ('modelo_email.html', '.')
    ],
    hiddenimports=['babel.numbers', 'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'email', 'smtplib'],
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
