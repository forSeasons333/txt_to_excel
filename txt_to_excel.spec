# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['txt_to_excel.py'],
    pathex=[],
    binaries=[],
    datas=[('icons\\folder_icon.png', 'icons'), ('icons\\start_icon.png', 'icons'), ('icons\\tte.png', 'icons')],
    hiddenimports=[],
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
    name='txt_to_excel',
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
    icon=['icons\\tte.ico'],
)
