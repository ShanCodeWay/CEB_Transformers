# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('src/add_icon.png', 'src'), ('src/update_icon.png', 'src'), ('src/delete_icon.png', 'src'), ('src/cancel_icon.png', 'src'), ('src/terminate_icon.png', 'src'), ('src/Export_icon.png', 'src'), ('src/Export_View_icon.png', 'src'), ('src/Search_icon.png', 'src'), ('src/BG.png', 'src'), ('src/Transformers.jpg', 'src'), ('src/LOGO.png', 'src'), ('src/assets.ico', 'src')],
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
    name='main',
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
    icon=['assets.ico'],
)
