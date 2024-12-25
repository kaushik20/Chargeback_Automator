# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Automator.py'],
    pathex=[],
    binaries=[],
    datas=[('Circular_Chargeback_Automator.png', '.'), ('ae7cd05d9438e3a42f955718affa1c9b.gif', '.')],
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
    name='Automator',
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
    icon=['Circular_Chargeback_Automator.png'],
)
