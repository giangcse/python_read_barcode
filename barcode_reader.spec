# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['barcode_reader.py'],
    pathex=[],
    binaries=[],
    datas=[('libiconv.dll', 'pyzbar'), ('libzbar-64.dll', 'pyzbar')],
    hiddenimports=['pyzbar', 'opencv-python-headless', 'pyperclip', 'pyqt5'],
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
    name='barcode_reader',
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
    icon=['barcode-scan.ico'],
)
