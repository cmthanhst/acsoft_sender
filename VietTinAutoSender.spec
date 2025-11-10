# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],    
    datas=[('dark.qss', '.'), ('light.qss', '.'), ('logo_base64.txt', '.')],
    hiddenimports=[
        'PySide6',
        'pandas',
        'pynput',
        'pynput.keyboard._win32', # Hook cho pynput keyboard trên Windows
        'pynput.mouse._win32',    # Hook cho pynput mouse trên Windows
        'pywin32', # Đảm bảo tất cả các module win32 được bao gồm
        'win32gui',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,

)
pyz = PYZ(a.pure)

coll = COLLECT(
    EXE(
        pyz,
        a.scripts,
        [],
        name='VietTinAutoSender',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon='logo.ico',
    ),
    a.binaries,
    a.datas,
    name='VietTinAutoSender',
)
