# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# SỬA: Thêm các file dữ liệu cần thiết vào đây
# Cú pháp: ('đường dẫn file nguồn', 'thư mục đích trong file exe')
# Dấu '.' nghĩa là thư mục gốc, cùng cấp với file exe.
added_files = [
    ('dark.qss', '.'),
    ('light.qss', '.'),
    ('logo_base64.txt', '.')
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,  # SỬA: Sử dụng biến added_files ở đây
    hiddenimports=[
        'pynput.keyboard._win32', 
        'pynput.mouse._win32',
        'win32process' # Thêm các thư viện có thể bị ẩn
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ACSoftSender',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',
    onefile=True,  # Thêm dòng này để tạo một file .exe duy nhất
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ACSoftSender',
)
