# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller配置文件
用于将简历信息提取工具打包为独立的Windows可执行文件
"""

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],  # 无需额外数据文件
    hiddenimports=[
        'openpyxl',
        'pdfplumber',
        'PyPDF2',
        'openpyxl.cell._writer',
        'PIL._tkinter_finder',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='简历信息提取工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 使用UPX压缩可执行文件
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 显示控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以在这里指定图标文件路径
)
