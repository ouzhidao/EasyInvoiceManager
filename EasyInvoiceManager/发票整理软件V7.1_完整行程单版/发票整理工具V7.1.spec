# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main_v6.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['PyQt5.sip', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.QtPrintSupport', 'PIL', 'PIL._imagingtk', 'PIL._tkinter_finder', 'requests', 'openpyxl', 'openpyxl.cell._writer', 'fitz', 'fitz.fitz', 'dateutil', 'dateutil.tz', 'dateutil.parser', 'gui.main_window', 'gui.folder_dialog', 'gui.print_dialog', 'gui.password_dialog', 'gui.main_window_v2', 'utils.helpers'],
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
    name='发票整理工具V7.1',
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
)
