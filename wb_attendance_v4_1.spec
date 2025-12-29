# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

# 빌드 산출물(영문 파일명) + DB 포함
datas = [('favicon.ico', '.'), ('leave_attendance.db', '.')]
binaries = []

hiddenimports = [
    'pandas',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.skiplist',
    'pandas._libs.algos',
    'pandas._libs.window.aggregations',
    'openpyxl',
    'openpyxl.cell._writer',
    'openpyxl.workbook.external_link.external',
    'openpyxl.packaging.workbook',
    'sqlite3',
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
]
hiddenimports += collect_submodules('pandas')
hiddenimports += collect_submodules('openpyxl')

tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PySide6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['Attendance and Leave Management Program.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='WB_Attendance_Manager_v4_1',
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
    icon=['favicon.ico'],
)


