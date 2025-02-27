# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['excel_formula_assistant.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.')],
    hiddenimports=[
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'PyQt6.sip',
        'PyQt6.QtPrintSupport'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name='excel_formula_assistant',
    debug=True,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=True,
    icon=['icon.ico'],
)