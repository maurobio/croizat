# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['Croizat3.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('C:\\Users\\mauro\\My Documents\\Projetos\\Croizat2\\images', 'images'),
        ('C:\\Python314\\Lib\\site-packages\\mpl_toolkits\\basemap_data', 'mpl_toolkits\\basemap_data'),
    ],
    hiddenimports=['mpl_toolkits.basemap_data'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['hook-basemap.py'],
    excludes=['FixTk', 'tcl', 'tk', '_tkinter', 'tkinter', 'Tkinter', 'wx', 'scipy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Croizat3',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Croizat3',
    icon='C:\\Users\\mauro\\My Documents\\Projetos\\Croizat2\\images\\croizat.ico'
)