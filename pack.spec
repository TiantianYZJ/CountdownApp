# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['CountdownApp.py'], 
    pathex=[],  
    binaries=[], 
    datas=[
        ('logo.ico', '.'), 
        ('OIAPI.png', '.'), 
    ],
    hiddenimports=[], 
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
    name='CountdownApp', 
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, 
    icon="logo.ico",
    uac_admin=False
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='CountdownApp' 
)