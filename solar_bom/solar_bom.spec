# -*- mode: python ; coding: utf-8 -*-
import sys
from version import get_version  # Import the version

block_cipher = None

a = Analysis(
    ['build_app.py'],
    pathex=[],
    binaries=[],
    datas=[
    ('data/module_templates.json', 'data'),
    ('data/tracker_templates.json', 'data'),
    ('data', 'data'),
    ('README.txt', '.'),
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=f'Solar eBOS BOM Generator v{get_version()}',  # Include version in name
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
    icon='BOM_Tool.ico'
)