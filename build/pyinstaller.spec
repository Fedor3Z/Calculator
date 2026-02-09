# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.building.datastruct import Tree

block_cipher = None

hiddenimports = collect_submodules('scipy') + collect_submodules('matplotlib')

# Bundle all UI assets (json maps, formulas, Excel template, etc.)
assets_tree = Tree('assets', prefix='assets')

a = Analysis(
    ['app/main.py'],
    pathex=['.'],
    binaries=[],
    datas=[assets_tree],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='KinematicsCalc',
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='KinematicsCalc',
)
