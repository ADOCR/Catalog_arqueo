# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

heif_datas, heif_binaries, heif_hidden = collect_all("pillow_heif")

a = Analysis(
    ["catalogador_pozos.py"],
    pathex=[],
    binaries=heif_binaries,
    datas=[
        ("assets", "assets"),
    ] + heif_datas,
    hiddenimports=heif_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib", "pytest"],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="CatalogadorPozos",
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
    icon="assets/cucharilla.ico",
    version="packaging/version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="CatalogadorPozos",
)
