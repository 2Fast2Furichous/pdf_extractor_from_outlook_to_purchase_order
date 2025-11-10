# -*- mode: python ; coding: utf-8 -*-
import os
import ttkbootstrap

# Get ttkbootstrap package location
ttk_path = os.path.dirname(ttkbootstrap.__file__)

a = Analysis(
    ["pdf_extractor_app.py"],
    pathex=[],
    binaries=[],
    datas=[
        (os.path.join(ttk_path, "themes"), "ttkbootstrap/themes"),
        ("../icon.ico", "."),  # Include icon in the bundle
    ],
    hiddenimports=[
        "ttkbootstrap",
        "ttkbootstrap.themes",
        "ttkbootstrap.scrolled",
        "pandas",
        "pandas._config",
        "pandas._config.localization",
        "pandas._libs.tslibs.timedeltas",
        "pandas._libs.tslibs.np_datetime",
        "pandas._libs.tslibs.nattype",
        "pandas._libs.tslib",
        "pandas.io.formats.excel",
        "pandas.io.excel._openpyxl",
        "openpyxl",
        "openpyxl.cell",
        "openpyxl.styles",
        "win32timezone",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pandas.tests",
        "numpy.tests",
        "pytest",
        "hypothesis",
        "matplotlib",
        "scipy",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PDF_Extractor",
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
    icon="../icon.ico",
)
