# -*- mode: python ; coding: utf-8 -*-
import os
import sys

a = Analysis(
    ['pdf_extractor_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('web', 'web'),
        ('eel_files', 'eel'),  # From the diagnostic script
    ],
    hiddenimports=[
        # Eel and web framework
        'eel',
        'bottle',
        'bottle_websocket',
        'gevent',
        'gevent.socket',
        'gevent.monkey',
        'geventwebsocket',
        
        # PDF processing
        'pdfplumber',
        'pdfplumber.utils',
        'pdfminer',
        'pdfminer.six',
        
        # Pandas and data processing
        'pandas',
        'pandas._config',
        'pandas._config.localization',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslib',
        'pandas.io.formats.excel',
        'pandas.io.excel._openpyxl',
        
        # Excel handling
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.utils.dataframe',
        
        # Windows/Outlook integration - CRITICAL
        'win32com',
        'win32com.client',
        'win32com.server',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32con',
        'win32timezone',  # THIS IS THE MISSING MODULE
        
        # Additional commonly needed modules
        'numpy',
        'dateutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pandas.tests',
        'numpy.tests',
        'pytest',
        'hypothesis',
        'matplotlib',
        'scipy',
        'notebook',
        'ipython',
        'jupyter',
        'setuptools',
        'pip',
        'wheel',
        'lxml',
        'html5lib',
        'jinja2',
        'sqlalchemy',
        'pyarrow',
        'numba',
        'bottleneck',
        'numexpr',
        'tables',
        'xlrd',
        'xlwt',
        'odfpy',
        'pyxlsb',
        'bs4',
        'beautifulsoup4',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF_Extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'python*.dll', 'pywintypes*.dll'],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
