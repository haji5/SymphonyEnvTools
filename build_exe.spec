# -*- mode: python ; coding: utf-8 -*-
import sys
import os

block_cipher = None

# Define the main script
main_script = 'centralUI.py'

# All Python files to include as modules
py_modules = [
    'centralUI',
    'PRIZM_Report_GUI',
    'PrizmExcelMerger',
    'PRIZM_data_grabber',
    'Subdivision_data_grabber',
    'Transformation_3'
]

# Hidden imports for packages that might not be detected automatically
hiddenimports = [
    'pandas',
    'selenium',
    'selenium.webdriver',
    'selenium.webdriver.chrome',
    'selenium.webdriver.chrome.options',
    'selenium.webdriver.common.by',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
    'selenium.webdriver.support',
    'win32com.client',
    'pythoncom',
    'openpyxl',
    'xlsxwriter',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'threading',
    'subprocess',
    'glob',
    'math',
    'importlib.util',
    'time'
]

# Data files to include (if any)
datas = []

# Binaries to include
binaries = []

a = Analysis(
    [main_script],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='EnvironicsDataTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI application
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None  # You can add an icon file path here if you have one
)
