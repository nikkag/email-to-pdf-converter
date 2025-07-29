# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path

# Get the Playwright browsers path
playwright_browsers_path = None
if os.name == 'nt':  # Windows
    playwright_browsers_path = os.path.join(os.environ.get('USERPROFILE', ''), 'AppData', 'Local', 'ms-playwright')
else:  # Unix-like
    playwright_browsers_path = os.path.join(os.environ.get('HOME', ''), '.cache', 'ms-playwright')

# Add Playwright browsers to the bundle
datas = []
if playwright_browsers_path and os.path.exists(playwright_browsers_path):
    for browser_dir in os.listdir(playwright_browsers_path):
        browser_path = os.path.join(playwright_browsers_path, browser_dir)
        if os.path.isdir(browser_path):
            datas.append((browser_path, browser_dir))

a = Analysis(
    ['eml_to_pdf_converter.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'email',
        'email.parser',
        'email.policy',
        'email.utils',
        'extract_msg',
        'beautifulsoup4',
        'fpdf',
        'playwright',
        'playwright.async_api',
        'playwright.sync_api',
        'playwright._impl._browser_context',
        'playwright._impl._browser',
        'playwright._impl._page',
        'playwright._impl._connection',
        'playwright._impl._transport',
        'playwright._impl._driver',
        'playwright._impl._browser_type',
        'playwright._impl._playwright',
        'playwright._impl._api_types',
        'playwright._impl._api_structures',
        'playwright._impl._errors',
        'playwright._impl._object_factory',
        'playwright._impl._network',
        'playwright._impl._cdp_session',
        'playwright._impl._frame',
        'playwright._impl._element_handle',
        'playwright._impl._js_handle',
        'playwright._impl._worker',
        'playwright._impl._route',
        'playwright._impl._request',
        'playwright._impl._response',
        'playwright._impl._websocket',
        'playwright._impl._video',
        'playwright._impl._tracing',
        'playwright._impl._accessibility',
        'playwright._impl._locator',
        'playwright._impl._expect',
        'asyncio',
        'html',
        'unicodedata',
        'warnings',
        'datetime',
        'pathlib',
        're',
        'sys',
        'os',
    ],
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=['hooks/runtime-hook-playwright.py'],
    excludes=[],
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
    name='EmailToPDFConverter',
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
)