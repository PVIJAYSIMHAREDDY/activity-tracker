# PyInstaller spec — used for Windows and macOS builds.
# Usage: pyinstaller ActivityTracker.spec

import sys, os
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static',    'static'),
        ('icon.png',  '.'),
    ],
    hiddenimports=[
        # Flask / Werkzeug
        'flask', 'flask.templating', 'jinja2', 'jinja2.ext',
        'werkzeug', 'werkzeug.serving', 'werkzeug.routing',
        'click', 'itsdangerous',
        # openpyxl
        'openpyxl', 'openpyxl.styles', 'openpyxl.chart',
        'openpyxl.chart.series', 'openpyxl.utils',
        'openpyxl.formatting.rule', 'openpyxl.worksheet.datavalidation',
        'et_xmlfile',
        # pywebview — platform-specific backends
        'webview',
        'webview.platforms.winforms',   # Windows (Edge WebView2)
        'webview.platforms.edgechromium',
        'webview.platforms.cocoa',      # macOS
        'webview.platforms.gtk',        # Linux (fallback)
        'webview.http',
        'webview.util',
        'webview.event',
        'webview.menu',
        'webview.dom',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'scipy', 'PIL'],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Windows / Linux: single-dir build ──────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ActivityTracker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='icon.ico' if sys.platform == 'win32' else (
         'icon.icns' if sys.platform == 'darwin' else 'icon.png'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ActivityTracker',
)

# ── macOS: wrap in .app bundle ─────────────────────────────────────────────
if sys.platform == 'darwin':
    app_bundle = BUNDLE(
        coll,
        name='ActivityTracker.app',
        icon='icon.icns',
        bundle_identifier='com.activitytracker.app',
        info_plist={
            'CFBundleName': 'Daily Activity Tracker',
            'CFBundleDisplayName': 'Daily Activity Tracker',
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'NSHighResolutionCapable': True,
            'NSRequiresAquaSystemAppearance': False,
        },
    )
