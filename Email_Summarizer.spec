# -*- mode: python ; coding: utf-8 -*-
import os
import dash
import dash_bootstrap_components

dash_dir = os.path.dirname(dash.__file__)
dbc_dir  = os.path.dirname(dash_bootstrap_components.__file__)

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Dash 내장 에셋 (JS/CSS)
        (os.path.join(dash_dir, 'dash-renderer'),  'dash/dash-renderer'),
        (os.path.join(dash_dir, 'dcc'),             'dash/dcc'),
        (os.path.join(dash_dir, 'html'),            'dash/html'),
        (os.path.join(dash_dir, 'dash_table'),      'dash/dash_table'),
        (os.path.join(dash_dir, 'favicon.ico'),     'dash'),
        # 프로젝트 파일
        ('assets',                                   'assets'),
        ('outlook_manager.py',                       '.'),
        ('ai_processor.py',                          '.'),
        ('notion_sync.py',                           '.'),
    ],
    hiddenimports=[
        'dash',
        'dash_bootstrap_components',
        'diskcache',
        'flask',
        'flask_compress',
        'win32com.client',
        'win32com',
        'pythoncom',
        'pywintypes',
        'google.genai',
        'dotenv',
        'notion_sync',
        'ai_processor',
        'outlook_manager',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'pandas', 'scipy'],
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
    name='Email_Summarizer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 콘솔창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/favicon.ico',
)
