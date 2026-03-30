# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('data.py', '.'),
        ('INVENTRA.json', '.'),
        ('logo_app.ico', '.'),
        ('logo_app.png', '.'),
        ('logo_trsp.png', '.'),
        ('tutorial_mapping.png', '.'),
        ('tutorial_preview.png', '.'),
        ('tutorial_reset.png', '.'),
        ('tutorial_upload.png', '.'),
        ('rumus_slide', '.'),
        ('icon_drp.png', '.'),
        ('icon_settingan.png', '.')
    ],
    hiddenimports=[
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.skiplist',
        'numpy',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.cell',
        'xlrd',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'tzdata',
        'pywinstyles',
        'urllib.request',
        'urllib.parse',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyd = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyd,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='INVENTRA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='logo_app.ico'
)
