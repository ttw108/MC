# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['r2e.py'],
    pathex=['C:\\Users\\ttw\\PycharmProjects\\MC'],
    binaries=[],
    datas=[('C:\\Users\\ttw\\PycharmProjects\\MC\\venv\\lib\\site-packages\\sklearn', 'sklearn'),('C:\\Users\\ttw\\PycharmProjects\\MC\\venv\\lib\\site-packages\\matplotlib', 'matplotlib')],
    hiddenimports=['matplotlib.pyplot','seaborn',],
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
    name='r2e',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
