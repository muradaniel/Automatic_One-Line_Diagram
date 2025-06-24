# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('matplotlib') + collect_data_files('seaborn')
datas += [('raio.png', 'raio.png', 'DATA')]


a = Analysis(
    ['Novo_diagrama.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'scipy.stats._distn_infrastructure',
        'scipy._lib.messagestream',
        'xlwings._xlplatform'
],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Novo_diagrama',
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
