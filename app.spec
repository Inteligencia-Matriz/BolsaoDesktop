# -*- mode: python ; coding: utf-8 -*-
import os

spec_dir = SPECPATH
block_cipher = None

# CORRIGIDO: Adicionada a vírgula que faltava
datas_list = [
    (os.path.join(spec_dir, 'carta.html'), '.'),
    (os.path.join(spec_dir, 'style.css'), '.'),
    (os.path.join(spec_dir, 'images'), 'images'), # <-- VÍRGULA ADICIONADA AQUI
    (os.path.join(spec_dir, 'dist', 'updater.exe'), '.')
]

a = Analysis(
    [os.path.join(spec_dir, 'app.py')],
    pathex=[spec_dir],
    binaries=[],
    datas=datas_list,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GestorBolsao',
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
    icon=os.path.join(spec_dir, 'images', 'matriz.ico')
)