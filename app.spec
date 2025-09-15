# -*- mode: python ; coding: utf-8 -*-
import os

# Usa a variável SPECPATH, fornecida pelo PyInstaller, 
# para encontrar o diretório onde o arquivo .spec está.
spec_dir = SPECPATH

block_cipher = None

# Define a lista de arquivos de dados que serão incluídos no executável
# O formato é ('caminho/completo/do/arquivo', 'pasta_de_destino_no_exe')
datas_list = [
    (os.path.join(spec_dir, 'carta.html'), '.'),
    (os.path.join(spec_dir, 'style.css'), '.'),
    (os.path.join(spec_dir, 'images'), 'images')
    (os.path.join(spec_dir, 'dist', 'updater.exe'), '.')
]

a = Analysis(
    [os.path.join(spec_dir, 'app.py')], # Usa o caminho absoluto para o script principal
    pathex=[spec_dir], # Adiciona a pasta do projeto ao path de busca
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
    name='GestorBolsao', # Nome do arquivo .exe final
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Garante que a janela de console não apareça
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(spec_dir, 'images', 'matriz.ico')  # Define o ícone do programa
)