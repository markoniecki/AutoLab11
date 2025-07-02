# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('ROD_11_/*.lbx', 'ROD_11_'),
        ('ROD_11_/*.py',  'ROD_11_'),

        ('tamuryn/bazolec.lbx', 'tamuryn'),
        ('tamuryn/bazolecROD.lbx', 'tamuryn'),
        ('tamuryn/bazolecSL.lbx', 'tamuryn'),
        ('tamuryn/bazolecSH.lbx', 'tamuryn'),
        ('tamuryn/bazolecSH_SL.lbx', 'tamuryn'),
        ('tamuryn/class15.py', 'tamuryn'),
        #('tamuryn/class15.py', 'tamuryn'),

        ('utils/odczyt_tabela_material_NewSheet.py', 'utils')

        # Dodaj inne pliki z utils/, jeśli są potrzebne
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[]
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AutoLabApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True  # Zmień na False, jeśli robisz GUI i nie chcesz konsoli
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='AutoLabApp'
)
