# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['document_analyzer_cli.py'],
    pathex=[],
    binaries=[],
    datas=[('section_words_form.html', '.'), ('formatting_rules.json', '.')],
    hiddenimports=[],
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
    name='document_analyzer_cli',
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

import shutil
import os
spec_dir = os.path.dirname(os.path.abspath(SPEC))
dist_dir = os.path.join(spec_dir, 'dist')
if os.path.exists(dist_dir):
    for filename in ['section_words_form.html', 'formatting_rules.json', 'uk_dict.txt', 'us_dict.txt']:
        src = os.path.join(spec_dir, filename)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(dist_dir, filename))
            print(f"Copied {filename} to {dist_dir}")

