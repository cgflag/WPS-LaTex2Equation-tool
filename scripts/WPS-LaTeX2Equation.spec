# PyInstaller spec — bundle latex2mathml data (unimathsymbols.txt) and GUI deps.
# Build from repo root: pyinstaller --noconfirm scripts/WPS-LaTeX2Equation.spec

import os

from PyInstaller.utils.hooks import collect_all

block_cipher = None

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
APP_SCRIPT = os.path.join(ROOT, "gui", "app.py")

latex2mathml_datas, latex2mathml_binaries, latex2mathml_hidden = collect_all("latex2mathml")

a = Analysis(
    [APP_SCRIPT],
    pathex=[ROOT],
    binaries=latex2mathml_binaries,
    datas=latex2mathml_datas,
    hiddenimports=latex2mathml_hidden + ["windnd", "gui.dnd_win", "customtkinter"],
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
    name="WPS-LaTeX2Equation",
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
