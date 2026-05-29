# 调试版：带控制台，便于查看 exe 启动报错
# pyinstaller --noconfirm scripts/WPS-LaTeX2Equation-debug.spec

import os

from PyInstaller.utils.hooks import collect_all

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
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="WPS-LaTeX2Equation-debug",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
