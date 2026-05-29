# Windows 单文件 exe（可选，Release 打包用）

## GUI 版（推荐分发）

```powershell
pip install -e ".[gui]"
pip install pyinstaller
# 必须用 spec：打包 latex2mathml 的 unimathsymbols.txt 等数据文件
pyinstaller --noconfirm scripts/WPS-LaTeX2Equation.spec
```

或双击 `scripts\build_gui_exe.bat`。

> 不要仅用 `pyinstaller gui/app.py`，否则会报错找不到 `unimathsymbols.txt`。

**exe 闪退 / 打不开时**，先打调试版看控制台报错：

```powershell
pyinstaller --noconfirm scripts/WPS-LaTeX2Equation-debug.spec
dist\WPS-LaTeX2Equation-debug.exe
```

打包前请**关闭正在运行的** `WPS-LaTeX2Equation.exe`，否则会 Permission denied。

产物：`dist/WPS-LaTeX2Equation.exe`。需随附说明：本机需 Office 的 `MML2OMML.XSL`。

## CLI 版

```powershell
pip install pyinstaller
pyinstaller --onefile --name patent-math convert_latex_docx.py
```

用户仍需本机有 `MML2OMML.XSL`，或通过环境变量 `PATENT_MATH_MML2OMML` 指定。

将 exe 附到 GitHub Release 时，在 `.github/workflows/release.yml` 中取消 `windows-exe` job 注释。
