# Windows 单文件 exe（可选，Release 打包用）

在已验证 `python convert_latex_docx.py` 正常后执行：

```powershell
pip install pyinstaller
pyinstaller --onefile --name patent-math convert_latex_docx.py
```

产物：`dist/patent-math.exe`。用户仍需本机有 `MML2OMML.XSL`（Office 安装），或通过环境变量 `PATENT_MATH_MML2OMML` 指定。

将 exe 附到 GitHub Release `v1.0.0` 时，在 `.github/workflows/release.yml` 中取消 `windows-exe` job 注释。
