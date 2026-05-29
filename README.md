# WPS-LaTex2Equation-tool

[![Release](https://img.shields.io/github/v/release/cgflag/WPS-LaTex2Equation-tool)](https://github.com/cgflag/WPS-LaTex2Equation-tool/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

**专为 WPS / Word：把 AI 写进 docx 的 `$...$`、`$$...$$` 批量转成原生可编辑公式。不启 Word、不用 MathType、不依赖 COM 宏；块级公式自动居中并编号 (1)(2)。**

> 一句话：**不依赖 Word COM / OMath API，WPS 可直接打开；专为 AI 输出的 `$...$` 批量转原生公式。**

[English](#english) · [示例对比](examples/) · [更新日志](CHANGELOG.md)

<!-- 录屏 GIF 放到 docs/assets/demo.gif 后取消注释 -->
<!-- ![演示](docs/assets/demo.gif) -->

| 转换前 | 转换后 |
|--------|--------|
| [`examples/demo_before.docx`](examples/demo_before.docx) | [`examples/demo_after.docx`](examples/demo_after.docx) |

用 **WPS 文字** 或 Word 打开上表两个文件即可对比效果。

---

## 适合谁用

- 用 ChatGPT / Claude 等写**专利说明书、论文**，导出 docx 后公式仍是 `$E=mc^2$` 纯文本
- 日常用 **WPS 文字**，不想装 MathType，VBA 宏又经常报错
- 需要块级公式 **居中 + (1)(2)(3)… 编号**（制表位实现，符合国内专利排版习惯）

---

## 解决什么问题

| 痛点 | 本工具 |
|------|--------|
| docx 里公式是 LaTeX 字符串，无法直接当公式编辑 | 批量转为 **WPS/Word 原生 OMML 公式** |
| MathType 收费；WPS 上 VBA + `OMaths` 不稳定 | **无需 MathType**；运行时 **不调用 Word COM** |
| 独立段公式要居中并编号 | 识别 `$$...$$` / 整段 `$...$`，**制表位居中 + (n)** |
| 怕转换失败毁掉整篇文档 | 失败处 **保留原 `$...$` 文本** |

**原理**：直接改写 docx 内的 `document.xml`（LaTeX → MathML → OMML），**不启动 Word 进程**。生成文件用 WPS 打开即可双击编辑公式。

---

## 快速开始

### 环境要求

| 项目 | 说明 |
|------|------|
| Python | 3.10 及以上 |
| `MML2OMML.XSL` | 来自 **Microsoft Office** 安装目录（仅用于格式转换，**不要求日常用 Word 排版**） |
| WPS | 用于查看/编辑输出 docx（推荐） |

常见 XSL 路径（Windows）：

```
C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
```

### 安装

```bash
git clone https://github.com/cgflag/WPS-LaTex2Equation-tool.git
cd WPS-LaTex2Equation-tool
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

也可：`pip install -r requirements.txt`（不含命令行别名）。

### 图形界面（GUI，v1.1）

```bash
pip install -e ".[gui]"
wps-latex2equation-gui
# 或
python -m gui
# Windows 双击
scripts\run_gui.bat
```

- 选择或 **拖拽** docx 到窗口/拖放区 → 点击 **开始转换** → 默认另存为 `原名_公式版.docx`
- 启动时检测 `MML2OMML.XSL`，未找到则 **禁止转换** 并提示
- 显示成功/失败数量，失败公式列表；完成后可 **打开输出文件夹**

### 一行命令（推荐先试示例）

```bash
python convert_latex_docx.py examples/demo_before.docx examples/demo_after.docx
```

安装 `-e` 后可用短命令：

```bash
patent-math 你的文档.docx 输出.docx
# 别名：docx-latex-math
```

### 常用参数

```bash
# 默认输出：input_converted.docx
python convert_latex_docx.py 说明书.docx

# 指定字号（小四 = 12pt）
python convert_latex_docx.py 说明书.docx 说明书_公式版.docx --size 小四

# 块级公式制表位：按页面宽度自适应（默认）
python convert_latex_docx.py 说明书.docx out.docx --tab-mode page

# 按字符数设制表位（兼容旧宏习惯）
python convert_latex_docx.py 说明书.docx out.docx --tab-mode chars --tab-center-chars 21 --tab-right-chars 44

# 指定 XSL（仅 WPS、未装 Office 时若另有 XSL 副本）
set PATENT_MATH_MML2OMML=C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
python convert_latex_docx.py 说明书.docx out.docx --xsl "%PATENT_MATH_MML2OMML%"
```

### 公式写法

| 写法 | 效果 |
|------|------|
| 句中 `$E=mc^2$` | 行内公式 |
| 整段只有 `$$...$$` 或单独一行 `$...$` | 块级：居中 + 右侧 `(1)(2)...` |
| 某条公式解析失败 | 保留 `$...$` 原文，其余照常转换 |

---

## 开源范围

| 内容 | 说明 |
|------|------|
| ✅ **MIT 开源** | 单文件 CLI、LaTeX→OMML、行内/块级识别、制表位编号 |
| 🔒 未包含（可 Pro / 闭源） | 批量文件夹、GUI、专利模板 preset、题注/目录一体化 |

---

## 项目结构

```
WPS-LaTex2Equation-tool/
├── convert_latex_docx.py    # 主程序
├── pyproject.toml
├── requirements.txt
├── examples/                # 转换前后对比
├── tests/
├── .github/workflows/       # CI
├── scripts/build_examples.py
├── docs/xiaohongshu/        # 推广文案（可选）
└── legacy/                  # 旧版 WPS VBA 宏
```

重新生成示例：`python scripts/build_examples.py`

---

## 开发

```bash
pip install -e ".[dev]"
pytest -q
```

---

## 已知限制

- 仅处理正文 `word/document.xml`（**不含**页眉页脚、文本框、表格内文字）
- LaTeX 命令集受 `latex2mathml` 限制，极复杂宏可能失败
- `MML2OMML.XSL` 版权归 Microsoft，**不能随本仓库分发**，需本机 Office 或自行指定路径

---

## 许可证

[MIT](LICENSE) · Copyright (c) 2026 [cgflag](https://github.com/cgflag)

---

<a id="english"></a>

## English

**WPS-LaTex2Equation-tool** — Batch convert `$...$` / `$$...$$` LaTeX in `.docx` to native **OMML** equations. No MathType, no Word COM at runtime. Display equations: centered with `(1)(2)...` tab numbering. Output opens and edits in **WPS Office** and Microsoft Word.

```bash
git clone https://github.com/cgflag/WPS-LaTex2Equation-tool.git
cd WPS-LaTex2Equation-tool && pip install -e .
python convert_latex_docx.py examples/demo_before.docx examples/demo_after.docx
```

Requires `MML2OMML.XSL` from Microsoft Office (`--xsl` or env `PATENT_MATH_MML2OMML`).

License: [MIT](LICENSE)
