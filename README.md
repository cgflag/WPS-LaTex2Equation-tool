# WPS-LaTex2Equation-tool

[![Release](https://img.shields.io/github/v/release/cgflag/WPS-LaTex2Equation-tool)](https://github.com/cgflag/WPS-LaTex2Equation-tool/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

**专为 WPS / Word：把 AI 写进 docx 的 `$...$`、`$$...$$` 批量转成原生可编辑公式。不启 Word、不用 MathType、不依赖 COM 宏；块级公式自动居中并编号 (1)(2)。**

> 一句话：**不依赖 Word COM / OMath API，WPS 可直接打开；专为 AI 输出的 `$...$` 批量转原生公式。**

[English](#english) · [示例对比](examples/) · [更新日志](CHANGELOG.md)

| 转换前 | 转换后 |
|--------|--------|
| [`examples/demo_before.docx`](examples/demo_before.docx) | [`examples/demo_after.docx`](examples/demo_after.docx) |

用 **WPS 文字** 或 Word 打开上表两个脱敏示例即可对比效果。

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

### 图形界面（GUI）

```bash
pip install -e ".[gui]"
wps-latex2equation-gui
# 或
python -m gui
```

- 选择或 **拖拽** docx 到拖放区 → **开始转换** → 默认另存为 `原名_公式版.docx`
- 启动时检测 `MML2OMML.XSL`，未找到则禁止转换并提示
- 显示成功/失败数量；完成后可打开输出文件夹

### 命令行

```bash
python convert_latex_docx.py examples/demo_before.docx examples/demo_after.docx
```

安装 `-e` 后：

```bash
patent-math 你的文档.docx 输出.docx
```

### 常用参数

```bash
python convert_latex_docx.py 说明书.docx 说明书_公式版.docx --size 小四
python convert_latex_docx.py 说明书.docx out.docx --tab-mode page
set PATENT_MATH_MML2OMML=C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
python convert_latex_docx.py 说明书.docx out.docx --xsl "%PATENT_MATH_MML2OMML%"
```

### 公式写法

| 写法 | 效果 |
|------|------|
| 句中 `$E=mc^2$` | 行内公式 |
| 整段只有 `$$...$$` 或单独一行 `$...$` | 块级：居中 + 右侧 `(1)(2)...` |
| 某条公式解析失败 | 保留 `$...$` 原文 |

---

## 开源范围

| 内容 | 说明 |
|------|------|
| ✅ **MIT 开源** | CLI、GUI、LaTeX→OMML、行内/块级识别、制表位编号 |
| 🔒 未包含 | 批量文件夹、专利模板 preset、题注/目录一体化 |

---

## 项目结构

```
WPS-LaTex2Equation-tool/
├── convert_latex_docx.py
├── gui/
├── pyproject.toml
├── requirements.txt
├── examples/          # 脱敏演示 docx
├── tests/
└── .github/workflows/
```

---

## 开发

```bash
pip install -e ".[dev]"
pytest -q
```

---

## 已知限制

- 仅处理正文 `word/document.xml`（不含页眉页脚、文本框、表格）
- LaTeX 子集受 `latex2mathml` 限制
- `MML2OMML.XSL` 需本机 Office，不能随仓库分发

---

## 许可证

[MIT](LICENSE) · Copyright (c) 2026 [cgflag](https://github.com/cgflag)

---

<a id="english"></a>

## English

Batch convert `$...$` / `$$...$$` LaTeX in `.docx` to native **OMML** for **WPS** and Word. No MathType, no Word COM at runtime.

```bash
pip install -e ".[gui]"
python -m gui
# or
python convert_latex_docx.py examples/demo_before.docx examples/demo_after.docx
```

Requires `MML2OMML.XSL` from Microsoft Office.

License: [MIT](LICENSE)
