# patent-math-tool

**不依赖 Word COM / OMath API，WPS 可直接打开；专为 AI 输出的 `$...$` 批量转原生公式。**

Batch LaTeX-in-docx → native OMML for **WPS** & Word · no MathType · no COM automation

**中文** | [English](#english)

> 搜索关键词：`latex` `docx` `equation` `wps` `word` `omml` `patent` `formula`  
> 仓库别名也可考虑：`docx-latex-math`、`patent-formula-converter`

<!-- 将录屏 GIF 放到 docs/assets/demo.gif 后取消下行注释 -->
<!-- ![Demo](docs/assets/demo.gif) -->

| 转换前 | 转换后 |
|--------|--------|
| [`examples/demo_before.docx`](examples/demo_before.docx) | [`examples/demo_after.docx`](examples/demo_after.docx) |

---

## 解决什么问题

| 痛点 | 本工具 |
|------|--------|
| AI / Markdown 写专利，公式是 `$...$` 纯文本 | 一键批量转成 **Word/WPS 原生公式** |
| MathType 收费、宏在 WPS 上不稳定 | **无需 MathType**；运行时 **不调用 Word COM / OMath API** |
| 块级公式要居中 + (1)(2) 编号 | 自动识别 `$$...$$` / 独立段 `$...$`，**制表位居中 + 编号** |
| 转换失败怕整篇文档坏掉 | 失败处 **保留原 `$...$` 文本** |

技术路径：直接改写 docx 内 `document.xml`（LaTeX → MathML → OMML），输出文件 **WPS 与 Word 均可打开编辑**。

---

## 开源范围

| 模块 | 许可 |
|------|------|
| ✅ 单文件 CLI、基础 LaTeX→OMML、行内/块级识别、制表位编号 | **MIT 开源**（本仓库） |
| 🔒 批量文件夹、GUI、专利模板 preset、题注/目录一体化 | 可闭源 / Pro（未包含） |

---

## 特性

- ✅ **WPS 文字** 打开即可编辑（OOXML 原生公式）
- ✅ **不需要 MathType**
- ✅ 支持 `$...$`（行内）与 `$$...$$`（块级）
- ✅ 不依赖运行时 Word COM / `OMaths.BuildUp`（与 VBA 宏方案不同）
- ⚠️ 需要本机 **Microsoft Office 的 `MML2OMML.XSL`**（仅作 MathML→OMML 转换，不启动 Word）

---

## 快速开始

### 环境

- Python 3.10+
- `MML2OMML.XSL`（通常位于 Office 安装目录；可用 `--xsl` 或环境变量指定）

### 安装

```bash
git clone https://github.com/YOUR_USER/patent-math-tool.git
cd patent-math-tool
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -e .                # 推荐：含 pyproject.toml 入口
# 或: pip install -r requirements.txt
```

安装后可使用命令：

```bash
patent-math input.docx output.docx
# 等价别名
docx-latex-math input.docx
```

### 一行示例

```bash
patent-math examples/demo_before.docx examples/demo_after.docx
```

### 常用参数

```bash
python convert_latex_docx.py input.docx                    # → input_converted.docx
python convert_latex_docx.py input.docx out.docx --size 小四
python convert_latex_docx.py input.docx out.docx --tab-mode page
set PATENT_MATH_MML2OMML=C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
```

| 写法 | 效果 |
|------|------|
| 句中 `$E=mc^2$` | 行内公式 |
| 整段仅 `$$...$$` 或单独一行 `$...$` | 块级：居中 + `(n)` |
| 转换失败 | 保留 `$...$` 原文 |

重新生成示例：`python scripts/build_examples.py`

---

## 项目结构

```
patent-math-tool/
├── convert_latex_docx.py
├── pyproject.toml
├── requirements.txt
├── README.md
├── CHANGELOG.md
├── LICENSE
├── examples/
│   ├── demo_before.docx
│   └── demo_after.docx
├── tests/
├── .github/workflows/      # CI + Release
├── docs/xiaohongshu/
├── scripts/
└── legacy/                 # 旧 VBA 宏（可选）
```

---

## 开发

```bash
pip install -e ".[dev]"
pytest -q
```

CI：全平台单元测试；Windows runner 在检测到 Office XSL 时跑完整 docx 转换冒烟。

---

## 发布 GitHub 时建议

1. 仓库名：`patent-math-tool`（或 `docx-latex-math`）
2. **Topics**：`latex`, `docx`, `equation`, `wps`, `word`, `omml`, `patent`, `python`
3. About 描述填 Star 差异化那句（见文首）
4. 打 tag `v1.0.0` 触发 Release（附 examples；exe 见 `scripts/build_exe.md`）

---

## 已知限制

- 仅正文 `word/document.xml`（不含页眉页脚、文本框、表格）
- LaTeX 子集受 `latex2mathml` 限制
- 需 `MML2OMML.XSL`（不能随仓库分发，需用户本机 Office）

---

## 许可证

[MIT](LICENSE) — 见上方「开源范围」。

---

<a id="english"></a>

## English

**No Word COM / OMath API at runtime — output opens in WPS; built for batch-converting AI-generated `$...$` in docx to native equations.**

Convert `$...$` (inline) and `$$...$$` (display) LaTeX inside `.docx` to native **OMML** equations. Display math: **centered** + **(1)(2)...** via tab stops.

### Why this exists

- Patent/spec drafts from LLMs often leave formulas as literal `$...$` text
- MathType costs money; VBA + `OMaths` is fragile on WPS
- This tool rewrites OOXML directly (LaTeX → MathML → OMML) — **no MathType**, **no COM**

### Open source vs Pro

| Included (MIT) | Not in repo (Pro / private) |
|----------------|----------------------------|
| Single-file CLI, inline/display detection, tab numbering | Folder batch, GUI, patent templates, caption/TOC integration |

### Install & run

```bash
pip install -e .
patent-math examples/demo_before.docx examples/demo_after.docx
```

Requires `MML2OMML.XSL` from Microsoft Office (`--xsl` or `PATENT_MATH_MML2OMML`).

License: [MIT](LICENSE)
