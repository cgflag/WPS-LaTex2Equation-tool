# 小红书发布脚本（本地素材，可不提交 GIF 大文件）

## 标题建议

专利/论文 docx 里的 `$公式$` 一键变 Word 原生公式｜WPS 可用、不用 MathType

## 正文模板

写专利说明书时公式经常是 LaTeX 写法：`$E=mc^2$`、`$$\int_0^1 x^2 dx$$`  
手动敲公式太费时间，MathType 又要钱。

这个小工具 **批量把 docx 里的 `$...$` / `$$...$$` 转成 Word/WPS 原生公式**，块级公式自动 **居中 + (1)(2) 编号**（制表位实现）。

✅ WPS 文字可打开编辑  
✅ 不需要 MathType  
✅ 命令行一条搞定  

GitHub 开源：`<你的仓库链接>`

对比图放 `docs/xiaohongshu/`（转换前/后截图各 1 张）。

## 录屏 / GIF 步骤（约 30 秒）

1. 打开 `examples/demo_before.docx`（WPS 或 Word）
2. 终端运行：
   ```bash
   python convert_latex_docx.py examples/demo_before.docx examples/demo_out.docx
   ```
3. 打开 `demo_out.docx`，展示行内公式、块级居中编号
4. 用 ScreenToGif / OBS 导出 GIF，放到 README 的 `docs/assets/demo.gif`（可选）

## 标签

#专利撰写 #科研工具 #LaTeX #WPS #Word公式 #开源工具
