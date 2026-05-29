#!/usr/bin/env python3
"""WPS-LaTeX2Equation GUI — desktop tool for docx LaTeX → OMML conversion."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError as exc:
    raise SystemExit(
        "未安装 GUI 依赖。请运行：pip install -e \".[gui]\""
    ) from exc

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from convert_latex_docx import (  # noqa: E402
    TabLayoutConfig,
    __version__,
    convert_docx,
    find_mml2omml_xsl,
    parse_size_arg,
)
from gui.dnd_win import hook_dropfiles_once  # noqa: E402
from gui.settings import load_settings, save_settings  # noqa: E402

GITHUB_URL = "https://github.com/cgflag/WPS-LaTex2Equation-tool"
DEMO_DOCX = ROOT / "examples" / "demo_before.docx"
FONT_CHOICES = ("继承文档", "小四", "四号", "五号")


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_公式版.docx")


def resolve_xsl(custom: str | None) -> Path | None:
    if custom and Path(custom).is_file():
        return Path(custom)
    try:
        return find_mml2omml_xsl()
    except FileNotFoundError:
        return None


def parse_font_choice(choice: str) -> int | None:
    if choice == "继承文档":
        return None
    return parse_size_arg(choice)


class ConverterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(f"WPS 公式转换工具 v{__version__}")
        self.geometry("640x520")
        self.minsize(560, 480)

        self._settings = load_settings()
        self._input_path: Path | None = None
        self._output_path: Path | None = None
        self._xsl_path: Path | None = None
        self._busy = False
        self._failed_formulas: list[str] = []
        self._drop_queue: queue.Queue = queue.Queue()

        self._build_ui()
        self._refresh_xsl_status()
        self._poll_drop_queue()
        # 等窗口创建完成后再挂拖拽，且每个 HWND 只挂一次
        self.after(200, self._enable_drag_drop)

    def _build_ui(self) -> None:
        pad = {"padx": 16, "pady": 6}

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", **pad)
        ctk.CTkLabel(
            header,
            text="专为 AI 输出的 $...$ 批量转 WPS/Word 原生公式",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="不用 MathType · 不启动 Word · 默认另存为",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).pack(anchor="w")

        self._xsl_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12))
        self._xsl_label.pack(fill="x", padx=16)

        self._drop_zone = ctk.CTkFrame(
            self,
            height=72,
            border_width=2,
            border_color=("#94a3b8", "#64748b"),
            fg_color=("#f1f5f9", "#1e293b"),
        )
        self._drop_zone.pack(fill="x", padx=16, pady=(0, 6))
        self._drop_zone.pack_propagate(False)
        self._drop_hint = ctk.CTkLabel(
            self._drop_zone,
            text="将 .docx 文件拖拽到此处（或拖到窗口任意位置）",
            font=ctk.CTkFont(size=14),
            text_color=("#475569", "#cbd5e1"),
        )
        self._drop_hint.pack(expand=True)

        self._file_frame = ctk.CTkFrame(self)
        self._file_frame.pack(fill="x", **pad)
        file_frame = self._file_frame

        ctk.CTkLabel(file_frame, text="输入文档 (.docx)", anchor="w").pack(
            fill="x", padx=12, pady=(10, 0)
        )
        row_in = ctk.CTkFrame(file_frame, fg_color="transparent")
        row_in.pack(fill="x", padx=12, pady=4)
        self._input_var = ctk.StringVar()
        ctk.CTkEntry(row_in, textvariable=self._input_var, placeholder_text="选择或拖拽 docx 到窗口").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ctk.CTkButton(row_in, text="浏览…", width=72, command=self._pick_input).pack(side="left")

        ctk.CTkLabel(file_frame, text="输出文档", anchor="w").pack(fill="x", padx=12, pady=(8, 0))
        row_out = ctk.CTkFrame(file_frame, fg_color="transparent")
        row_out.pack(fill="x", padx=12, pady=(4, 12))
        self._output_var = ctk.StringVar()
        ctk.CTkEntry(row_out, textvariable=self._output_var, placeholder_text="默认同目录：原名_公式版.docx").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ctk.CTkButton(row_out, text="另存为…", width=72, command=self._pick_output).pack(side="left")

        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", **pad)
        ctk.CTkLabel(opts, text="公式字号").pack(side="left")
        self._font_combo = ctk.CTkComboBox(opts, values=list(FONT_CHOICES), width=140)
        self._font_combo.set("继承文档")
        self._font_combo.pack(side="left", padx=(8, 0))

        self._convert_btn = ctk.CTkButton(
            self,
            text="开始转换",
            height=44,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self._start_convert,
        )
        self._convert_btn.pack(fill="x", padx=16, pady=8)

        self._progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=16)
        self._progress.stop()
        self._progress.pack_forget()

        self._status_var = ctk.StringVar(value="请选择 docx 文件")
        ctk.CTkLabel(self, textvariable=self._status_var, anchor="w").pack(fill="x", padx=16, pady=4)

        self._fail_box = ctk.CTkTextbox(self, height=100, font=ctk.CTkFont(family="Consolas", size=12))
        self._fail_box.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        self._fail_box.insert("1.0", "转换失败的公式将显示在这里。\n")
        self._fail_box.configure(state="disabled")

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(bottom, text="打开示例文档", width=120, command=self._load_demo).pack(
            side="left", padx=(0, 8)
        )
        ctk.CTkButton(bottom, text="GitHub", width=80, command=lambda: webbrowser.open(GITHUB_URL)).pack(
            side="left"
        )

        self._advanced_open = False
        self._adv_btn = ctk.CTkButton(
            self, text="▸ 高级设置", fg_color="transparent", text_color=("gray10", "gray90"),
            anchor="w", command=self._toggle_advanced,
        )
        self._adv_btn.pack(fill="x", padx=12)
        self._adv_frame = ctk.CTkFrame(self)
        ctk.CTkLabel(self._adv_frame, text="MML2OMML.XSL 路径（可选，留空则自动检测）", anchor="w").pack(
            fill="x", padx=12, pady=(8, 0)
        )
        xsl_row = ctk.CTkFrame(self._adv_frame, fg_color="transparent")
        xsl_row.pack(fill="x", padx=12, pady=8)
        self._xsl_var = ctk.StringVar(value=self._settings.get("custom_xsl", ""))
        ctk.CTkEntry(xsl_row, textvariable=self._xsl_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(xsl_row, text="浏览…", width=72, command=self._pick_xsl).pack(side="left")
        ctk.CTkButton(
            self._adv_frame, text="应用 XSL 路径", command=self._apply_custom_xsl
        ).pack(padx=12, pady=(0, 12), anchor="w")

    def _toggle_advanced(self) -> None:
        self._advanced_open = not self._advanced_open
        if self._advanced_open:
            self._adv_btn.configure(text="▾ 高级设置")
            self._adv_frame.pack(fill="x", padx=16, pady=(0, 8))
        else:
            self._adv_frame.pack_forget()
            self._adv_btn.configure(text="▸ 高级设置")

    def _enable_drag_drop(self) -> None:
        if sys.platform != "win32":
            self._drop_hint.configure(text="拖拽仅支持 Windows；请使用「浏览…」选择文件")
            return

        ok_zone = hook_dropfiles_once(self._drop_zone, self._drop_queue)
        ok_frame = hook_dropfiles_once(self._file_frame, self._drop_queue)

        if ok_zone or ok_frame:
            self._status_var.set("可将 docx 拖到上方拖放区")
        else:
            self._drop_hint.configure(
                text="拖拽不可用，请用「浏览…」选择文件（或重装带 [gui] 依赖）"
            )

    @staticmethod
    def _decode_drop_path(raw) -> Path:
        if isinstance(raw, bytes):
            for enc in ("utf-8", "gbk", "utf-16-le"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                text = raw.decode(errors="replace")
        else:
            text = str(raw)
        text = text.strip().strip("{}").strip('"').strip("'")
        return Path(text)

    def _poll_drop_queue(self) -> None:
        """主线程轮询 windnd 放入的路径，避免 GIL 崩溃。"""
        try:
            while True:
                batch = self._drop_queue.get_nowait()
                try:
                    self._handle_dropped_files(batch)
                except Exception as exc:
                    messagebox.showerror("拖拽处理失败", str(exc))
        except queue.Empty:
            pass
        self.after(50, self._poll_drop_queue)

    def _handle_dropped_files(self, paths: list) -> None:
        if self._busy:
            messagebox.showinfo("提示", "正在转换中，请稍候完成后再拖入文件。")
            return

        docx_paths: list[Path] = []
        for raw in paths:
            p = self._decode_drop_path(raw)
            if p.suffix.lower() == ".docx" and p.is_file():
                docx_paths.append(p.resolve())

        if not docx_paths:
            self._flash_drop_zone(ok=False)
            messagebox.showwarning(
                "无法添加",
                "请拖入有效的 .docx 文件。\n（快捷方式或非 docx 文件不支持）",
            )
            return

        self._set_input(docx_paths[0])
        self._flash_drop_zone(ok=True)
        if len(docx_paths) > 1:
            messagebox.showinfo(
                "提示",
                f"已添加：{docx_paths[0].name}\n\n"
                f"共拖入 {len(docx_paths)} 个文件，当前仅支持单文件转换，已使用第一个。",
            )

    def _flash_drop_zone(self, ok: bool) -> None:
        if ok:
            self._drop_zone.configure(border_color=("#16a34a", "#22c55e"))
            self._drop_hint.configure(
                text="✓ 已添加输入文档",
                text_color=("#15803d", "#4ade80"),
            )
        else:
            self._drop_zone.configure(border_color=("#dc2626", "#f87171"))
            self._drop_hint.configure(
                text="拖入失败，请使用 .docx 文件",
                text_color=("#b91c1c", "#fca5a5"),
            )
        self.after(2000, self._reset_drop_zone_hint)

    def _reset_drop_zone_hint(self) -> None:
        self._drop_zone.configure(border_color=("#94a3b8", "#64748b"))
        self._drop_hint.configure(
            text="将 .docx 文件拖拽到此处",
            text_color=("#475569", "#cbd5e1"),
        )

    def _refresh_xsl_status(self) -> None:
        custom = self._xsl_var.get().strip() if hasattr(self, "_xsl_var") else ""
        custom = custom or self._settings.get("custom_xsl", "")
        self._xsl_path = resolve_xsl(custom or None)
        if self._xsl_path:
            self._xsl_label.configure(
                text=f"✓ 转换组件已就绪：{self._xsl_path.name}",
                text_color=("#1a7f37", "#2ea043"),
            )
            self._convert_btn.configure(state="normal")
        else:
            self._xsl_label.configure(
                text="✗ 未找到 MML2OMML.XSL。请安装 Microsoft Office，或在高级设置中指定路径。",
                text_color=("#b42318", "#f85149"),
            )
            self._convert_btn.configure(state="disabled")

    def _apply_custom_xsl(self) -> None:
        path = self._xsl_var.get().strip()
        self._settings["custom_xsl"] = path
        save_settings(self._settings)
        self._refresh_xsl_status()
        if self._xsl_path:
            messagebox.showinfo("XSL", "已应用自定义 XSL 路径。")
        else:
            messagebox.showwarning(
                "XSL",
                "仍未找到有效 XSL 文件。\n\n"
                "请安装 Office，或将 MML2OMML.XSL 的完整路径填入后重试。\n"
                "常见位置：\nC:\\Program Files\\Microsoft Office\\root\\Office16\\MML2OMML.XSL",
            )

    def _pick_xsl(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 MML2OMML.XSL",
            filetypes=[("XSL", "*.xsl"), ("All", "*.*")],
        )
        if path:
            self._xsl_var.set(path)
            self._apply_custom_xsl()

    def _pick_input(self) -> None:
        initial = self._settings.get("last_dir", str(Path.home()))
        path = filedialog.askopenfilename(
            title="选择 docx",
            initialdir=initial,
            filetypes=[("Word 文档", "*.docx")],
        )
        if path:
            self._set_input(Path(path))

    def _pick_output(self) -> None:
        if not self._input_path:
            messagebox.showwarning("提示", "请先选择输入文档。")
            return
        initial = str(self._output_path.parent if self._output_path else self._input_path.parent)
        path = filedialog.asksaveasfilename(
            title="另存为",
            initialdir=initial,
            initialfile=default_output_path(self._input_path).name,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
        )
        if path:
            self._output_path = Path(path)
            self._output_var.set(str(self._output_path))

    def _set_input(self, path: Path) -> None:
        self._input_path = path.resolve()
        self._input_var.set(str(self._input_path))
        self._output_path = default_output_path(self._input_path)
        self._output_var.set(str(self._output_path))
        self._settings["last_dir"] = str(self._input_path.parent)
        save_settings(self._settings)
        self._status_var.set(f"已选择：{self._input_path.name}")
        if hasattr(self, "_drop_hint"):
            self._drop_hint.configure(
                text=f"当前文件：{self._input_path.name}",
                text_color=("#1d4ed8", "#60a5fa"),
            )

    def _load_demo(self) -> None:
        if not DEMO_DOCX.is_file():
            messagebox.showerror("错误", f"未找到示例文件：\n{DEMO_DOCX}")
            return
        self._set_input(DEMO_DOCX)

    def _start_convert(self) -> None:
        if self._busy or not self._xsl_path:
            return
        if not self._input_path or not self._input_path.is_file():
            messagebox.showwarning("提示", "请选择有效的输入 docx 文件。")
            return

        out_text = self._output_var.get().strip()
        output_path = Path(out_text) if out_text else default_output_path(self._input_path)
        if output_path.resolve() == self._input_path.resolve():
            messagebox.showerror("错误", "输出文件不能与输入文件相同。\n请使用另存为（默认 _公式版.docx）。")
            return

        try:
            override_sz = parse_font_choice(self._font_combo.get())
        except ValueError as e:
            messagebox.showerror("字号错误", str(e))
            return

        self._output_path = output_path
        self._busy = True
        self._convert_btn.configure(state="disabled")
        self._progress.pack(fill="x", padx=16, pady=4)
        self._progress.start()
        self._status_var.set("正在转换，请稍候…")
        self._clear_failures()

        thread = threading.Thread(
            target=self._run_convert,
            args=(self._input_path, self._output_path, override_sz, self._xsl_path),
            daemon=True,
        )
        thread.start()

    def _run_convert(
        self,
        input_path: Path,
        output_path: Path,
        override_sz: int | None,
        xsl_path: Path,
    ) -> None:
        try:
            ok, fail, failures = convert_docx(
                input_path,
                output_path,
                override_sz,
                TabLayoutConfig(mode="page"),
                xsl_path,
            )
            self.after(0, lambda: self._on_convert_done(True, ok, fail, failures, output_path, None))
        except Exception as e:
            self.after(0, lambda: self._on_convert_done(False, 0, 0, [], output_path, str(e)))

    def _on_convert_done(
        self,
        success: bool,
        ok: int,
        fail: int,
        failures: list[str],
        output_path: Path,
        error: str | None,
    ) -> None:
        self._busy = False
        self._progress.stop()
        self._progress.pack_forget()
        self._refresh_xsl_status()

        if not success:
            self._status_var.set("转换失败")
            messagebox.showerror("转换失败", error or "未知错误")
            return

        self._failed_formulas = failures
        self._status_var.set(f"完成：成功 {ok} 个，失败 {fail} 个 → {output_path.name}")
        if failures:
            self._show_failures(failures)

        msg = f"转换完成！\n\n成功：{ok} 个\n失败：{fail} 个\n\n输出文件：\n{output_path}"
        if fail:
            msg += "\n\n失败公式仍保留为 $...$ 文本，详见下方列表。"
        if messagebox.askyesno("转换完成", msg + "\n\n是否打开输出文件所在文件夹？"):
            self._open_folder(output_path.parent)

    def _clear_failures(self) -> None:
        self._fail_box.configure(state="normal")
        self._fail_box.delete("1.0", "end")
        self._fail_box.configure(state="disabled")

    def _show_failures(self, failures: list[str]) -> None:
        lines = "\n".join(f"· ${f}$" if not f.startswith("$") else f"· {f}" for f in failures)
        self._fail_box.configure(state="normal")
        self._fail_box.delete("1.0", "end")
        self._fail_box.insert("1.0", f"以下 {len(failures)} 条公式未能转换：\n\n{lines}\n")
        self._fail_box.configure(state="disabled")

    @staticmethod
    def _open_folder(folder: Path) -> None:
        folder = folder.resolve()
        if sys.platform == "win32":
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(folder)], check=False)
        else:
            subprocess.run(["xdg-open", str(folder)], check=False)


def run() -> None:
    try:
        app = ConverterApp()
        app.mainloop()
    except Exception as exc:
        try:
            root = ctk.CTk()
            root.withdraw()
            messagebox.showerror("启动失败", f"{exc}\n\n请尝试：pip install -e \".[gui]\"")
            root.destroy()
        except Exception:
            print("启动失败:", exc, file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    run()
