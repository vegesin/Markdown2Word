from __future__ import annotations

import json
import os
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk, font as tkfont
from tkinter.scrolledtext import ScrolledText

from ..win32.dialog_guard import MathTypeInlineDialogGuard
from ..core.pipeline import (
    NUMBER_FORMATS,
    ConfigurationError,
    ConversionConfig,
    PipelineRunner,
    find_pandoc,
    find_python,
    validate,
)


APP_NAME = "Markdown2Word"


class Markdown2WordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Markdown2Word 模板转换工具")
        self.geometry("920x720")
        self.minsize(780, 620)
        self.runner = PipelineRunner()
        self.dialog_guard = MathTypeInlineDialogGuard(
            on_click=lambda line: self.events.put(("log", line))
        )
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.settings_path = Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME / "settings.json"

        self.markdown_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.python_var = tk.StringVar(value=str(find_python()))
        self.pandoc_var = tk.StringVar(value=str(find_pandoc()))
        self.convert_formulas_var = tk.BooleanVar(value=True)
        self.add_numbers_var = tk.BooleanVar(value=True)
        self.number_format_var = tk.StringVar(value=next(iter(NUMBER_FORMATS)))
        self.create_visio_var = tk.BooleanVar(value=False)
        self.auto_word_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=True)
        self.skip_integrity_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="就绪")

        self.dark_var = tk.BooleanVar(value=False)
        self._build_ui()
        self._load_settings()
        self._apply_theme()
        self._sync_controls()
        self.after(100, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 16))
        ttk.Label(header, text="Markdown2Word", style="Title.TLabel").pack(side="left")
        ttk.Checkbutton(header, text="Dark mode", variable=self.dark_var, command=self._apply_theme).pack(side="right")
        files = ttk.LabelFrame(root, text="文件", padding=10)
        files.pack(fill="x")
        self._path_row(files, 0, "Markdown", self.markdown_var, self._browse_markdown)
        self._path_row(files, 1, "Word 模板", self.template_var, self._browse_template, optional=True)
        self._path_row(files, 2, "输出 DOCX", self.output_var, self._browse_output)
        files.columnconfigure(1, weight=1)

        options = ttk.LabelFrame(root, text="转换选项", padding=10)
        options.pack(fill="x", pady=(10, 0))
        ttk.Checkbutton(options, text="转换为可编辑 MathType 公式", variable=self.convert_formulas_var, command=self._sync_controls).grid(row=0, column=0, sticky="w", padx=(0, 18))
        self.numbers_check = ttk.Checkbutton(options, text="给行间公式添加右编号", variable=self.add_numbers_var, command=self._sync_controls)
        self.numbers_check.grid(row=0, column=1, sticky="w")
        ttk.Label(options, text="编号格式").grid(row=1, column=0, sticky="w", pady=(9, 0))
        self.number_combo = ttk.Combobox(options, textvariable=self.number_format_var, values=list(NUMBER_FORMATS), state="readonly", width=24)
        self.number_combo.grid(row=1, column=1, sticky="w", pady=(9, 0))
        ttk.Checkbutton(options, text="自动启动/连接 Microsoft Word", variable=self.auto_word_var).grid(row=2, column=0, sticky="w", pady=(9, 0))
        ttk.Checkbutton(options, text="允许覆盖已有输出", variable=self.overwrite_var).grid(row=2, column=1, sticky="w", pady=(9, 0))
        ttk.Checkbutton(options, text="将支持的 Mermaid 流程图转为可编辑 Visio 对象", variable=self.create_visio_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=(9, 0))
        ttk.Checkbutton(options, text="跳过非公式正文完整性检查（仅故障恢复时使用）", variable=self.skip_integrity_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(9, 0))

        tools = ttk.LabelFrame(root, text="工具路径（通常无需修改）", padding=10)
        tools.pack(fill="x", pady=(10, 0))
        self._path_row(tools, 0, "Python", self.python_var, lambda: self._browse_tool(self.python_var, "选择 Python", [("Python", "python.exe"), ("所有文件", "*.*")]))
        self._path_row(tools, 1, "Pandoc", self.pandoc_var, lambda: self._browse_tool(self.pandoc_var, "选择 Pandoc", [("Pandoc", "pandoc.exe"), ("所有文件", "*.*")]))
        tools.columnconfigure(1, weight=1)

        action = ttk.Frame(root)
        action.pack(fill="x", pady=(10, 6))
        self.start_button = ttk.Button(action, text="开始转换", command=self._start)
        self.start_button.pack(side="left")
        self.cancel_button = ttk.Button(action, text="取消", command=self._cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))
        ttk.Button(action, text="打开输出目录", command=self._open_output_directory).pack(side="left", padx=(8, 0))
        ttk.Label(action, textvariable=self.status_var).pack(side="right")

        log_box = ttk.LabelFrame(root, text="运行日志", padding=6)
        log_box.pack(fill="both", expand=True)
        self.log = ScrolledText(log_box, wrap="word", height=16, font=("Consolas", 9), state="disabled")
        self.log.pack(fill="both", expand=True)

    def _apply_theme(self) -> None:
        dark = self.dark_var.get()
        bg, panel, fg, muted, accent = (
            ("#18191c", "#25262b", "#f5f5f7", "#9b9da5", "#0a84ff") if dark else
            ("#f5f5f7", "#ffffff", "#1d1d1f", "#6e6e73", "#007aff")
        )
        available = set(tkfont.families(self))
        family = next((name for name in ("SF Pro Text", "SF Pro Display", "PingFang SC", "Helvetica Neue", "Microsoft YaHei UI", "Segoe UI") if name in available), "TkDefaultFont")
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            tkfont.nametofont(name).configure(family=family, size=10)
        style = ttk.Style(self)
        style.configure(".", font=(family, 10), background=bg, foreground=fg)
        style.configure("Title.TLabel", font=(family, 21, "bold"))
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TLabelframe.Label", foreground=muted)
        style.configure("TButton", padding=(12, 7), background=panel, foreground=fg)
        style.map("TButton", background=[("active", accent)], foreground=[("active", "white"), ("disabled", muted)])
        for widget in ("TEntry", "TCombobox", "TSpinbox"):
            style.configure(widget, fieldbackground=panel, foreground=fg, insertcolor=fg, padding=5)
            style.map(widget, fieldbackground=[("readonly", panel), ("disabled", bg)], foreground=[("disabled", muted), ("readonly", fg)])
        style.map("TCheckbutton", background=[("active", bg)], foreground=[("disabled", muted)])
        self.configure(background=bg)
        self.log.configure(background=panel, foreground=fg, insertbackground=fg, selectbackground=accent, font=(family, 10), relief="flat", borderwidth=0, padx=12, pady=10)
        self.option_add("*TCombobox*Listbox.background", panel)
        self.option_add("*TCombobox*Listbox.foreground", fg)

    def _path_row(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar, command, optional: bool = False) -> None:
        ttk.Label(parent, text=label + ("（可选）" if optional else "")).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="浏览…", command=command, width=9).grid(row=row, column=2, padx=(8, 0), pady=4)

    def _browse_markdown(self) -> None:
        value = filedialog.askopenfilename(title="选择 Markdown", filetypes=[("Markdown", "*.md *.markdown"), ("所有文件", "*.*")])
        if value:
            self.markdown_var.set(value)
            if not self.output_var.get().strip():
                self.output_var.set(str(Path(value).with_suffix(".docx")))

    def _browse_template(self) -> None:
        value = filedialog.askopenfilename(title="选择 Word 模板", filetypes=[("Word 模板", "*.docx *.dotx *.docm"), ("所有文件", "*.*")])
        if value:
            self.template_var.set(value)

    def _browse_output(self) -> None:
        initial = Path(self.markdown_var.get()).with_suffix(".docx") if self.markdown_var.get() else Path("output.docx")
        value = filedialog.asksaveasfilename(title="保存输出 Word", defaultextension=".docx", initialfile=initial.name, filetypes=[("Word 文档", "*.docx")])
        if value:
            self.output_var.set(value)

    def _browse_tool(self, variable: tk.StringVar, title: str, filetypes) -> None:
        value = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if value:
            variable.set(value)

    def _config(self) -> ConversionConfig:
        template = self.template_var.get().strip()
        label = self.number_format_var.get()
        return ConversionConfig(
            markdown=Path(self.markdown_var.get().strip()),
            template=Path(template) if template else None,
            output=Path(self.output_var.get().strip()),
            python=Path(self.python_var.get().strip()),
            pandoc=Path(self.pandoc_var.get().strip()),
            convert_formulas=self.convert_formulas_var.get(),
            add_numbers=self.add_numbers_var.get(),
            number_format=NUMBER_FORMATS.get(label, "Simple"),
            create_visio=self.create_visio_var.get(),
            auto_start_word=self.auto_word_var.get(),
            overwrite=self.overwrite_var.get(),
            skip_text_integrity_check=self.skip_integrity_var.get(),
        )

    def _sync_controls(self) -> None:
        formula_state = "normal" if self.convert_formulas_var.get() else "disabled"
        self.numbers_check.configure(state=formula_state)
        if not self.convert_formulas_var.get():
            self.add_numbers_var.set(False)
        self.number_combo.configure(state="readonly" if self.add_numbers_var.get() and self.convert_formulas_var.get() else "disabled")

    def _start(self) -> None:
        try:
            config = validate(self._config())
        except (ConfigurationError, OSError) as exc:
            messagebox.showerror("参数检查失败", str(exc), parent=self)
            return
        self._save_settings()
        self._set_running(True)
        self._append_log("\n============================================================\n")
        self._append_log(f"开始转换：{config.markdown}\n输出文件：{config.output}\n")
        if config.convert_formulas:
            # Backend UI Automation handles dialogs without moving the mouse.
            self._append_log("GUI MathType 行内公式弹窗保护已启动。\n")
        self.worker = threading.Thread(target=self._run_worker, args=(config,), daemon=True)
        self.worker.start()

    def _run_worker(self, config: ConversionConfig) -> None:
        try:
            code = self.runner.run(config, lambda line: self.events.put(("log", line)))
            self.events.put(("done", (code, config.output)))
        except Exception as exc:
            self.events.put(("error", str(exc)))
        finally:
            self.dialog_guard.stop()

    def _cancel(self) -> None:
        self.status_var.set("正在取消…")
        self.runner.cancel()

    def _drain_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "done":
                    code, output = payload
                    self._set_running(False)
                    if code == 0:
                        self.status_var.set("转换成功")
                        self._append_log(f"\n转换成功：{output}\n")
                        messagebox.showinfo("转换完成", f"已生成：\n{output}", parent=self)
                    elif code == 130:
                        self.status_var.set("已取消")
                        self._append_log("\n用户取消了转换。已生成的阶段文件保留用于排查。\n")
                    else:
                        self.status_var.set("转换失败")
                        messagebox.showerror("转换失败", "请查看日志末尾的原因与处理建议。中间文件已保留。", parent=self)
                elif kind == "error":
                    self._set_running(False)
                    self.status_var.set("转换失败")
                    self._append_log(f"\nGUI 调度错误：{payload}\n")
                    messagebox.showerror("转换失败", str(payload), parent=self)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")
        self.status_var.set("转换中，请勿操作 Word…" if running else "就绪")

    def _open_output_directory(self) -> None:
        raw = self.output_var.get().strip()
        directory = (Path(raw).expanduser().parent if raw else Path.cwd()).resolve()
        directory.mkdir(parents=True, exist_ok=True)
        os.startfile(directory)  # type: ignore[attr-defined]

    def _save_settings(self) -> None:
        data = {
            "template": self.template_var.get(),
            "output": self.output_var.get(),
            "python": self.python_var.get(),
            "pandoc": self.pandoc_var.get(),
            "convert_formulas": self.convert_formulas_var.get(),
            "add_numbers": self.add_numbers_var.get(),
            "number_format": self.number_format_var.get(),
            "create_visio": self.create_visio_var.get(),
            "auto_word": self.auto_word_var.get(),
            "overwrite": self.overwrite_var.get(),
            "dark_mode": self.dark_var.get(),
        }
        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            self.settings_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _load_settings(self) -> None:
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for key, variable in {
            "template": self.template_var,
            "output": self.output_var,
            "python": self.python_var,
            "pandoc": self.pandoc_var,
            "convert_formulas": self.convert_formulas_var,
            "add_numbers": self.add_numbers_var,
            "number_format": self.number_format_var,
            "create_visio": self.create_visio_var,
            "auto_word": self.auto_word_var,
            "overwrite": self.overwrite_var,
            "dark_mode": self.dark_var,
        }.items():
            if key in data:
                variable.set(data[key])

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("转换仍在运行", "确定取消转换并退出吗？", parent=self):
                return
            self.runner.cancel()
        self.dialog_guard.stop()
        self._save_settings()
        self.destroy()


def main() -> None:
    Markdown2WordApp().mainloop()


if __name__ == "__main__":
    main()
