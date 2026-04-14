from __future__ import annotations

import json
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import standardize_docx_paper as formatter


class DocxFormatterGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DOCX 论文格式化工具")
        self.root.geometry("1100x760")

        cwd = Path.cwd().resolve()
        self.input_file_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value=str((cwd / "output" / "doc").resolve()))
        self.config_path_var = tk.StringVar(value=str(formatter.CONFIG_PATH.resolve()))
        self.overwrite_var = tk.BooleanVar(value=False)
        self.header_footer_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.load_config_into_editor()


    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(3, weight=1)

        path_frame = ttk.LabelFrame(self.root, text="文档与配置")
        path_frame.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))
        path_frame.columnconfigure(1, weight=1)
        path_frame.columnconfigure(4, weight=0)

        ttk.Label(path_frame, text="导入文档").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(path_frame, textvariable=self.input_file_var).grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(path_frame, text="选择文档", command=self.choose_input_file).grid(row=0, column=2, padx=8, pady=8)

        ttk.Label(path_frame, text="导出目录").grid(row=1, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(path_frame, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(path_frame, text="选择目录", command=self.choose_output_dir).grid(row=1, column=2, padx=8, pady=8)

        ttk.Label(path_frame, text="配置文件").grid(row=2, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(path_frame, textvariable=self.config_path_var).grid(row=2, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(path_frame, text="选择配置", command=self.choose_config_path).grid(row=2, column=2, padx=8, pady=8)

        option_frame = ttk.LabelFrame(self.root, text="运行选项")
        option_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=6)

        ttk.Checkbutton(option_frame, text="允许覆盖输出文件", variable=self.overwrite_var).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        ttk.Checkbutton(option_frame, text="写入页眉页脚", variable=self.header_footer_var).grid(row=0, column=1, padx=10, pady=8, sticky="w")

        button_frame = ttk.Frame(option_frame)
        button_frame.grid(row=0, column=2, padx=10, pady=8, sticky="e")
        option_frame.columnconfigure(2, weight=1)

        self.reload_button = ttk.Button(button_frame, text="重新加载配置", command=self.load_config_into_editor)
        self.reload_button.grid(row=0, column=0, padx=4)

        self.save_button = ttk.Button(button_frame, text="保存配置", command=self.save_config_from_editor)
        self.save_button.grid(row=0, column=1, padx=4)

        self.run_button = ttk.Button(button_frame, text="开始格式化", command=self.start_formatting)
        self.run_button.grid(row=0, column=2, padx=4)

        config_frame = ttk.LabelFrame(self.root, text="配置编辑器（保存后下次格式化立即生效）")
        config_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=6)
        config_frame.columnconfigure(0, weight=1)
        config_frame.rowconfigure(0, weight=1)

        self.config_text = ScrolledText(config_frame, wrap="none", font=("Menlo", 12))
        self.config_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        log_frame = ttk.LabelFrame(self.root, text="运行日志")
        log_frame.grid(row=3, column=0, sticky="nsew", padx=12, pady=(6, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Menlo", 11), height=10, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)


    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


    def choose_input_file(self) -> None:
        initial_path = self.input_file_var.get().strip()
        initial_dir = str(Path(initial_path).expanduser().parent) if initial_path else str(Path.cwd())
        selected = filedialog.askopenfilename(
            title="选择导入文档",
            initialdir=initial_dir,
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        self.input_file_var.set(selected)

        output_path = Path(self.output_dir_var.get()).expanduser()
        if not str(output_path).strip() or output_path == Path.cwd() / "output" / "doc":
            self.output_dir_var.set(str((Path(selected).parent / "output" / "doc").resolve()))


    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择导出目录", initialdir=self.output_dir_var.get() or str(Path.cwd()))
        if selected:
            self.output_dir_var.set(selected)


    def choose_config_path(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择配置文件",
            initialdir=str(Path(self.config_path_var.get()).expanduser().parent),
            filetypes=[("JSONC 文件", "*.jsonc"), ("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        self.config_path_var.set(selected)
        self.load_config_into_editor()


    def load_config_into_editor(self) -> None:
        try:
            config_path, config_text, config = formatter.load_config_text(Path(self.config_path_var.get()))
        except Exception as exc:
            messagebox.showerror("配置加载失败", str(exc))
            return

        self.config_path_var.set(str(config_path))
        self.config_text.delete("1.0", "end")
        self.config_text.insert("1.0", config_text)
        self.apply_config_options(config)
        self.append_log(f"已加载配置: {config_path}")


    def apply_config_options(self, config: dict) -> None:
        general = config.get("general", {})
        self.overwrite_var.set(bool(general.get("overwrite", False)))
        self.header_footer_var.set(bool(general.get("add_header_footer", True)))

        if not self.output_dir_var.get().strip():
            input_file = Path(self.input_file_var.get()).expanduser()
            default_subdir = general.get("default_output_subdir", "output/doc")
            if input_file.exists():
                self.output_dir_var.set(str((input_file.parent / Path(default_subdir)).resolve()))


    def save_config_from_editor(self, show_message: bool = True) -> tuple[Path, dict] | None:
        raw_text = self.config_text.get("1.0", "end").strip()
        try:
            config_path, merged_config = formatter.save_config_text(raw_text, Path(self.config_path_var.get()))
        except json.JSONDecodeError as exc:
            messagebox.showerror("配置格式错误", f"JSONC 解析失败：{exc}")
            return None
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))
            return None

        self.config_path_var.set(str(config_path))
        self.config_text.delete("1.0", "end")
        self.config_text.insert("1.0", formatter.render_config_text(merged_config))
        self.apply_config_options(merged_config)
        self.append_log(f"已保存配置: {config_path}")

        if show_message:
            messagebox.showinfo("保存成功", f"配置已保存到：\n{config_path}")
        return config_path, merged_config


    def set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in (self.reload_button, self.save_button, self.run_button):
            button.configure(state=state)


    def start_formatting(self) -> None:
        input_file = Path(self.input_file_var.get()).expanduser()
        if not input_file.exists() or not input_file.is_file() or input_file.suffix.lower() != ".docx":
            messagebox.showerror("导入文档无效", "请选择一个存在的 .docx 文档。")
            return

        output_dir = Path(self.output_dir_var.get()).expanduser()
        save_result = self.save_config_from_editor(show_message=False)
        if save_result is None:
            return
        config_path, _ = save_result
        overwrite = self.overwrite_var.get()
        add_header_footer = self.header_footer_var.get()

        self.set_controls_enabled(False)
        self.append_log(f"开始格式化，输入文档：{input_file}")
        self.append_log(f"导出目录：{output_dir}")

        worker = threading.Thread(
            target=self._run_formatting_worker,
            args=(input_file, output_dir, config_path, overwrite, add_header_footer),
            daemon=True,
        )
        worker.start()


    def _run_formatting_worker(
        self,
        input_file: Path,
        output_dir: Path,
        config_path: Path,
        overwrite: bool,
        add_header_footer: bool,
    ) -> None:
        try:
            produced, errors = formatter.run_batch(
                [input_file],
                output_dir=output_dir,
                overwrite=overwrite,
                add_header_footer=add_header_footer,
                config_path=config_path,
            )
        except Exception as exc:
            self.root.after(0, self._finish_formatting, [], [str(exc)])
            return

        self.root.after(0, self._finish_formatting, produced, errors)


    def _finish_formatting(self, produced: list[Path], errors: list[str]) -> None:
        self.set_controls_enabled(True)

        if produced:
            self.append_log(f"格式化完成，共生成 {len(produced)} 个文件：")
            for path in produced:
                self.append_log(f"  {path}")

        if errors:
            self.append_log("出现以下问题：")
            for message in errors:
                self.append_log(f"  {message}")
            messagebox.showwarning("处理完成（含问题）", "\n".join(errors[:5]))
            return

        messagebox.showinfo("处理完成", f"已成功生成 {len(produced)} 个文件。")


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    DocxFormatterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
