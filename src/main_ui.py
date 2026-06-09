"""CustomTkinter desktop UI for batch invoice extraction."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import customtkinter as ctk

from src.excel_builder import AVAILABLE_OUTPUT_COLUMNS
from src.file_manager import DEFAULT_EXPORT_COLUMNS, BatchProgress, BatchResult, process_invoice_folder


APP_TITLE = "数电发票提取工具"
DEFAULT_OUTPUT_NAME = "发票汇总.xlsx"
MODE_LABELS = {
    "汇总+明细": "items",
    "发票一行": "invoice",
}
UI_DEFAULT_EXPORT_COLUMNS = [
    "发票号码",
    "销售方名称",
    "金额合计",
    "税额合计",
    "税率",
    "价税合计",
    "开票日期",
]

COLOR_BG = "#EEF2FA"
COLOR_PANEL = "#FFFFFF"
COLOR_INPUT = "#F8FAFF"
COLOR_INPUT_SOFT = "#F7F9FE"
COLOR_TEXT = "#2D3142"
COLOR_MUTED = "#8B93A7"
COLOR_ACCENT = "#7775F2"
COLOR_ACCENT_HOVER = "#6866DD"
COLOR_SECONDARY = "#EDF0FB"
COLOR_SECONDARY_HOVER = "#E2E7F7"
COLOR_ERROR = "#E85D75"


def _font_family() -> str:
    if sys.platform == "darwin":
        return "PingFang SC"
    if sys.platform == "win32":
        return "Microsoft YaHei UI"
    return "Arial"


class InvoiceExtractorApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1240x760")
        self.minsize(1060, 680)
        self.configure(fg_color=COLOR_BG)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.font_family = _font_family()
        self.font = ctk.CTkFont(family=self.font_family, size=13)
        self.font_small = ctk.CTkFont(family=self.font_family, size=12)
        self.font_title = ctk.CTkFont(family=self.font_family, size=16, weight="bold")
        self.font_section = ctk.CTkFont(family=self.font_family, size=13, weight="bold")
        self.font_button = ctk.CTkFont(family=self.font_family, size=14, weight="bold")
        self.font_log = ctk.CTkFont(family=self.font_family, size=13)

        self.invoice_dir = ctk.StringVar(value="")
        self.save_dir = ctk.StringVar(value=str(Path.cwd() / "outputs"))
        self.output_name = ctk.StringVar(value=DEFAULT_OUTPUT_NAME)
        self.mode_label = ctk.StringVar(value="汇总+明细")
        self.status_text = ctk.StringVar(value="空闲")
        self.is_running = False
        self.last_result: BatchResult | None = None
        self.worker_thread: threading.Thread | None = None
        self.ui_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.column_vars: dict[str, ctk.BooleanVar] = {}

        self._build_layout()
        self.after(120, self._drain_ui_queue)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=0, minsize=352)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_view()

    def _build_sidebar(self) -> None:
        sidebar = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        sidebar.grid(row=0, column=0, padx=(22, 12), pady=22, sticky="nsew")
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_rowconfigure(9, weight=1)

        ctk.CTkLabel(sidebar, text=APP_TITLE, font=self.font_title, text_color=COLOR_TEXT, anchor="w").grid(
            row=0, column=0, sticky="ew", pady=(0, 18)
        )

        self._section_label(sidebar, 1, "路径")
        self._path_picker(sidebar, 2, "发票文件夹", self.invoice_dir, self.select_invoice_folder)
        self._path_picker(sidebar, 3, "保存文件夹", self.save_dir, self.select_save_folder)

        self._section_label(sidebar, 4, "输出")
        ctk.CTkEntry(
            sidebar,
            textvariable=self.output_name,
            font=self.font,
            height=36,
            corner_radius=10,
            border_width=0,
            fg_color=COLOR_INPUT,
            text_color=COLOR_TEXT,
        ).grid(row=5, column=0, sticky="ew", pady=(0, 10))

        self.mode_segment = ctk.CTkSegmentedButton(
            sidebar,
            values=list(MODE_LABELS.keys()),
            variable=self.mode_label,
            font=self.font,
            height=34,
            corner_radius=10,
            selected_color=COLOR_ACCENT,
            selected_hover_color=COLOR_ACCENT_HOVER,
            unselected_color=COLOR_SECONDARY,
            unselected_hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_TEXT,
        )
        self.mode_segment.grid(row=6, column=0, sticky="ew", pady=(0, 18))

        field_header = ctk.CTkFrame(sidebar, fg_color="transparent")
        field_header.grid(row=7, column=0, sticky="ew", pady=(0, 8))
        field_header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(field_header, text="输出字段", font=self.font_section, text_color=COLOR_TEXT, anchor="w").grid(
            row=0, column=0, sticky="ew"
        )
        self._tiny_button(field_header, "清空", self.clear_all_columns).grid(row=0, column=1, padx=(8, 0))
        self._tiny_button(field_header, "默认", self.reset_default_columns).grid(row=0, column=2, padx=(6, 0))
        self._tiny_button(field_header, "全选", self.select_all_columns).grid(row=0, column=3, padx=(6, 0))

        fields_frame = ctk.CTkScrollableFrame(
            sidebar,
            fg_color=COLOR_INPUT,
            corner_radius=14,
            border_width=0,
            scrollbar_button_color=COLOR_SECONDARY,
            scrollbar_button_hover_color=COLOR_SECONDARY_HOVER,
        )
        fields_frame.grid(row=8, column=0, sticky="nsew", pady=(0, 16))
        for row, column_name in enumerate(AVAILABLE_OUTPUT_COLUMNS):
            enabled = column_name in UI_DEFAULT_EXPORT_COLUMNS
            variable = ctk.BooleanVar(value=enabled)
            self.column_vars[column_name] = variable
            checkbox = ctk.CTkCheckBox(
                fields_frame,
                text=column_name,
                variable=variable,
                font=self.font,
                corner_radius=4,
                border_width=1,
                fg_color=COLOR_ACCENT,
                hover_color=COLOR_ACCENT_HOVER,
                text_color=COLOR_TEXT,
                checkmark_color="#FFFFFF",
            )
            checkbox.grid(row=row, column=0, padx=12, pady=5, sticky="w")

        self.extract_button = ctk.CTkButton(
            sidebar,
            text="提取执行",
            font=self.font_button,
            height=44,
            corner_radius=12,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#FFFFFF",
            command=self.start_extract,
        )
        self.extract_button.grid(row=10, column=0, sticky="ew", pady=(0, 10))

        utility = ctk.CTkFrame(sidebar, fg_color="transparent")
        utility.grid(row=11, column=0, sticky="ew")
        utility.grid_columnconfigure((0, 1, 2), weight=1)
        self.preview_button = self._secondary_button(utility, "预览", self.preview_data)
        self.preview_button.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        self.open_button = self._secondary_button(utility, "打开", self.open_save_folder)
        self.open_button.grid(row=0, column=1, padx=3, sticky="ew")
        self.summary_button = self._secondary_button(utility, "汇总", self.show_summary)
        self.summary_button.grid(row=0, column=2, padx=(6, 0), sticky="ew")
        self.action_buttons = [self.extract_button, self.preview_button, self.open_button, self.summary_button]

    def _build_main_view(self) -> None:
        main = ctk.CTkFrame(self, fg_color=COLOR_PANEL, corner_radius=20, border_width=0)
        main.grid(row=0, column=1, padx=(12, 22), pady=22, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=3)
        main.grid_rowconfigure(4, weight=1)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, padx=24, pady=(24, 12), sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="处理日志", font=self.font_title, text_color=COLOR_TEXT, anchor="w").grid(
            row=0, column=0, sticky="ew"
        )
        ctk.CTkLabel(header, textvariable=self.status_text, font=self.font, text_color=COLOR_MUTED, anchor="e").grid(
            row=0, column=1
        )

        self.progress_bar = ctk.CTkProgressBar(
            main,
            height=6,
            corner_radius=3,
            progress_color=COLOR_ACCENT,
            fg_color=COLOR_SECONDARY,
        )
        self.progress_bar.grid(row=1, column=0, padx=24, pady=(0, 14), sticky="ew")
        self.progress_bar.set(0)

        self.message_text = ctk.CTkTextbox(
            main,
            wrap="word",
            state="disabled",
            font=self.font_log,
            corner_radius=14,
            fg_color=COLOR_INPUT_SOFT,
            border_width=0,
            text_color=COLOR_TEXT,
        )
        self.message_text.grid(row=2, column=0, padx=24, pady=(0, 18), sticky="nsew")

        error_header = ctk.CTkFrame(main, fg_color="transparent")
        error_header.grid(row=3, column=0, padx=24, pady=(0, 8), sticky="ew")
        error_header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(error_header, text="错误日志", font=self.font_section, text_color=COLOR_ERROR, anchor="w").grid(
            row=0, column=0, sticky="ew"
        )

        self.error_text = ctk.CTkTextbox(
            main,
            wrap="word",
            state="disabled",
            font=self.font_log,
            corner_radius=14,
            fg_color=COLOR_INPUT_SOFT,
            border_width=0,
            text_color=COLOR_ERROR,
        )
        self.error_text.grid(row=4, column=0, padx=24, pady=(0, 24), sticky="nsew")

    def _section_label(self, parent: ctk.CTkFrame, row: int, text: str) -> None:
        ctk.CTkLabel(parent, text=text, font=self.font_section, text_color=COLOR_TEXT, anchor="w").grid(
            row=row, column=0, sticky="ew", pady=(0, 8)
        )

    def _path_picker(self, parent: ctk.CTkFrame, row: int, label: str, variable: ctk.StringVar, command) -> None:
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        wrapper.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(wrapper, text=label, font=self.font_small, text_color=COLOR_MUTED, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5)
        )
        ctk.CTkEntry(
            wrapper,
            textvariable=variable,
            font=self.font,
            height=36,
            corner_radius=10,
            border_width=0,
            fg_color=COLOR_INPUT,
            text_color=COLOR_TEXT,
        ).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self._tiny_button(wrapper, "浏览...", command, width=70, height=36).grid(row=1, column=1)

    def _tiny_button(self, parent: ctk.CTkFrame, text: str, command, width: int = 54, height: int = 28) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            font=self.font_small,
            width=width,
            height=height,
            corner_radius=10,
            fg_color=COLOR_SECONDARY,
            hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_ACCENT,
            command=command,
        )

    def _secondary_button(self, parent: ctk.CTkFrame, text: str, command) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text=text,
            font=self.font,
            height=36,
            corner_radius=10,
            fg_color=COLOR_SECONDARY,
            hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_ACCENT,
            command=command,
        )

    def select_invoice_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择发票文件夹")
        if folder:
            self.invoice_dir.set(str(Path(folder)))

    def select_save_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择保存文件夹")
        if folder:
            self.save_dir.set(str(Path(folder)))

    def select_all_columns(self) -> None:
        for variable in self.column_vars.values():
            variable.set(True)

    def clear_all_columns(self) -> None:
        for variable in self.column_vars.values():
            variable.set(False)

    def reset_default_columns(self) -> None:
        for column_name, variable in self.column_vars.items():
            variable.set(column_name in UI_DEFAULT_EXPORT_COLUMNS)

    def start_extract(self) -> None:
        if self.is_running:
            return

        input_dir = Path(self.invoice_dir.get()).expanduser()
        save_dir = Path(self.save_dir.get()).expanduser()
        output_name = self.output_name.get().strip() or DEFAULT_OUTPUT_NAME
        if not output_name.lower().endswith(".xlsx"):
            output_name = f"{output_name}.xlsx"
        output_path = save_dir / output_name

        if not input_dir.exists() or not input_dir.is_dir():
            self._append_error("请选择有效的发票文件夹。")
            return

        selected_columns = self.get_selected_columns()
        if not selected_columns:
            self._append_error("请至少选择一个输出字段。")
            return

        self.is_running = True
        self.last_result = None
        self._set_buttons_enabled(False)
        self.progress_bar.set(0)
        self.status_text.set("正在处理")
        self._append_message(f"开始处理：{input_dir}")

        mode = MODE_LABELS.get(self.mode_label.get(), "items")
        self.worker_thread = threading.Thread(
            target=self._run_extract_worker,
            args=(input_dir, output_path, mode, selected_columns),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_extract_worker(self, input_dir: Path, output_path: Path, mode: str, selected_columns: list[str]) -> None:
        try:
            result = process_invoice_folder(
                input_dir,
                output_path,
                mode=mode,  # type: ignore[arg-type]
                selected_columns=selected_columns,
                progress_callback=lambda progress: self.ui_queue.put(("progress", progress)),
            )
            self.ui_queue.put(("result", result))
        except PermissionError as exc:
            self.ui_queue.put(("error", format_exception_message(exc)))
        except Exception as exc:
            self.ui_queue.put(("error", format_exception_message(exc)))
        finally:
            self.ui_queue.put(("worker_done", None))

    def _drain_ui_queue(self) -> None:
        while True:
            try:
                event, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if event == "progress":
                self._handle_progress(payload)
            elif event == "result":
                self.last_result = payload
            elif event == "error":
                self._append_error(str(payload))
            elif event == "worker_done":
                self.is_running = False
                self._set_buttons_enabled(True)
                self.status_text.set("已完成" if self.last_result else "空闲")

        self.after(120, self._drain_ui_queue)

    def _handle_progress(self, progress: BatchProgress) -> None:
        prefix = f"[{progress.index}/{progress.total}] " if progress.total and progress.index else ""
        message = f"{prefix}{progress.message}"
        if progress.total and progress.index:
            self.progress_bar.set(progress.index / progress.total)
        elif progress.event == "done":
            self.progress_bar.set(1)

        if progress.event == "parse_skipped":
            self._append_error(message)
        else:
            self._append_message(message)

    def get_selected_columns(self) -> list[str]:
        selected = {column_name for column_name, variable in self.column_vars.items() if variable.get()}
        return resolve_selected_column_order(selected)

    def preview_data(self) -> None:
        if not self.last_result:
            messagebox.showinfo(APP_TITLE, "暂无可预览数据。")
            return

        lines = []
        for record in self.last_result.successful_records[:10]:
            lines.append(
                f"{record.get('文件名称', '')} | {record.get('发票号码', '')} | "
                f"{record.get('购买方名称', '')} | {record.get('销售方名称', '')}"
            )
        messagebox.showinfo(APP_TITLE, "\n".join(lines) or "暂无成功记录。")

    def show_summary(self) -> None:
        if not self.last_result:
            messagebox.showinfo(APP_TITLE, "暂无汇总数据。")
            return

        message = (
            f"PDF 数量：{self.last_result.total_count}\n"
            f"成功：{self.last_result.success_count}\n"
            f"跳过：{self.last_result.skipped_count}\n"
            f"输出：{self.last_result.output_path}"
        )
        messagebox.showinfo(APP_TITLE, message)

    def open_save_folder(self) -> None:
        folder = Path(self.save_dir.get()).expanduser()
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception as exc:
            self._append_error(f"打开保存文件夹失败：{exc}")

    def _set_buttons_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in self.action_buttons:
            button.configure(state=state)

    def _append_message(self, message: str) -> None:
        self._append_text(self.message_text, message)

    def _append_error(self, message: str) -> None:
        self._append_text(self.error_text, message)

    def _append_text(self, textbox: ctk.CTkTextbox, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        textbox.configure(state="normal")
        textbox.insert("end", f"{timestamp}  {message}\n")
        textbox.see("end")
        textbox.configure(state="disabled")


def resolve_selected_column_order(selected_columns: set[str]) -> list[str]:
    """Keep UI-selected fields in the stable export order users expect."""
    ordered = [column for column in DEFAULT_EXPORT_COLUMNS if column in selected_columns]
    ordered.extend(column for column in AVAILABLE_OUTPUT_COLUMNS if column in selected_columns and column not in ordered)
    return ordered


def format_exception_message(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        filename = getattr(exc, "filename", "") or ""
        if filename:
            return f"写入失败：没有权限访问文件「{filename}」。请先关闭已打开的 Excel 文件，或更换输出文件名后重试。"
        return "写入失败：没有权限访问输出文件。请先关闭已打开的 Excel 文件，或更换输出文件名后重试。"
    return f"处理失败：{exc}"


def main() -> int:
    app = InvoiceExtractorApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
