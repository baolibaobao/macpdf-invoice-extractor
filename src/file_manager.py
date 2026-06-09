"""Batch file scanning and invoice export orchestration."""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from src.excel_builder import ExportMode, export_invoices_to_excel
from src.invoice_parser import parse_invoice_pdf


DEFAULT_EXPORT_COLUMNS = [
    "文件路径",
    "文件名称",
    "发票号码",
    "发票类型",
    "开票日期",
    "购买方名称",
    "购买方纳税人识别号",
    "销售方名称",
    "销售方纳税人识别号",
    "项目名称",
    "明细项目名称",
    "规格型号",
    "单位",
    "数量",
    "单价",
    "税率",
    "明细金额",
    "明细税额",
    "金额合计",
    "税额合计",
    "价税合计",
    "价税合计大写",
    "备注",
]


@dataclass(frozen=True)
class BatchProgress:
    event: str
    message: str
    index: int = 0
    total: int = 0
    pdf_path: Path | None = None


@dataclass
class BatchResult:
    input_dir: Path
    output_path: Path
    pdf_files: list[Path] = field(default_factory=list)
    records: list[dict[str, Any]] = field(default_factory=list)
    successful_records: list[dict[str, Any]] = field(default_factory=list)
    skipped_records: list[dict[str, Any]] = field(default_factory=list)

    @property
    def total_count(self) -> int:
        return len(self.pdf_files)

    @property
    def success_count(self) -> int:
        return len(self.successful_records)

    @property
    def skipped_count(self) -> int:
        return len(self.skipped_records)


ProgressCallback = Callable[[BatchProgress], None]


def find_pdf_files(input_dir: str | Path) -> list[Path]:
    """Recursively find PDF files in a folder, case-insensitively."""
    root = Path(input_dir).expanduser()
    if not root.exists():
        raise FileNotFoundError(f"Input folder does not exist: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {root}")

    return sorted(path for path in root.rglob("*") if path.is_file() and path.suffix.lower() == ".pdf")


def process_invoice_folder(
    input_dir: str | Path,
    output_path: str | Path,
    mode: ExportMode = "items",
    selected_columns: list[str] | None = None,
    progress_callback: ProgressCallback | None = None,
) -> BatchResult:
    """Parse all PDF invoices under ``input_dir`` and export an Excel file."""
    source_dir = Path(input_dir).expanduser()
    target_path = Path(output_path).expanduser()
    columns = selected_columns if selected_columns is not None else DEFAULT_EXPORT_COLUMNS

    _emit(progress_callback, "scan", f"正在扫描文件夹：{source_dir}")
    pdf_files = find_pdf_files(source_dir)
    result = BatchResult(input_dir=source_dir, output_path=target_path, pdf_files=pdf_files)
    _emit(progress_callback, "scan_done", f"发现 {len(pdf_files)} 个 PDF 文件。", total=len(pdf_files))

    for index, pdf_path in enumerate(pdf_files, start=1):
        _emit(progress_callback, "parse_start", f"正在解析：{pdf_path.name}", index=index, total=len(pdf_files), pdf_path=pdf_path)
        record = parse_invoice_pdf(pdf_path)
        result.records.append(record)

        if record.get("_status") == "ok":
            result.successful_records.append(record)
            item_count = len(record.get("_items", []))
            _emit(
                progress_callback,
                "parse_ok",
                f"解析成功：{pdf_path.name}（{item_count} 条明细）",
                index=index,
                total=len(pdf_files),
                pdf_path=pdf_path,
            )
        else:
            result.skipped_records.append(record)
            error_code = record.get("_error_code", "UNKNOWN")
            _emit(
                progress_callback,
                "parse_skipped",
                f"已跳过：{pdf_path.name}（{error_code}）",
                index=index,
                total=len(pdf_files),
                pdf_path=pdf_path,
            )

    _emit(progress_callback, "export_start", f"正在导出 Excel：{target_path}", total=len(pdf_files))
    export_invoices_to_excel(result.records, target_path, mode=mode, selected_columns=columns)
    _emit(
        progress_callback,
        "done",
        f"处理完成。成功 {result.success_count} 个，跳过 {result.skipped_count} 个，输出文件：{target_path}",
        total=len(pdf_files),
    )
    return result


def _emit(
    progress_callback: ProgressCallback | None,
    event: str,
    message: str,
    index: int = 0,
    total: int = 0,
    pdf_path: Path | None = None,
) -> None:
    if progress_callback is None:
        return
    progress_callback(BatchProgress(event=event, message=message, index=index, total=total, pdf_path=pdf_path))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Batch parse invoice PDFs and export Excel.")
    parser.add_argument("input_dir", help="Folder containing invoice PDFs")
    parser.add_argument("--output", required=True, help="Output .xlsx path")
    parser.add_argument("--mode", choices=["invoice", "items"], default="items")
    parser.add_argument("--columns", nargs="*", help="Optional output columns")
    args = parser.parse_args(argv)

    def print_progress(progress: BatchProgress) -> None:
        prefix = f"[{progress.index}/{progress.total}] " if progress.total and progress.index else ""
        print(f"{prefix}{progress.message}")

    process_invoice_folder(
        args.input_dir,
        args.output,
        mode=args.mode,
        selected_columns=args.columns,
        progress_callback=print_progress,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
