"""Manual helper to export real invoice fixtures to Excel."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.excel_builder import export_invoices_to_excel
from src.invoice_parser import parse_invoice_pdf


DEFAULT_COLUMNS = [
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Export real invoice fixture PDFs to Excel.")
    parser.add_argument(
        "--input-dir",
        default=str(Path("tests") / "fixtures" / "real_invoices"),
        help="Folder containing real PDF invoices.",
    )
    parser.add_argument(
        "--output",
        default=str(Path("outputs") / "real_invoice_items_export_test.xlsx"),
        help="Output Excel path.",
    )
    parser.add_argument("--mode", choices=["invoice", "items"], default="items")
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    records = [parse_invoice_pdf(path) for path in sorted(input_dir.glob("*.pdf"))]
    output_path = export_invoices_to_excel(records, args.output, mode=args.mode, selected_columns=DEFAULT_COLUMNS)

    print(f"exported: {output_path}")
    print(f"pdf_count: {len(records)}")
    for record in records:
        print(f"{record.get('文件名称', '')}: {record.get('_status', '')}, items={len(record.get('_items', []))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
