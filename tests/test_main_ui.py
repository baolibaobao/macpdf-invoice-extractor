from src.file_manager import DEFAULT_EXPORT_COLUMNS
from src.main_ui import DEFAULT_OUTPUT_NAME, MODE_LABELS, resolve_selected_column_order


def test_main_ui_defaults() -> None:
    assert DEFAULT_OUTPUT_NAME.endswith(".xlsx")
    assert MODE_LABELS["汇总+明细"] == "items"
    assert MODE_LABELS["发票一行"] == "invoice"


def test_selected_columns_keep_default_export_order() -> None:
    selected = {"税率", "文件名称", "购买方名称", "明细项目名称", "发票号码", "价税合计", "发票代码"}

    columns = resolve_selected_column_order(selected)

    assert columns[:6] == ["文件名称", "发票号码", "购买方名称", "明细项目名称", "税率", "价税合计"]
    assert columns[-1] == "发票代码"
    assert [column for column in DEFAULT_EXPORT_COLUMNS if column in selected] == columns[:-1]
