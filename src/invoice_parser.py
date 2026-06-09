"""PDF invoice parser for native digital invoices.

The parser intentionally avoids OCR and image conversion. It first tries to
read embedded XML files from a PDF, then falls back to native text extraction.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import fitz


INVOICE_COLUMNS = [
    "文件路径",
    "文件名称",
    "发票代码",
    "发票号码",
    "发票类型",
    "开票日期",
    "金额合计",
    "税额合计",
    "税率",
    "价税合计",
    "购买方名称",
    "购买方纳税人识别号",
    "销售方名称",
    "销售方纳税人识别号",
    "备注",
    "项目名称",
]

OPTIONAL_OUTPUT_COLUMNS = [
    "价税合计大写",
    "规格型号",
    "单位",
    "数量",
    "单价",
]

STATUS_OK = "ok"
STATUS_SKIPPED = "skipped"
ERROR_SCANNED_PDF = "SCANNED_PDF"


class InvoiceParserError(Exception):
    """Base exception for parser failures."""


class NotTextPDFError(InvoiceParserError):
    """Raised when a PDF has no useful native text stream."""


def empty_invoice_record(pdf_path: str | Path) -> dict[str, str]:
    """Return a record containing all required Excel columns."""
    path = Path(pdf_path)
    record = {column: "" for column in INVOICE_COLUMNS}
    record["文件路径"] = str(path)
    record["文件名称"] = path.name
    return record


def parse_invoice_pdf(pdf_path: str | Path) -> dict[str, Any]:
    """Parse one PDF invoice and return a standardized dictionary.

    Successful records always contain the required 16 columns plus internal
    metadata keys prefixed with ``_``. Scanned PDFs are returned as skipped
    records so callers can continue processing a batch safely.
    """
    path = Path(pdf_path)
    record = empty_invoice_record(path)

    try:
        with fitz.open(path) as document:
            text = _extract_native_text(document)
            if len(_compact_text(text)) < 20:
                raise NotTextPDFError("PDF has no useful native text stream")

            record.update(_parse_text_invoice(text))
            visual_fields = _parse_visual_invoice(document)
            for key, value in visual_fields.items():
                if value:
                    record[key] = value
            items = _extract_line_items_from_pages(document)
            if items:
                record["_items"] = items
                record["项目名称"] = "\n".join(dict.fromkeys(item["项目名称"] for item in items if item.get("项目名称")))
                if not record.get("税率"):
                    record["税率"] = "；".join(dict.fromkeys(item["税率"] for item in items if item.get("税率")))
            record["_status"] = STATUS_OK
            record["_source"] = "visual_text"
            return record
    except NotTextPDFError as exc:
        record["_status"] = STATUS_SKIPPED
        record["_error_code"] = ERROR_SCANNED_PDF
        record["_error_message"] = str(exc)
        return record
    except Exception as exc:  # Keep batch processing alive for malformed files.
        record["_status"] = STATUS_SKIPPED
        record["_error_code"] = exc.__class__.__name__
        record["_error_message"] = str(exc)
        return record


def _extract_embedded_xml(document: fitz.Document) -> bytes | None:
    count = document.embfile_count()
    for index in range(count):
        info = document.embfile_info(index) or {}
        names = [
            str(info.get("filename") or ""),
            str(info.get("ufilename") or ""),
            str(info.get("name") or ""),
        ]
        if any(name.lower().endswith(".xml") for name in names):
            payload = document.embfile_get(index)
            if payload:
                return bytes(payload)
    return None


def _extract_native_text(document: fitz.Document) -> str:
    page_text = []
    for page in document:
        page_text.append(page.get_text("text"))
    return "\n".join(page_text)


def _parse_visual_invoice(document: fitz.Document) -> dict[str, str]:
    if document.page_count == 0:
        return {}

    first_page = document[0]
    first_words = first_page.get_text("words")
    if not first_words:
        return {}

    fields: dict[str, str] = {}
    fields["发票类型"] = _find_visual_invoice_type(first_page)
    fields["发票号码"] = _find_visual_invoice_number(document)
    fields["开票日期"] = _find_visual_issue_date(document)

    party_fields = _find_visual_parties(first_page)
    fields.update(party_fields)

    total_fields = _find_visual_totals(document)
    fields.update(total_fields)
    return {key: value for key, value in fields.items() if value}


def _find_visual_invoice_type(page: fitz.Page) -> str:
    rows = _group_words_by_row(page.get_text("words"))
    page_height = float(page.rect.height)
    candidates: list[tuple[float, str]] = []
    for row_y, row_words in rows:
        text = _row_text(row_words)
        if "发票" not in text:
            continue
        if "号码" in text or "代码" in text or "开票" in text:
            continue
        if row_y > page_height * 0.25:
            continue
        candidates.append((row_y, text))
    return candidates[0][1] if candidates else ""


def _find_visual_invoice_number(document: fitz.Document) -> str:
    candidates = _find_word_regex_candidates(document, r"\b\d{20}\b")
    top_right = [
        candidate
        for candidate in candidates
        if candidate["page"] == 0 and candidate["y_ratio"] < 0.22 and candidate["x_ratio"] > 0.55
    ]
    if top_right:
        return str(top_right[0]["value"])
    return str(candidates[0]["value"]) if candidates else ""


def _find_visual_issue_date(document: fitz.Document) -> str:
    candidates = _find_word_regex_candidates(document, r"\b20\d{2}年\d{1,2}月\d{1,2}日\b")
    top_right = [
        candidate
        for candidate in candidates
        if candidate["page"] == 0 and candidate["y_ratio"] < 0.22 and candidate["x_ratio"] > 0.55
    ]
    value = str(top_right[0]["value"] if top_right else candidates[0]["value"]) if candidates else ""
    return _normalize_date(value)


def _find_word_regex_candidates(document: fitz.Document, pattern: str) -> list[dict[str, Any]]:
    regex = re.compile(pattern)
    candidates: list[dict[str, Any]] = []
    for page_index, page in enumerate(document):
        width = float(page.rect.width) or 1.0
        height = float(page.rect.height) or 1.0
        for word in page.get_text("words"):
            x0, y0, _x1, _y1, text, *_ = word
            for match in regex.finditer(str(text)):
                candidates.append(
                    {
                        "value": match.group(0),
                        "page": page_index,
                        "x": float(x0),
                        "y": float(y0),
                        "x_ratio": float(x0) / width,
                        "y_ratio": float(y0) / height,
                    }
                )
    return sorted(candidates, key=lambda item: (item["page"], item["y"], item["x"]))


def _find_visual_parties(page: fitz.Page) -> dict[str, str]:
    words = page.get_text("words")
    if not words:
        return {}

    width = float(page.rect.width)
    height = float(page.rect.height)
    rows = _group_words_by_row(words)
    header_y = _find_row_y_containing(rows, "项目名称")
    bottom = header_y if header_y is not None else height * 0.42
    top = height * 0.14

    buyer_words = _filter_words(words, 0, width * 0.5, top, bottom)
    seller_words = _filter_words(words, width * 0.5, width, top, bottom)
    buyer = _parse_party_region(buyer_words)
    seller = _parse_party_region(seller_words)

    fields: dict[str, str] = {}
    if buyer.get("name"):
        fields["购买方名称"] = buyer["name"]
    if buyer.get("tax_no"):
        fields["购买方纳税人识别号"] = buyer["tax_no"]
    if seller.get("name"):
        fields["销售方名称"] = seller["name"]
    if seller.get("tax_no"):
        fields["销售方纳税人识别号"] = seller["tax_no"]
    return fields


def _filter_words(words: list[tuple], left: float, right: float, top: float, bottom: float) -> list[tuple]:
    return [
        word
        for word in words
        if left <= float(word[0]) < right and top <= float(word[1]) <= bottom
    ]


def _parse_party_region(words: list[tuple]) -> dict[str, str]:
    rows = _group_words_by_row(words)
    row_texts = [_row_text(row_words) for _, row_words in rows]
    joined = "\n".join(row_texts)
    result = {"name": "", "tax_no": ""}

    for text in row_texts:
        match = re.search(r"名称[:：]\s*(.+)", text)
        if match and match.group(1).strip():
            result["name"] = _trim_value(match.group(1))
            break

    tax_match = re.search(r"\b9[0-9A-Z]{17}\b", joined)
    if tax_match:
        result["tax_no"] = tax_match.group(0)
    return result


def _find_visual_totals(document: fitz.Document) -> dict[str, str]:
    fields: dict[str, str] = {}
    money_pattern = r"[¥￥]\s*(-?[0-9]+(?:\.[0-9]{1,2})?)"
    upper_pattern = r"[零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+"

    for page in reversed(list(document)):
        rows = _group_words_by_row(page.get_text("words"))
        for _row_y, row_words in reversed(rows):
            text = _row_text(row_words)
            if "价税合计" in text or "小写" in text:
                money_values = re.findall(money_pattern, text)
                if money_values and not fields.get("价税合计"):
                    fields["价税合计"] = _normalize_money(money_values[-1])
                upper_match = re.search(upper_pattern, text)
                if upper_match and not fields.get("价税合计大写"):
                    fields["价税合计大写"] = upper_match.group(0)

            if "合计" in text and "价税合计" not in text:
                money_values = re.findall(money_pattern, text)
                if len(money_values) >= 2:
                    fields["金额合计"] = _normalize_money(money_values[-2])
                    fields["税额合计"] = _normalize_money(money_values[-1])

            if fields.get("金额合计") and fields.get("税额合计") and fields.get("价税合计") and fields.get("价税合计大写"):
                return fields
    return fields


def _parse_xml_invoice(xml_payload: bytes) -> dict[str, str]:
    root = ElementTree.fromstring(xml_payload)
    flattened = _flatten_xml(root)

    def first(*names: str) -> str:
        for name in names:
            value = _lookup_xml_value(flattened, name)
            if value:
                return value
        return ""

    fields = {
        "发票代码": first("InvoiceCode", "Fpdm", "FPDM", "发票代码"),
        "发票号码": first("InvoiceNumber", "Fphm", "FPHM", "发票号码", "发票号码/数电票号码"),
        "发票类型": first("InvoiceType", "Fplx", "FPLX", "发票类型"),
        "开票日期": _normalize_date(first("IssueDate", "Kprq", "KPRQ", "开票日期")),
        "金额合计": _normalize_money(first("TotalAmount", "Hjje", "HJJE", "金额合计", "合计金额")),
        "税额合计": _normalize_money(first("TotalTax", "Hjse", "HJSE", "税额合计", "合计税额")),
        "税率": first("TaxRate", "Sl", "SL", "税率"),
        "价税合计": _normalize_money(first("TotalTaxIncludedAmount", "Jshj", "JSHJ", "价税合计")),
        "价税合计大写": first("TotalTaxIncludedAmountCN", "JshjDx", "JSHJDX", "价税合计大写"),
        "购买方名称": first("BuyerName", "Gfmc", "GFMC", "购买方名称"),
        "购买方纳税人识别号": first("BuyerTaxNo", "Gfnsrsbh", "GFNSRSBH", "购买方纳税人识别号"),
        "销售方名称": first("SellerName", "Xfmc", "XFMC", "销售方名称"),
        "销售方纳税人识别号": first("SellerTaxNo", "Xfnsrsbh", "XFNSRSBH", "销售方纳税人识别号"),
        "备注": first("Remark", "Bz", "BZ", "备注"),
        "项目名称": _join_items_from_xml(flattened),
    }
    return _clean_required_fields(fields)


def _flatten_xml(root: ElementTree.Element) -> dict[str, list[str]]:
    values: dict[str, list[str]] = {}
    for element in root.iter():
        tag = _strip_namespace(element.tag)
        text = (element.text or "").strip()
        if text:
            values.setdefault(tag, []).append(text)
    return values


def _strip_namespace(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _lookup_xml_value(flattened: dict[str, list[str]], name: str) -> str:
    lowered_name = name.lower()
    for key, values in flattened.items():
        if key == name or key.lower() == lowered_name:
            return values[0].strip()
    return ""


def _join_items_from_xml(flattened: dict[str, list[str]]) -> str:
    item_keys = (
        "ItemName",
        "ProjectName",
        "Xmmc",
        "XMMC",
        "货物或应税劳务、服务名称",
        "项目名称",
    )
    items: list[str] = []
    for key in item_keys:
        items.extend(value.strip() for value in flattened.get(key, []) if value.strip())
    return "；".join(dict.fromkeys(items))


def _parse_text_invoice(text: str) -> dict[str, str]:
    normalized = _normalize_text(text)
    fields = {
        "发票代码": _search(normalized, [r"发票代码[:：]?\s*([0-9]{8,20})"]),
        "发票号码": _search(
            normalized,
            [
                r"(?:发票号码|数电票号码)[:：]?\s*([0-9]{8,30})",
                r"发票号码/数电票号码[:：]?\s*([0-9]{8,30})",
            ],
        ),
        "发票类型": _detect_invoice_type(normalized),
        "开票日期": _normalize_date(
            _search(normalized, [r"开票日期[:：]?\s*([0-9]{4}[年/-][0-9]{1,2}[月/-][0-9]{1,2}日?)"])
        ),
        "金额合计": _normalize_money(
            _search(normalized, [r"(?:金额合计|合计金额)[:：]?\s*[¥￥]?\s*([0-9,]+(?:\.[0-9]{1,2})?)"])
        ),
        "税额合计": _normalize_money(
            _search(normalized, [r"(?:税额合计|合计税额)[:：]?\s*[¥￥]?\s*([0-9,]+(?:\.[0-9]{1,2})?)"])
        ),
        "税率": _search(normalized, [r"税率[:：]?\s*([0-9]+(?:\.[0-9]+)?%)"]),
        "价税合计": _normalize_money(
            _search(
                normalized,
                [
                    r"价税合计(?:\s*[(（]小写[)）])?[:：]?\s*[¥￥]?\s*([0-9,]+(?:\.[0-9]{1,2})?)",
                    r"小写[:：]?\s*[¥￥]?\s*([0-9,]+(?:\.[0-9]{1,2})?)",
                ],
            )
        ),
        "价税合计大写": _search(
            normalized,
            [
                r"价税合计(?:\s*[(（]大写[)）])?[:：]?\s*([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+)",
                r"大写[:：]?\s*([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+)",
            ],
        ),
        "购买方名称": _search_party_field(normalized, "购买方", "名称"),
        "购买方纳税人识别号": _search_party_field(normalized, "购买方", "纳税人识别号"),
        "销售方名称": _search_party_field(normalized, "销售方", "名称"),
        "销售方纳税人识别号": _search_party_field(normalized, "销售方", "纳税人识别号"),
        "备注": _search(normalized, [r"备注[:：]?\s*(.*?)(?:\s+(?:开票人|收款人|复核人)[:：]|$)"]),
        "项目名称": _extract_item_names(normalized),
    }
    _fill_from_stacked_layout(normalized, fields)
    _fill_from_labeled_lines(normalized, fields)
    _fill_summary_totals_from_text(normalized, fields)
    return _clean_required_fields(fields)


def _normalize_text(text: str) -> str:
    lines = [line.strip() for line in text.replace("\u3000", " ").splitlines()]
    return "\n".join(line for line in lines if line)


def _compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text)


def _search(text: str, patterns: list[str]) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            return _trim_value(match.group(1))
    return ""


def _search_party_field(text: str, party: str, field: str) -> str:
    pattern = rf"{party}(?:信息)?[\s\S]{{0,160}}?{field}[:：]?\s*([^\n]+)"
    value = _search(text, [pattern])
    return _trim_value(value)


def _extract_item_names(text: str) -> str:
    items: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("*"):
            continue
        item = re.split(r"\s{2,}|[\t]", stripped, maxsplit=1)[0]
        if item:
            items.append(item)
    return "；".join(dict.fromkeys(items))


def _fill_from_stacked_layout(text: str, fields: dict[str, str]) -> None:
    """Fill fields for PDFs whose text layer emits labels before values.

    Some native invoice PDFs extract as a block of static labels followed by a
    block of dynamic values. The visible page is correct, but the text stream is
    not in reading order, so label-adjacent regexes cannot see the values.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return

    invoice_index = _find_line_index(lines, r"^[0-9]{8,30}$")
    date_index = _find_line_index(lines, r"^[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日$")

    if not fields.get("发票号码") and invoice_index is not None:
        fields["发票号码"] = lines[invoice_index]
    if not fields.get("开票日期") and date_index is not None:
        fields["开票日期"] = _normalize_date(lines[date_index])

    if date_index is not None:
        _fill_parties_from_value_block(lines, date_index + 1, fields)
        _fill_amounts_from_value_block(lines, date_index + 1, fields)

    if not fields.get("备注"):
        remark = next((line for line in lines if line.startswith(("备注:", "备注：", "保单号:", "保单号："))), "")
        fields["备注"] = remark


def _fill_from_labeled_lines(text: str, fields: dict[str, str]) -> None:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return

    invoice_number = _next_value_after_label(lines, r"^(?:发票号码|数电票号码)[:：]?$", r"^[0-9]{8,30}$")
    issue_date = _next_value_after_label(lines, r"^开票日期[:：]?$", r"^[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日$")
    if invoice_number:
        fields["发票号码"] = invoice_number
    if issue_date:
        fields["开票日期"] = _normalize_date(issue_date)

    names = [
        value
        for value in (_trim_value(line.split("：", 1)[-1].split(":", 1)[-1]) for line in lines if line.startswith("名称"))
        if value
    ]
    tax_numbers = [
        value
        for value in (
            _trim_value(line.split("：", 1)[-1].split(":", 1)[-1])
            for line in lines
            if line.startswith("统一社会信用代码/纳税人识别号")
        )
        if value
    ]
    if len(names) >= 2:
        fields["购买方名称"] = names[0]
        fields["销售方名称"] = names[1]
    if len(tax_numbers) >= 2:
        fields["购买方纳税人识别号"] = tax_numbers[0]
        fields["销售方纳税人识别号"] = tax_numbers[1]


def _next_value_after_label(lines: list[str], label_pattern: str, value_pattern: str, max_lookahead: int = 8) -> str:
    label_regex = re.compile(label_pattern)
    value_regex = re.compile(value_pattern)
    for index, line in enumerate(lines):
        if not label_regex.match(line):
            continue
        for candidate in lines[index + 1 : index + 1 + max_lookahead]:
            if value_regex.match(candidate):
                return candidate
    return ""


def _fill_summary_totals_from_text(text: str, fields: dict[str, str]) -> None:
    compact = re.sub(r"\s+", "", text)
    money_pattern = r"[¥￥]\s*(-?[0-9]+(?:\.[0-9]{1,2})?)"

    total_match = re.search(r"小写[)）]?" + money_pattern, compact)
    if total_match:
        fields["价税合计"] = _normalize_money(total_match.group(1))

    subtotal_match = re.search(r"(?<!价税)合计" + money_pattern + money_pattern, compact)
    if subtotal_match:
        fields["金额合计"] = _normalize_money(subtotal_match.group(1))
        fields["税额合计"] = _normalize_money(subtotal_match.group(2))

    upper_match = re.search(r"价税合计[（(]大写[）)]([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+)", compact)
    if upper_match:
        fields["价税合计大写"] = upper_match.group(1)


def _fill_parties_from_value_block(lines: list[str], start: int, fields: dict[str, str]) -> None:
    tax_pattern = re.compile(r"^[0-9A-Z]{15,20}$")
    tax_positions = [index for index in range(start, len(lines)) if tax_pattern.match(lines[index])]
    if len(tax_positions) < 2:
        return

    buyer_tax_index, seller_tax_index = tax_positions[:2]
    buyer_name = _previous_value_line(lines, buyer_tax_index, start)
    seller_name = _previous_value_line(lines, seller_tax_index, buyer_tax_index + 1)

    if buyer_name and not fields.get("购买方名称"):
        fields["购买方名称"] = buyer_name
    if not fields.get("购买方纳税人识别号"):
        fields["购买方纳税人识别号"] = lines[buyer_tax_index]
    if seller_name and not fields.get("销售方名称"):
        fields["销售方名称"] = seller_name
    if not fields.get("销售方纳税人识别号"):
        fields["销售方纳税人识别号"] = lines[seller_tax_index]


def _fill_amounts_from_value_block(lines: list[str], start: int, fields: dict[str, str]) -> None:
    tax_rate_index = _find_line_index(lines[start:], r"^[0-9]+(?:\.[0-9]+)?%$")
    if tax_rate_index is None:
        return
    tax_rate_index += start

    if not fields.get("税率"):
        fields["税率"] = lines[tax_rate_index]

    number_values_before_tax = [
        line for line in lines[start:tax_rate_index] if re.fullmatch(r"[0-9]+(?:\.[0-9]{1,2})?", line)
    ]
    if number_values_before_tax and not fields.get("金额合计"):
        fields["金额合计"] = number_values_before_tax[-1]

    if tax_rate_index + 1 < len(lines) and not fields.get("税额合计"):
        tax_value = _normalize_money(lines[tax_rate_index + 1])
        if tax_value:
            fields["税额合计"] = tax_value

    money_lines = [_normalize_money(line) for line in lines[tax_rate_index + 1 :] if line.startswith(("¥", "￥"))]
    money_lines = [line for line in money_lines if line]
    if money_lines:
        if not fields.get("金额合计"):
            fields["金额合计"] = money_lines[0]
        if len(money_lines) >= 2 and not fields.get("税额合计"):
            fields["税额合计"] = money_lines[1]
        if not fields.get("价税合计"):
            fields["价税合计"] = money_lines[-1]

    if not fields.get("价税合计大写"):
        upper_value = next(
            (
                line
                for line in lines[tax_rate_index + 1 :]
                if re.fullmatch(r"[零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+", line)
            ),
            "",
        )
        fields["价税合计大写"] = upper_value


def _extract_line_items_from_pages(document: fitz.Document) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    for page in document:
        items.extend(_extract_line_items_from_page(page))
    return items


def _extract_line_items_from_page(page: fitz.Page) -> list[dict[str, str]]:
    words = page.get_text("words")
    if not words:
        return []

    rows = _group_words_by_row(words)
    header_y = _find_row_y_containing(rows, "项目名称")
    total_y = _find_row_y_containing(rows, "合计")
    if header_y is None:
        return []
    if total_y is None:
        total_y = page.rect.height

    detail_rows = [
        (y, row_words)
        for y, row_words in rows
        if header_y + 3 < y < total_y - 3 and _row_text(row_words)
    ]
    return _parse_detail_rows(detail_rows)


def _group_words_by_row(words: list[tuple]) -> list[tuple[float, list[tuple]]]:
    sorted_words = sorted(words, key=lambda word: (round(float(word[1]), 1), float(word[0])))
    rows: list[tuple[float, list[tuple]]] = []
    for word in sorted_words:
        y0 = float(word[1])
        for row_index, (row_y, row_words) in enumerate(rows):
            if abs(row_y - y0) <= 3:
                row_words.append(word)
                rows[row_index] = ((row_y + y0) / 2, row_words)
                break
        else:
            rows.append((y0, [word]))
    return [(y, sorted(row_words, key=lambda word: float(word[0]))) for y, row_words in rows]


def _find_row_y_containing(rows: list[tuple[float, list[tuple]]], text: str) -> float | None:
    for y, row_words in rows:
        if text in _row_text(row_words):
            return y
    return None


def _row_text(row_words: list[tuple]) -> str:
    return "".join(str(word[4]) for word in row_words).strip()


def _parse_detail_rows(detail_rows: list[tuple[float, list[tuple]]]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    current: dict[str, str] | None = None

    for _, row_words in detail_rows:
        row = _split_detail_row(row_words)
        if not row["项目名称"]:
            continue
        starts_new_item = bool(row["税率"] or row["金额"] or row["税额"] or row["规格型号"] or row["数量"] or row["单价"])
        if starts_new_item:
            current = row
            items.append(current)
        elif current is not None:
            current["项目名称"] = "\n".join(part for part in [current["项目名称"], row["项目名称"]] if part)
        else:
            current = row
            items.append(current)

    return items


def _split_detail_row(row_words: list[tuple]) -> dict[str, str]:
    columns = {
        "项目名称": (0, 115),
        "规格型号": (115, 185),
        "单位": (185, 235),
        "数量": (235, 310),
        "单价": (310, 385),
        "金额": (385, 445),
        "税率": (445, 520),
        "税额": (520, 620),
    }
    row = {key: "" for key in columns}
    for key, (left, right) in columns.items():
        parts = [str(word[4]) for word in row_words if left <= float(word[0]) < right]
        row[key] = "".join(parts).strip()
    row["项目名称"], inferred_spec = _split_item_name_and_spec(row["项目名称"])
    if inferred_spec and not row["规格型号"]:
        row["规格型号"] = inferred_spec
    return row


def _split_item_name_and_spec(value: str) -> tuple[str, str]:
    value = value.strip()
    match = re.search(r"(\*{2,})$", value)
    if not match:
        return value, ""
    spec = match.group(1)
    return value[: -len(spec)].strip(), spec


def _previous_value_line(lines: list[str], index: int, lower_bound: int) -> str:
    for cursor in range(index - 1, lower_bound - 1, -1):
        line = lines[cursor]
        if _looks_like_static_label(line):
            continue
        return line
    return ""


def _looks_like_static_label(line: str) -> bool:
    labels = {
        "名称：",
        "名称:",
        "统一社会信用代码/纳税人识别号：",
        "统一社会信用代码/纳税人识别号:",
        "项目名称",
        "规格型号",
        "单 位",
        "数 量",
        "单 价",
        "金 额",
        "税率/征收率",
        "税  额",
        "合",
        "计",
        "备",
        "注",
    }
    return line in labels


def _find_line_index(lines: list[str], pattern: str) -> int | None:
    regex = re.compile(pattern)
    for index, line in enumerate(lines):
        if regex.match(line):
            return index
    return None


def _detect_invoice_type(text: str) -> str:
    explicit = _search(text, [r"发票类型[:：]?\s*([^\n]+)"])
    if explicit:
        return explicit
    for line in text.splitlines():
        if "发票" in line:
            return _trim_value(line)
    return ""


def _normalize_money(value: str) -> str:
    value = _trim_value(value).replace(",", "")
    if not value:
        return ""
    match = re.search(r"-?[0-9]+(?:\.[0-9]{1,2})?", value)
    return match.group(0) if match else value


def _normalize_date(value: str) -> str:
    value = _trim_value(value)
    match = re.search(r"([0-9]{4})[年/-]([0-9]{1,2})[月/-]([0-9]{1,2})日?", value)
    if not match:
        return value
    year, month, day = match.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def _trim_value(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip()
    return value.strip("：:，,;；")


def _clean_required_fields(fields: dict[str, Any]) -> dict[str, str]:
    output_columns = [*INVOICE_COLUMNS, *OPTIONAL_OUTPUT_COLUMNS]
    return {
        column: str(fields.get(column, "") or "").strip()
        for column in output_columns
        if column in fields
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse one native digital invoice PDF.")
    parser.add_argument("pdf_path", help="Path to a PDF invoice")
    args = parser.parse_args(argv)

    result = parse_invoice_pdf(args.pdf_path)
    json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return 0 if result.get("_status") == STATUS_OK else 1


if __name__ == "__main__":
    raise SystemExit(main())
