from pathlib import Path

import fitz

from src.invoice_parser import (
    ERROR_SCANNED_PDF,
    INVOICE_COLUMNS,
    STATUS_OK,
    STATUS_SKIPPED,
    _parse_text_invoice,
    parse_invoice_pdf,
)


def test_parse_pdf_with_embedded_xml(tmp_path: Path) -> None:
    pdf_path = tmp_path / "xml_invoice.pdf"
    xml = """<?xml version="1.0" encoding="UTF-8"?>
<Invoice>
  <InvoiceCode>044001900111</InvoiceCode>
  <InvoiceNumber>24992000000000012345</InvoiceNumber>
  <InvoiceType>电子发票（普通发票）</InvoiceType>
  <IssueDate>2026年06月09日</IssueDate>
  <TotalAmount>100.00</TotalAmount>
  <TotalTax>6.00</TotalTax>
  <TaxRate>6%</TaxRate>
  <TotalTaxIncludedAmount>106.00</TotalTaxIncludedAmount>
  <BuyerName>测试购买方有限公司</BuyerName>
  <BuyerTaxNo>91310000BUYER001</BuyerTaxNo>
  <SellerName>测试销售方有限公司</SellerName>
  <SellerTaxNo>91310000SELLER001</SellerTaxNo>
  <Remark>XML备注</Remark>
  <ItemName>*信息技术服务*软件服务</ItemName>
</Invoice>
""".encode("utf-8")

    doc = fitz.open()
    doc.new_page()
    doc.embfile_add("invoice.xml", xml, filename="invoice.xml")
    doc.save(pdf_path)
    doc.close()

    result = parse_invoice_pdf(pdf_path)

    assert set(INVOICE_COLUMNS).issubset(result.keys())
    assert result["_status"] == STATUS_SKIPPED
    assert result["_error_code"] == ERROR_SCANNED_PDF
    assert result["文件路径"] == str(pdf_path)
    assert result["文件名称"] == "xml_invoice.pdf"
    assert result["发票号码"] == ""


def test_parse_native_text_invoice() -> None:
    text = """
电子发票（普通发票）
发票代码：044001900222
发票号码：24992000000000054321
开票日期：2026年06月09日
购买方信息
名称：文本购买方有限公司
纳税人识别号：91310000TEXTBUYER
销售方信息
名称：文本销售方有限公司
纳税人识别号：91310000TEXTSELLER
项目名称 税率 金额 税额
*现代服务*咨询服务 6% 200.00 12.00
金额合计：￥200.00
税额合计：￥12.00
税率：6%
价税合计（小写）：￥212.00
备注：文本备注 开票人：张三
"""

    result = _parse_text_invoice(text)

    assert result["发票号码"] == "24992000000000054321"
    assert result["开票日期"] == "2026-06-09"
    assert result["金额合计"] == "200.00"
    assert result["税额合计"] == "12.00"
    assert result["价税合计"] == "212.00"
    assert result["购买方名称"] == "文本购买方有限公司"
    assert result["销售方名称"] == "文本销售方有限公司"


def test_scanned_like_pdf_is_skipped(tmp_path: Path) -> None:
    pdf_path = tmp_path / "empty_scan_like.pdf"

    doc = fitz.open()
    doc.new_page()
    doc.save(pdf_path)
    doc.close()

    result = parse_invoice_pdf(pdf_path)

    assert set(INVOICE_COLUMNS).issubset(result.keys())
    assert result["_status"] == STATUS_SKIPPED
    assert result["_error_code"] == ERROR_SCANNED_PDF
