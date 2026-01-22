"""Tests for doc_code-only de-duplication (Phase 4.2)."""

from openpyxl import Workbook

from app.core.models import Product
from app.parser.workbook import _dedupe_products_by_doc_code, parse_workbook


def test_dedupe_products_by_doc_code_keeps_first_occurrence() -> None:
    products = [
        Product(doc_code="A1", product_name="first"),
        Product(doc_code="A1", product_name="second"),
        Product(doc_code="A1 ", product_name="third-with-trailing-space"),
        Product(doc_code="B2", product_name="unique"),
    ]

    deduped = _dedupe_products_by_doc_code(products)

    assert [p.doc_code for p in deduped] == ["A1", "B2"]
    assert deduped[0].product_name == "first"


def test_dedupe_products_by_doc_code_keeps_all_missing_doc_codes() -> None:
    products = [
        Product(doc_code=None, product_name="missing-1"),
        Product(doc_code="", product_name="missing-empty"),
        Product(doc_code="   ", product_name="missing-whitespace"),
        Product(doc_code=None, product_name="missing-2"),
    ]

    deduped = _dedupe_products_by_doc_code(products)

    assert len(deduped) == 4
    assert [p.product_name for p in deduped] == [
        "missing-1",
        "missing-empty",
        "missing-whitespace",
        "missing-2",
    ]


def test_parse_workbook_applies_doc_code_dedup_across_sheets() -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Schedule"
    ws2 = wb.create_sheet("Sales Schedule")

    # Minimal headers that the sheet detector / column mapper can understand.
    headers = ["CODE", "DESCRIPTION"]
    ws1.append(headers)
    ws2.append(headers)

    ws1.append(["DUP1", "First occurrence"])
    ws2.append(["DUP1", "Second occurrence (should be dropped)"])
    ws2.append(["UNIQ", "Unique product"])
    ws2.append([None, "No doc_code 1"])
    ws2.append([None, "No doc_code 2"])

    parsed = parse_workbook(wb, filename="test.xlsx")
    doc_codes = [p.doc_code for p in parsed.products]

    assert parsed.schedule_name
    assert doc_codes == ["DUP1", "UNIQ", None, None]

