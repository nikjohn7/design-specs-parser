"""Synthetic tests for mutated schedules (Phase 5.6).

These tests validate that the parser is robust to realistic mutations applied to
real sample schedules via `tools/generate_programa_test_schedules.py`.

Mutated schedules have `.meta.json` metadata but no ground-truth products.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.parser.sheet_detector import HEADER_SYNONYMS, is_schedule_sheet
from app.parser.workbook import load_workbook_safe
from tests.synthetic_helpers import iter_string_values, load_json, post_parse, validate_schema

pytestmark = [pytest.mark.synthetic, pytest.mark.synthetic_mutated]


def _load_meta(meta_path: Path) -> dict[str, Any]:
    meta = load_json(meta_path)
    assert isinstance(meta.get("mutations"), list), f"{meta_path.name}: missing mutations list"
    return meta


def _products_from_response(data: dict[str, Any], *, name: str) -> list[dict[str, Any]]:
    products = data.get("products")
    assert isinstance(products, list), f"{name}: expected products list"
    return products


def _doc_codes(products: list[dict[str, Any]]) -> list[str]:
    codes: list[str] = []
    for product in products:
        code = (product.get("doc_code") or "").strip()
        if code:
            codes.append(code)
    return codes


def _any_field_contains(products: list[dict[str, Any]], phrase: str) -> bool:
    phrase_norm = phrase.casefold()
    for product in products:
        for value in iter_string_values(product):
            if phrase_norm in value.casefold():
                return True
    return False


def _header_tokens_casefold() -> set[str]:
    tokens: set[str] = set()
    for synonyms in HEADER_SYNONYMS.values():
        for s in synonyms:
            if isinstance(s, str) and s.strip():
                tokens.add(s.strip().casefold())
    return tokens


def test_mutated_files_schema_valid(client: TestClient, mutated_files: list[tuple[Path, Path]]) -> None:
    for xlsx_path, meta_path in mutated_files:
        data = post_parse(client, xlsx_path)
        validate_schema(data)
        assert (data.get("schedule_name") or "").strip(), f"{xlsx_path.name}: empty schedule_name"


def test_mutated_files_non_empty_products_and_doc_codes(client: TestClient, mutated_files: list[tuple[Path, Path]]) -> None:
    for xlsx_path, meta_path in mutated_files:
        data = post_parse(client, xlsx_path)
        products = _products_from_response(data, name=xlsx_path.name)
        assert products, f"{xlsx_path.name}: expected non-empty products"
        assert _doc_codes(products), f"{xlsx_path.name}: expected at least one doc_code"


def test_repeat_header_mid_sheet_does_not_create_header_products(
    client: TestClient, mutated_files: list[tuple[Path, Path]]
) -> None:
    header_tokens = _header_tokens_casefold()
    header_like_doc_codes = {"spec code", "code", "ref", "reference", "doc code", "document code"}

    for xlsx_path, meta_path in mutated_files:
        meta = _load_meta(meta_path)
        if "repeat_header_mid_sheet" not in meta["mutations"]:
            continue

        data = post_parse(client, xlsx_path)
        products = _products_from_response(data, name=xlsx_path.name)
        doc_codes = _doc_codes(products)

        assert len(doc_codes) == len(set(doc_codes)), f"{xlsx_path.name}: duplicate doc_codes found"

        for product in products:
            doc_code = (product.get("doc_code") or "").strip().casefold()
            if doc_code and (doc_code in header_tokens or doc_code in header_like_doc_codes):
                raise AssertionError(f"{xlsx_path.name}: header row emitted as product doc_code={doc_code!r}")


def test_insert_noise_rows_are_skipped(client: TestClient, mutated_files: list[tuple[Path, Path]]) -> None:
    phrases = [
        "IMAGES AND COSTS ARE INDICATIVE ONLY.",
        "REFER TO DRAWINGS AND SPECIFICATION.",
        "VERIFY ON SITE.",
        "PROJECT: Mutated Schedule",
    ]

    for xlsx_path, meta_path in mutated_files:
        meta = _load_meta(meta_path)
        if "insert_noise_rows" not in meta["mutations"]:
            continue

        data = post_parse(client, xlsx_path)
        products = _products_from_response(data, name=xlsx_path.name)
        for phrase in phrases:
            assert not _any_field_contains(products, phrase), f"{xlsx_path.name}: noise text leaked into products ({phrase})"


def test_add_extra_sheet_is_not_treated_as_schedule(mutated_files: list[tuple[Path, Path]]) -> None:
    needle = "This sheet is not part of the schedule."

    for xlsx_path, meta_path in mutated_files:
        meta = _load_meta(meta_path)
        if "add_extra_sheet" not in meta["mutations"]:
            continue

        wb = load_workbook_safe(xlsx_path.read_bytes())
        extra_sheets = []
        for ws in wb.worksheets:
            a1 = ws.cell(row=1, column=1).value
            if isinstance(a1, str) and needle.casefold() in a1.casefold():
                extra_sheets.append(ws)
        assert extra_sheets, f"{xlsx_path.name}: expected an extra sheet containing marker text"
        for ws in extra_sheets:
            assert not is_schedule_sheet(ws), f"{xlsx_path.name}: extra sheet {ws.title!r} detected as schedule sheet"


def test_swap_two_columns_column_mapping_still_extracts_key_fields(
    client: TestClient, mutated_files: list[tuple[Path, Path]]
) -> None:
    for xlsx_path, meta_path in mutated_files:
        meta = _load_meta(meta_path)
        if "swap_two_columns" not in meta["mutations"]:
            continue

        data = post_parse(client, xlsx_path)
        products = _products_from_response(data, name=xlsx_path.name)
        assert len(products) >= 10, f"{xlsx_path.name}: unexpectedly few products after column swap"

        with_doc_code = sum(1 for p in products if (p.get("doc_code") or "").strip())
        with_name = sum(1 for p in products if (p.get("product_name") or "").strip())
        assert with_doc_code / len(products) >= 0.30, f"{xlsx_path.name}: low doc_code fill-rate after column swap"
        assert with_name / len(products) >= 0.30, f"{xlsx_path.name}: low product_name fill-rate after column swap"


def test_rename_headers_fuzzy_matching_handles_synonyms(
    client: TestClient, mutated_files: list[tuple[Path, Path]]
) -> None:
    candidates = []
    for xlsx_path, meta_path in mutated_files:
        meta = _load_meta(meta_path)
        if "rename_headers" in meta["mutations"]:
            candidates.append((xlsx_path, meta_path))

    if not candidates:
        pytest.skip("No mutated files with rename_headers mutation found")

    for xlsx_path, meta_path in candidates:
        data = post_parse(client, xlsx_path)
        products = _products_from_response(data, name=xlsx_path.name)
        assert products, f"{xlsx_path.name}: expected non-empty products after header renames"
        assert _doc_codes(products), f"{xlsx_path.name}: expected doc_codes after header renames"

