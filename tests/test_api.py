"""API integration tests (Phase 5.2).

These tests exercise the FastAPI app end-to-end:
- GET /health
- POST /parse against the provided sample schedules
- POST /parse error handling for invalid uploads

Assertions are intentionally non-brittle: we prefer schema checks, non-empty
results, and a small number of spot-checks of known values.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient


def _post_parse(client: TestClient, xlsx_path: Path) -> Any:
    response = client.post(
        "/parse",
        files={
            "file": (
                xlsx_path.name,
                xlsx_path.read_bytes(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert response.headers.get("content-type", "").startswith("application/json")
    return response


def _find_product(products: list[dict[str, Any]], doc_code: str) -> dict[str, Any]:
    for product in products:
        if (product.get("doc_code") or "").strip() == doc_code:
            return product
    raise AssertionError(f"Product with doc_code={doc_code!r} not found")


def test_health_returns_ok(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_parse_sample1_returns_valid_json(client: TestClient, sample1_path: Path) -> None:
    response = _post_parse(client, sample1_path)
    assert response.status_code == 200

    data = response.json()
    assert isinstance(data.get("schedule_name"), str)
    assert data["schedule_name"].strip()
    assert not data["schedule_name"].lstrip().startswith("=")

    products = data.get("products")
    assert isinstance(products, list)
    assert len(products) > 0

    # Spot-check known values from tasks.md (Phase 3.7).
    iconic = _find_product(products, "FCA-01 A")
    assert iconic.get("product_name") == "ICONIC"
    assert iconic.get("colour") == "SILVER SHADOW"
    assert iconic.get("width") == 3660


def test_parse_sample2_returns_valid_json(client: TestClient, sample2_path: Path) -> None:
    response = _post_parse(client, sample2_path)
    assert response.status_code == 200

    data = response.json()
    assert isinstance(data.get("schedule_name"), str)
    assert data["schedule_name"].strip()
    assert not data["schedule_name"].lstrip().startswith("=")

    products = data.get("products")
    assert isinstance(products, list)
    assert len(products) > 0

    # Spot-check known values from tasks.md (Phase 3.7).
    blink = _find_product(products, "FTI-01 A")
    assert blink.get("product_name") == "BLINK"
    assert blink.get("colour") == "BLANCO"
    assert blink.get("width") == 600
    assert blink.get("height") == 600


def test_parse_sample3_returns_valid_json(client: TestClient, sample3_path: Path) -> None:
    response = _post_parse(client, sample3_path)
    assert response.status_code == 200

    data = response.json()
    assert isinstance(data.get("schedule_name"), str)
    assert data["schedule_name"].strip()
    assert not data["schedule_name"].lstrip().startswith("=")

    products = data.get("products")
    assert isinstance(products, list)
    assert len(products) > 0

    # Spot-check known values from tasks.md (Phase 3.7).
    plinth = _find_product(products, "F88")
    assert plinth.get("brand") == "Eaglestone"
    assert plinth.get("product_name") == "Rectangular plinth coffee table"
    assert plinth.get("width") == 1200
    assert plinth.get("length") == 800
    assert plinth.get("height") == 330


def test_parse_rejects_non_xlsx_filename(client: TestClient) -> None:
    response = client.post(
        "/parse",
        files={"file": ("not-an-xlsx.txt", b"hello", "text/plain")},
    )
    assert response.status_code == 400

    data = response.json()
    assert data.get("error") == "Invalid file format"
    assert isinstance(data.get("detail"), str)


def test_parse_rejects_invalid_xlsx_bytes(client: TestClient) -> None:
    response = client.post(
        "/parse",
        files={
            "file": (
                "bad.xlsx",
                b"this is not a zip file",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert response.status_code == 400

    data = response.json()
    assert isinstance(data.get("error"), str)
    # Error message may vary depending on which validation triggers first.
    assert data["error"] in {"Invalid file", "Invalid file format", "Invalid Excel file", "Failed to load workbook"}
