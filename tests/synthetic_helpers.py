from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from fastapi.testclient import TestClient

from app.core.models import ParseResponse


def post_parse(client: TestClient, xlsx_path: Path) -> dict[str, Any]:
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
    assert response.status_code == 200, f"{xlsx_path.name}: {response.text[:400]}"
    data = response.json()
    assert isinstance(data, dict), f"{xlsx_path.name}: expected JSON object"
    return data


def validate_schema(data: dict[str, Any]) -> ParseResponse:
    if hasattr(ParseResponse, "model_validate"):  # Pydantic v2
        return ParseResponse.model_validate(data)
    return ParseResponse.parse_obj(data)  # Pydantic v1


def load_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    assert isinstance(data, dict), f"{path.name}: expected JSON object"
    return data


def iter_string_values(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        out: list[str] = []
        for item in value:
            out.extend(iter_string_values(item))
        return out
    if isinstance(value, dict):
        out = []
        for item in value.values():
            out.extend(iter_string_values(item))
        return out
    return []

