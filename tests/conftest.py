from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.main import create_app


@pytest.fixture()
def client() -> Iterator[TestClient]:
    app = create_app()
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture()
def repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


@pytest.fixture()
def sample1_path(repo_root: Path) -> Path:
    return repo_root / "data" / "schedule_sample1.xlsx"


@pytest.fixture()
def sample2_path(repo_root: Path) -> Path:
    return repo_root / "data" / "schedule_sample2.xlsx"


@pytest.fixture()
def sample3_path(repo_root: Path) -> Path:
    return repo_root / "data" / "schedule_sample3.xlsx"


@pytest.fixture()
def synthetic_generated_dir(repo_root: Path) -> Path:
    path = repo_root / "synthetic_out" / "generated"
    if not path.exists():
        pytest.skip(f"Missing synthetic generated directory: {path}")
    return path


@pytest.fixture()
def synthetic_mutated_dir(repo_root: Path) -> Path:
    path = repo_root / "synthetic_out" / "mutated"
    if not path.exists():
        pytest.skip(f"Missing synthetic mutated directory: {path}")
    return path


@pytest.fixture()
def generated_files(synthetic_generated_dir: Path) -> list[tuple[Path, Path]]:
    pairs: list[tuple[Path, Path]] = []
    for xlsx_path in sorted(synthetic_generated_dir.glob("*.xlsx")):
        truth_path = xlsx_path.with_suffix(".truth.json")
        if truth_path.exists():
            pairs.append((xlsx_path, truth_path))
    if not pairs:
        pytest.skip(f"No generated synthetic .xlsx/.truth.json pairs found in {synthetic_generated_dir}")
    return pairs


@pytest.fixture()
def mutated_files(synthetic_mutated_dir: Path) -> list[tuple[Path, Path]]:
    pairs: list[tuple[Path, Path]] = []
    for xlsx_path in sorted(synthetic_mutated_dir.glob("*.xlsx")):
        meta_path = xlsx_path.with_suffix(".meta.json")
        if meta_path.exists():
            pairs.append((xlsx_path, meta_path))
    if not pairs:
        pytest.skip(f"No mutated synthetic .xlsx/.meta.json pairs found in {synthetic_mutated_dir}")
    return pairs


@pytest.fixture()
def all_synthetic_files(
    generated_files: list[tuple[Path, Path]],
    mutated_files: list[tuple[Path, Path]],
) -> list[Path]:
    return [xlsx for xlsx, _ in generated_files] + [xlsx for xlsx, _ in mutated_files]

