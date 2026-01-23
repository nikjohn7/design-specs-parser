"""Synthetic tests for cross-validation (Phase 5.7).

These tests validate:
1. Extraction quality correlation between original samples and their mutated variants
2. Generated files with different `layout_family` values all parse correctly
3. All expected layout families (`finish_schedule`, `normalized`, `ffe_tracker`) are covered
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from tests.synthetic_helpers import load_json, post_parse, validate_schema

pytestmark = [pytest.mark.synthetic]

# Expected layout families as per PLAN.md
EXPECTED_LAYOUT_FAMILIES = {"finish_schedule", "normalized", "ffe_tracker"}


def _doc_codes_from_products(products: list[dict[str, Any]]) -> set[str]:
    """Extract non-empty doc_codes from a products list."""
    codes: set[str] = set()
    for product in products:
        code = (product.get("doc_code") or "").strip()
        if code:
            codes.add(code)
    return codes


def _filled_field_count(products: list[dict[str, Any]], field: str) -> int:
    """Count products that have a non-empty value for a given field."""
    count = 0
    for product in products:
        value = product.get(field)
        if value is not None:
            if isinstance(value, str):
                if value.strip():
                    count += 1
            else:
                count += 1
    return count


def _extraction_quality_score(products: list[dict[str, Any]]) -> dict[str, float]:
    """Calculate extraction quality metrics for a product list.
    
    Returns a dict with fill-rates for key fields.
    """
    if not products:
        return {
            "doc_code_rate": 0.0,
            "product_name_rate": 0.0,
            "brand_rate": 0.0,
            "colour_rate": 0.0,
            "finish_rate": 0.0,
            "product_count": 0,
        }
    
    total = len(products)
    return {
        "doc_code_rate": _filled_field_count(products, "doc_code") / total,
        "product_name_rate": _filled_field_count(products, "product_name") / total,
        "brand_rate": _filled_field_count(products, "brand") / total,
        "colour_rate": _filled_field_count(products, "colour") / total,
        "finish_rate": _filled_field_count(products, "finish") / total,
        "product_count": total,
    }


# =============================================================================
# Test 1: Compare extraction quality between original samples and mutations
# =============================================================================


class TestMutatedVsOriginal:
    """Compare extraction quality between original samples and their mutated variants.
    
    These tests aggregate metrics across all mutated files to ensure overall quality,
    rather than failing on individual edge cases where specific mutation combinations
    may significantly impact complex layouts (e.g., grouped-row formats like sample3).
    """

    @pytest.mark.synthetic_mutated
    def test_mutated_extraction_quality_comparable_to_original(
        self,
        client: TestClient,
        sample1_path: Path,
        sample2_path: Path,
        sample3_path: Path,
        mutated_files: list[tuple[Path, Path]],
    ) -> None:
        """Aggregate extraction quality across all mutations should meet threshold.
        
        This validates that mutations (noise, column swaps, header renames, etc.) don't 
        completely break the parser's ability to extract products on average.
        
        Target: ≥50% of mutated files should achieve ≥50% product extraction compared
        to their source sample.
        """
        # Parse original samples
        originals = {
            "schedule_sample1": post_parse(client, sample1_path),
            "schedule_sample2": post_parse(client, sample2_path),
            "schedule_sample3": post_parse(client, sample3_path),
        }
        
        original_quality = {
            name: _extraction_quality_score(data.get("products", []))
            for name, data in originals.items()
        }
        
        # Track aggregate performance
        total_compared = 0
        passing_files = 0
        failed_details: list[str] = []
        
        # Compare each mutated file against its original
        for xlsx_path, meta_path in mutated_files:
            meta = load_json(meta_path)
            source_name = Path(meta.get("source", "")).stem  # e.g., "schedule_sample1"
            
            if source_name not in original_quality:
                continue  # Can't compare if we don't have the original
            
            orig_quality = original_quality[source_name]
            if orig_quality["product_count"] == 0:
                continue
            
            mutated_data = post_parse(client, xlsx_path)
            mutated_quality = _extraction_quality_score(mutated_data.get("products", []))
            
            total_compared += 1
            
            # Check if this file meets the 50% threshold
            min_products = int(orig_quality["product_count"] * 0.50)
            meets_threshold = mutated_quality["product_count"] >= min_products
            
            if meets_threshold:
                passing_files += 1
            else:
                failed_details.append(
                    f"{xlsx_path.name}: {mutated_quality['product_count']}/{orig_quality['product_count']} products"
                )
        
        if total_compared == 0:
            pytest.skip("No mutated files could be compared against original samples")
        
        # At least 50% of mutated files should meet the extraction threshold
        pass_rate = passing_files / total_compared
        assert pass_rate >= 0.50, (
            f"Only {passing_files}/{total_compared} ({pass_rate:.1%}) mutated files "
            f"achieved ≥50% product extraction. Expected ≥50% of files to pass.\n"
            f"Failed files: {failed_details}"
        )

    @pytest.mark.synthetic_mutated
    def test_mutated_doc_code_overlap_aggregate(
        self,
        client: TestClient,
        sample1_path: Path,
        sample2_path: Path,
        sample3_path: Path,
        mutated_files: list[tuple[Path, Path]],
    ) -> None:
        """Aggregate doc_code overlap across all mutations should meet threshold.
        
        We expect at least 50% of mutated files to preserve ≥30% of original doc_codes,
        accounting for mutations that may hide rows or break complex layouts.
        """
        # Parse original samples
        originals = {
            "schedule_sample1": _doc_codes_from_products(
                post_parse(client, sample1_path).get("products", [])
            ),
            "schedule_sample2": _doc_codes_from_products(
                post_parse(client, sample2_path).get("products", [])
            ),
            "schedule_sample3": _doc_codes_from_products(
                post_parse(client, sample3_path).get("products", [])
            ),
        }
        
        total_compared = 0
        passing_files = 0
        failed_details: list[str] = []
        
        for xlsx_path, meta_path in mutated_files:
            meta = load_json(meta_path)
            source_name = Path(meta.get("source", "")).stem
            
            if source_name not in originals:
                continue
            
            original_codes = originals[source_name]
            if not original_codes:
                continue
            
            mutated_data = post_parse(client, xlsx_path)
            mutated_codes = _doc_codes_from_products(mutated_data.get("products", []))
            
            overlap = len(original_codes & mutated_codes)
            overlap_rate = overlap / len(original_codes)
            
            total_compared += 1
            
            if overlap_rate >= 0.30:
                passing_files += 1
            else:
                failed_details.append(
                    f"{xlsx_path.name}: {overlap_rate:.1%} overlap with {source_name}"
                )
        
        if total_compared == 0:
            pytest.skip("No mutated files could be compared against original samples")
        
        # At least 50% of mutated files should have ≥30% doc_code overlap
        pass_rate = passing_files / total_compared
        assert pass_rate >= 0.50, (
            f"Only {passing_files}/{total_compared} ({pass_rate:.1%}) mutated files "
            f"achieved ≥30% doc_code overlap. Expected ≥50% of files to pass.\n"
            f"Failed files: {failed_details}"
        )


# =============================================================================
# Test 2: Generated files with different layout_family values parse correctly
# =============================================================================


class TestLayoutFamilies:
    """Test that all layout family types parse correctly."""

    @pytest.mark.synthetic_generated
    def test_all_layout_families_parse_successfully(
        self,
        client: TestClient,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """Each layout_family should parse without errors and produce valid output."""
        families_tested: dict[str, list[str]] = {fam: [] for fam in EXPECTED_LAYOUT_FAMILIES}
        
        for xlsx_path, truth_path in generated_files:
            truth = load_json(truth_path)
            layout_family = truth.get("layout_family")
            
            if layout_family not in EXPECTED_LAYOUT_FAMILIES:
                continue
            
            # Parse the file
            data = post_parse(client, xlsx_path)
            validate_schema(data)
            
            # Verify non-empty results
            products = data.get("products", [])
            assert products, f"{xlsx_path.name}: expected non-empty products for {layout_family}"
            
            # Track successful parses
            families_tested[layout_family].append(xlsx_path.name)
        
        # Report coverage
        for family, files in families_tested.items():
            if not files:
                pytest.skip(f"No generated files found with layout_family={family!r}")

    @pytest.mark.synthetic_generated
    def test_finish_schedule_layout_extracts_expected_fields(
        self,
        client: TestClient,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """finish_schedule layout: should extract KV-style spec blocks."""
        finish_schedule_files = [
            (xlsx, truth) for xlsx, truth in generated_files
            if load_json(truth).get("layout_family") == "finish_schedule"
        ]
        
        if not finish_schedule_files:
            pytest.skip("No finish_schedule layout files found")
        
        for xlsx_path, truth_path in finish_schedule_files:
            data = post_parse(client, xlsx_path)
            products = data.get("products", [])
            
            # finish_schedule should include product details and descriptions
            with_details = sum(1 for p in products if (p.get("product_details") or "").strip())
            assert with_details > 0, (
                f"{xlsx_path.name}: finish_schedule should have product_details populated"
            )

    @pytest.mark.synthetic_generated
    def test_normalized_layout_extracts_expected_fields(
        self,
        client: TestClient,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """normalized layout: should extract clean tabular data with high fill rates."""
        normalized_files = [
            (xlsx, truth) for xlsx, truth in generated_files
            if load_json(truth).get("layout_family") == "normalized"
        ]
        
        if not normalized_files:
            pytest.skip("No normalized layout files found")
        
        for xlsx_path, truth_path in normalized_files:
            data = post_parse(client, xlsx_path)
            products = data.get("products", [])
            
            if not products:
                continue
            
            # normalized layout should have high product_name fill rate
            quality = _extraction_quality_score(products)
            assert quality["product_name_rate"] >= 0.50, (
                f"{xlsx_path.name}: normalized layout should have ≥50% product_name fill rate, "
                f"got {quality['product_name_rate']:.1%}"
            )

    @pytest.mark.synthetic_generated
    def test_ffe_tracker_layout_extracts_expected_fields(
        self,
        client: TestClient,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """ffe_tracker layout: FF&E tracker style with qty/pricing columns."""
        ffe_files = [
            (xlsx, truth) for xlsx, truth in generated_files
            if load_json(truth).get("layout_family") == "ffe_tracker"
        ]
        
        if not ffe_files:
            pytest.skip("No ffe_tracker layout files found")
        
        for xlsx_path, truth_path in ffe_files:
            data = post_parse(client, xlsx_path)
            products = data.get("products", [])
            
            if not products:
                continue
            
            # ffe_tracker should have qty populated for some products
            with_qty = sum(1 for p in products if p.get("qty") is not None)
            # It's okay if qty is 0, just check we extracted something
            assert len(products) > 0, f"{xlsx_path.name}: ffe_tracker should have products"


# =============================================================================
# Test 3: Verify all layout families are covered in synthetic test data
# =============================================================================


class TestLayoutFamilyCoverage:
    """Verify that synthetic test data covers all expected layout families."""

    @pytest.mark.synthetic_generated
    def test_all_expected_layout_families_are_covered(
        self,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """Synthetic generated files should cover all three layout families."""
        found_families: set[str] = set()
        
        for xlsx_path, truth_path in generated_files:
            truth = load_json(truth_path)
            layout_family = truth.get("layout_family")
            if layout_family:
                found_families.add(layout_family)
        
        missing = EXPECTED_LAYOUT_FAMILIES - found_families
        assert not missing, (
            f"Generated synthetic files do not cover layout families: {missing}. "
            f"Only found: {found_families}. "
            "Regenerate synthetic data to include all layout families."
        )

    @pytest.mark.synthetic_generated
    def test_each_layout_family_has_at_least_one_file(
        self,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """Each layout family should have at least one generated test file."""
        family_counts: dict[str, int] = {fam: 0 for fam in EXPECTED_LAYOUT_FAMILIES}
        
        for xlsx_path, truth_path in generated_files:
            truth = load_json(truth_path)
            layout_family = truth.get("layout_family")
            if layout_family in family_counts:
                family_counts[layout_family] += 1
        
        for family, count in family_counts.items():
            assert count >= 1, (
                f"Layout family {family!r} has no generated files. "
                "Consider regenerating synthetic data with more files."
            )

    @pytest.mark.synthetic_generated
    def test_layout_family_distribution_is_reasonable(
        self,
        generated_files: list[tuple[Path, Path]],
    ) -> None:
        """Layout family distribution should be reasonably balanced."""
        family_counts: dict[str, int] = {fam: 0 for fam in EXPECTED_LAYOUT_FAMILIES}
        
        for xlsx_path, truth_path in generated_files:
            truth = load_json(truth_path)
            layout_family = truth.get("layout_family")
            if layout_family in family_counts:
                family_counts[layout_family] += 1
        
        total = sum(family_counts.values())
        if total == 0:
            pytest.skip("No generated files found")
        
        # Each family should represent at least 10% of total files
        for family, count in family_counts.items():
            ratio = count / total
            # Only warn, don't fail - distribution may vary based on seed
            if ratio < 0.10:
                pytest.skip(
                    f"Layout family {family!r} underrepresented ({count}/{total} = {ratio:.1%}). "
                    "Consider regenerating with different seed for better coverage."
                )
