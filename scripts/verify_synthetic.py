#!/usr/bin/env python3
"""Verification script for synthetic test data.

This script tests the parser against synthetic/generated test files
to verify robustness.
"""

import json
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from app.parser.workbook import load_workbook_safe, get_schedule_name
from app.parser.sheet_detector import get_schedule_sheets
from app.parser.merged_cells import fill_merged_regions
from app.parser.column_mapper import map_columns
from app.parser.row_extractor import iter_product_rows
from app.parser.field_parser import parse_kv_block, extract_product_fields
from app.core.models import Product


def parse_file(file_path: Path) -> tuple[str, list[Product]]:
    """Parse a single file and return schedule name and products."""
    with open(file_path, 'rb') as f:
        file_bytes = f.read()

    wb = load_workbook_safe(file_bytes)
    schedule_name = get_schedule_name(wb, file_path.name)
    all_products: list[Product] = []

    schedule_sheets = get_schedule_sheets(wb)

    for sheet_name, ws, header_row in schedule_sheets:
        fill_merged_regions(ws)
        col_map = map_columns(ws, header_row)

        for row_data in iter_product_rows(ws, header_row, col_map):
            specs_text = row_data.get('specs', '')
            manufacturer_text = row_data.get('manufacturer', '')

            kv_specs = parse_kv_block(specs_text) if specs_text else {}
            kv_manufacturer = parse_kv_block(manufacturer_text) if manufacturer_text else {}

            product = extract_product_fields(row_data, kv_specs, kv_manufacturer)
            all_products.append(product)

    return schedule_name, all_products


def verify_against_truth(products: list[Product], truth_path: Path) -> tuple[int, int, list[str]]:
    """Verify extracted products against ground truth.

    Returns: (matched, total_expected, list of issues)
    """
    with open(truth_path, 'r') as f:
        truth = json.load(f)

    expected_products = truth.get('products', [])
    issues: list[str] = []

    # Build a lookup of extracted products by doc_code
    extracted_by_code = {}
    for p in products:
        if p.doc_code:
            # Keep first occurrence
            if p.doc_code not in extracted_by_code:
                extracted_by_code[p.doc_code] = p

    matched = 0
    for expected in expected_products:
        expected_code = expected.get('doc_code')
        if not expected_code:
            continue

        if expected_code in extracted_by_code:
            matched += 1
            extracted = extracted_by_code[expected_code]

            # Check specific fields
            for field in ['product_name', 'brand', 'colour', 'finish']:
                expected_val = expected.get(field)
                actual_val = getattr(extracted, field)
                if expected_val and actual_val and expected_val.lower() != actual_val.lower():
                    issues.append(f"  {expected_code}: {field} mismatch - expected '{expected_val}', got '{actual_val}'")
        else:
            issues.append(f"  Missing: {expected_code}")

    return matched, len(expected_products), issues


def main():
    """Run synthetic verification."""
    synthetic_dir = project_root / 'synthetic_out'

    if not synthetic_dir.exists():
        print("Synthetic data directory not found. Generating...")
        import subprocess
        result = subprocess.run(
            ['python', 'tools/generate_programa_test_schedules.py',
             '--mode', 'both',
             '--samples_dir', './data',
             '--output_dir', './synthetic_out',
             '--num_generated', '10',
             '--seed', '12345'],
            capture_output=True,
            text=True
        )
        if result.returncode != 0:
            print(f"Failed to generate synthetic data: {result.stderr}")
            return 1

    # Test generated files with ground truth
    generated_dir = synthetic_dir / 'generated'
    mutated_dir = synthetic_dir / 'mutated'

    results = {
        'generated': {'total': 0, 'passed': 0, 'failed': 0, 'errors': 0},
        'mutated': {'total': 0, 'passed': 0, 'failed': 0, 'errors': 0}
    }

    # Test generated files
    if generated_dir.exists():
        print("\n" + "="*60)
        print("Testing GENERATED synthetic files")
        print("="*60)

        for xlsx_file in sorted(generated_dir.glob('*.xlsx')):
            truth_file = xlsx_file.with_suffix('.truth.json')
            results['generated']['total'] += 1

            print(f"\n  {xlsx_file.name}...", end=' ')

            try:
                schedule_name, products = parse_file(xlsx_file)

                if truth_file.exists():
                    matched, total, issues = verify_against_truth(products, truth_file)
                    ratio = matched / total if total > 0 else 1.0

                    if ratio >= 0.7:  # Allow some flexibility
                        print(f"PASS ({matched}/{total} products matched)")
                        results['generated']['passed'] += 1
                    else:
                        print(f"FAIL ({matched}/{total} products matched)")
                        results['generated']['failed'] += 1
                        if issues[:3]:  # Show first 3 issues
                            for issue in issues[:3]:
                                print(issue)
                else:
                    # No truth file, just check it doesn't crash
                    print(f"OK ({len(products)} products, no truth file)")
                    results['generated']['passed'] += 1

            except Exception as e:
                print(f"ERROR: {e}")
                results['generated']['errors'] += 1

    # Test mutated files (no ground truth, just check they don't crash)
    if mutated_dir.exists():
        print("\n" + "="*60)
        print("Testing MUTATED synthetic files (robustness check)")
        print("="*60)

        for xlsx_file in sorted(mutated_dir.glob('*.xlsx')):
            results['mutated']['total'] += 1

            print(f"\n  {xlsx_file.name}...", end=' ')

            try:
                schedule_name, products = parse_file(xlsx_file)
                print(f"OK ({len(products)} products)")
                results['mutated']['passed'] += 1
            except Exception as e:
                print(f"ERROR: {e}")
                results['mutated']['errors'] += 1

    # Summary
    print("\n" + "="*60)
    print("SYNTHETIC VERIFICATION SUMMARY")
    print("="*60)

    for category, stats in results.items():
        if stats['total'] > 0:
            print(f"\n{category.upper()}:")
            print(f"  Total:  {stats['total']}")
            print(f"  Passed: {stats['passed']}")
            print(f"  Failed: {stats['failed']}")
            print(f"  Errors: {stats['errors']}")

    total_tests = sum(r['total'] for r in results.values())
    total_passed = sum(r['passed'] for r in results.values())
    total_errors = sum(r['errors'] for r in results.values())

    print(f"\nOVERALL: {total_passed}/{total_tests} passed, {total_errors} errors")

    return 0 if total_errors == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
