"""Phase 2 Verification Tests.

This module provides comprehensive verification tests for Phase 2:
Workbook Loading + Sheet Detection.

Verifies the following requirements from tasks.md 2.5:
- Sample1: `APARTMENTS` sheet detected, header at row 4
- Sample2: `Cover Sheet` skipped, `Schedule` + `Sales Schedule` detected, header at row 9
- Sample3: `Schedule` sheet detected, header at row 10
- Schedule name extracted correctly (not formula string)
"""

import json
import pytest
from pathlib import Path

from app.parser.workbook import (
    load_workbook_safe,
    get_schedule_name,
    WorkbookLoadError,
)
from app.parser.sheet_detector import (
    find_header_row,
    is_schedule_sheet,
    get_schedule_sheets,
    get_header_columns,
)
from app.parser.merged_cells import (
    fill_merged_regions,
)


class TestPhase2VerificationSample1:
    """Verification tests for Sample1 (schedule_sample1.xlsx).

    Requirements:
    - APARTMENTS sheet detected
    - Header at row 4
    """

    @pytest.fixture
    def workbook(self):
        """Load sample1 workbook."""
        with open('data/schedule_sample1.xlsx', 'rb') as f:
            return load_workbook_safe(f.read())

    def test_apartments_sheet_exists(self, workbook):
        """Verify APARTMENTS sheet exists in the workbook."""
        assert 'APARTMENTS' in workbook.sheetnames

    def test_apartments_sheet_is_only_sheet(self, workbook):
        """Verify APARTMENTS is the only sheet."""
        assert len(workbook.sheetnames) == 1
        assert workbook.sheetnames[0] == 'APARTMENTS'

    def test_apartments_sheet_detected_as_schedule(self, workbook):
        """Verify APARTMENTS sheet is detected as a schedule sheet."""
        ws = workbook['APARTMENTS']
        assert is_schedule_sheet(ws) is True

    def test_header_at_row_4(self, workbook):
        """Verify header is found at row 4."""
        ws = workbook['APARTMENTS']
        header_row = find_header_row(ws)
        assert header_row == 4, f"Expected header at row 4, got {header_row}"

    def test_header_columns_detected(self, workbook):
        """Verify all expected columns are detected."""
        ws = workbook['APARTMENTS']
        columns = get_header_columns(ws, 4)

        # Verify key columns exist
        assert 'doc_code' in columns, "doc_code column not found"
        assert 'image' in columns, "image column not found"
        assert 'item_location' in columns, "item_location column not found"
        assert 'specs' in columns, "specs column not found"
        assert 'manufacturer' in columns, "manufacturer column not found"
        assert 'notes' in columns, "notes column not found"
        assert 'cost' in columns, "cost column not found"

    def test_get_schedule_sheets_returns_apartments(self, workbook):
        """Verify get_schedule_sheets returns APARTMENTS with correct header."""
        sheets = get_schedule_sheets(workbook)

        assert len(sheets) == 1
        sheet_name, ws, header_row = sheets[0]

        assert sheet_name == 'APARTMENTS'
        assert header_row == 4

    def test_schedule_name_extracted(self, workbook):
        """Verify schedule name is extracted correctly."""
        name = get_schedule_name(workbook, 'schedule_sample1.xlsx')

        # Should extract "12006: GEM, WATERLINE PLACE, WILLIAMSTOWN"
        assert '12006' in name
        assert 'GEM' in name
        assert 'WATERLINE PLACE' in name
        assert 'WILLIAMSTOWN' in name

        # Should NOT be a formula or filename
        assert not name.startswith('=')
        assert 'schedule_sample1' not in name.lower()

    def test_merged_cells_filled(self, workbook):
        """Verify merged cells can be filled without errors."""
        ws = workbook['APARTMENTS']

        # Count merged cells before
        merged_before = len(ws.merged_cells.ranges)
        assert merged_before > 0, "Sample1 should have merged cells"

        # Fill merged regions
        fill_merged_regions(ws)

        # Verify FLOORING section header (row 6) is filled
        assert ws['A6'].value == 'FLOORING'
        assert ws['B6'].value == 'FLOORING'


class TestPhase2VerificationSample2:
    """Verification tests for Sample2 (schedule_sample2.xlsx).

    Requirements:
    - Cover Sheet skipped
    - Schedule sheet detected, header at row 9
    - Sales Schedule sheet detected, header at row 9
    """

    @pytest.fixture
    def workbook(self):
        """Load sample2 workbook."""
        with open('data/schedule_sample2.xlsx', 'rb') as f:
            return load_workbook_safe(f.read())

    def test_all_sheets_exist(self, workbook):
        """Verify all expected sheets exist."""
        sheets = [s.strip() for s in workbook.sheetnames]

        assert 'Cover Sheet' in sheets
        assert 'Schedule' in sheets
        assert 'Sales Schedule' in sheets

    def test_cover_sheet_skipped(self, workbook):
        """Verify Cover Sheet is NOT detected as a schedule sheet."""
        ws = workbook['Cover Sheet']
        assert is_schedule_sheet(ws) is False

    def test_schedule_sheet_detected(self, workbook):
        """Verify Schedule sheet is detected as a schedule sheet."""
        ws = workbook['Schedule']
        assert is_schedule_sheet(ws) is True

    def test_schedule_header_at_row_9(self, workbook):
        """Verify Schedule sheet header is at row 9."""
        ws = workbook['Schedule']
        header_row = find_header_row(ws)
        assert header_row == 9, f"Expected header at row 9, got {header_row}"

    def test_sales_schedule_detected(self, workbook):
        """Verify Sales Schedule sheet is detected (note: trailing space)."""
        # Handle trailing space in sheet name
        sheet_name = 'Sales Schedule '
        ws = workbook[sheet_name]
        assert is_schedule_sheet(ws) is True

    def test_sales_schedule_header_at_row_9(self, workbook):
        """Verify Sales Schedule header is at row 9."""
        ws = workbook['Sales Schedule ']  # Note trailing space
        header_row = find_header_row(ws)
        assert header_row == 9, f"Expected header at row 9, got {header_row}"

    def test_get_schedule_sheets_skips_cover(self, workbook):
        """Verify get_schedule_sheets returns both schedules but not Cover Sheet."""
        sheets = get_schedule_sheets(workbook)

        sheet_names = [s[0] for s in sheets]

        # Should have 2 schedule sheets
        assert len(sheets) == 2, f"Expected 2 schedule sheets, got {len(sheets)}"

        # Cover Sheet should NOT be included
        assert 'Cover Sheet' not in sheet_names

        # Schedule and Sales Schedule should be included
        assert 'Schedule' in sheet_names
        assert 'Sales Schedule ' in sheet_names or 'Sales Schedule' in [s.strip() for s in sheet_names]

    def test_both_schedules_have_correct_header_row(self, workbook):
        """Verify both schedule sheets have header at row 9."""
        sheets = get_schedule_sheets(workbook)

        for sheet_name, ws, header_row in sheets:
            assert header_row == 9, f"{sheet_name} header at {header_row}, expected 9"

    def test_schedule_name_extracted_from_formula(self, workbook):
        """Verify schedule name is resolved from formula reference."""
        name = get_schedule_name(workbook, 'schedule_sample2.xlsx')

        # Should extract "SCHEDULE 003- INTERNAL FINISHES" (from Cover Sheet A6)
        # NOT the formula string "='[1]Cover Sheet'!A6"
        assert 'SCHEDULE 003' in name or 'INTERNAL FINISHES' in name

        # Should NOT be a formula string
        assert not name.startswith('=')
        assert 'Cover Sheet' not in name
        assert '!' not in name


class TestPhase2VerificationSample3:
    """Verification tests for Sample3 (schedule_sample3.xlsx).

    Requirements:
    - Schedule sheet detected
    - Header at row 10
    """

    @pytest.fixture
    def workbook(self):
        """Load sample3 workbook."""
        with open('data/schedule_sample3.xlsx', 'rb') as f:
            return load_workbook_safe(f.read())

    def test_schedule_sheet_exists(self, workbook):
        """Verify Schedule sheet exists."""
        assert 'Schedule' in workbook.sheetnames

    def test_schedule_sheet_detected(self, workbook):
        """Verify Schedule sheet is detected as a schedule sheet."""
        ws = workbook['Schedule']
        assert is_schedule_sheet(ws) is True

    def test_header_at_row_10(self, workbook):
        """Verify header is found at row 10."""
        ws = workbook['Schedule']
        header_row = find_header_row(ws)
        assert header_row == 10, f"Expected header at row 10, got {header_row}"

    def test_header_columns_detected(self, workbook):
        """Verify expected columns are detected."""
        ws = workbook['Schedule']
        columns = get_header_columns(ws, 10)

        # Verify key columns exist
        assert 'doc_code' in columns, "doc_code column not found"
        assert 'item_location' in columns, "item_location/area column not found"
        assert 'qty' in columns, "qty column not found"
        assert 'cost' in columns, "cost column not found"

    def test_get_schedule_sheets_returns_schedule(self, workbook):
        """Verify get_schedule_sheets returns Schedule sheet."""
        sheets = get_schedule_sheets(workbook)

        assert len(sheets) == 1
        sheet_name, ws, header_row = sheets[0]

        assert sheet_name == 'Schedule'
        assert header_row == 10

    def test_schedule_name_fallback_to_filename(self, workbook):
        """Verify schedule name falls back to filename for sample3."""
        name = get_schedule_name(workbook, 'schedule_sample3.xlsx')

        # Sample3 doesn't have a clear title, should fallback to filename
        # The function returns "schedule sample3" (underscores replaced with spaces)
        assert 'schedule' in name.lower()
        assert 'sample3' in name.lower() or 'sample 3' in name.lower()

    def test_many_merged_cells_handled(self, workbook):
        """Verify large number of merged cells is handled efficiently."""
        ws = workbook['Schedule']

        # Sample3 has many merged cells (1234+)
        merged_count = len(ws.merged_cells.ranges)
        assert merged_count > 1000, f"Sample3 should have many merged cells, got {merged_count}"

        # Fill merged regions should complete without error
        fill_merged_regions(ws)


class TestPhase2VerificationSyntheticGenerated:
    """Verification tests using synthetic generated files."""

    @pytest.fixture
    def synthetic_dir(self):
        """Get synthetic generated files directory."""
        path = Path('synthetic_out/generated')
        if not path.exists():
            pytest.skip("Synthetic files not generated")
        return path

    def test_generated_files_load_successfully(self, synthetic_dir):
        """Verify all generated files load without error."""
        xlsx_files = list(synthetic_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")

        errors = []
        for xlsx_path in xlsx_files:
            try:
                with open(xlsx_path, 'rb') as f:
                    wb = load_workbook_safe(f.read())
            except WorkbookLoadError as e:
                errors.append(f"{xlsx_path.name}: {e}")

        assert not errors, f"Failed to load files:\n" + "\n".join(errors)

    def test_generated_files_schedule_detection(self, synthetic_dir):
        """Verify schedule sheets detected in generated files with code columns."""
        xlsx_files = list(synthetic_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")

        files_with_code_col = 0
        files_detected = 0

        for xlsx_path in xlsx_files:
            truth_path = xlsx_path.with_suffix('.truth.json')

            # Check if file has code column from ground truth
            has_code_col = True
            if truth_path.exists():
                with open(truth_path) as f:
                    truth = json.load(f)
                has_code_col = truth.get('notes', {}).get('include_code_col', True)

            if not has_code_col:
                continue

            files_with_code_col += 1

            with open(xlsx_path, 'rb') as f:
                wb = load_workbook_safe(f.read())

            sheets = get_schedule_sheets(wb)
            if len(sheets) >= 1:
                files_detected += 1

        if files_with_code_col > 0:
            rate = files_detected / files_with_code_col
            assert rate >= 0.5, f"Detection rate too low: {rate:.1%}"

    def test_generated_files_header_detection_accuracy(self, synthetic_dir):
        """Verify header rows detected close to ground truth."""
        xlsx_files = list(synthetic_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")

        checked = 0
        accurate = 0

        for xlsx_path in xlsx_files:
            truth_path = xlsx_path.with_suffix('.truth.json')
            if not truth_path.exists():
                continue

            with open(truth_path) as f:
                truth = json.load(f)

            expected_row = truth.get('notes', {}).get('header_row')
            has_code_col = truth.get('notes', {}).get('include_code_col', True)

            if expected_row is None or not has_code_col:
                continue

            with open(xlsx_path, 'rb') as f:
                wb = load_workbook_safe(f.read())

            sheets = get_schedule_sheets(wb)
            if not sheets:
                continue

            checked += 1
            detected_rows = [s[2] for s in sheets]

            # Allow a small, fixed tolerance for header row detection
            tolerance = 5
            if any(abs(d - expected_row) <= tolerance for d in detected_rows):
                accurate += 1

        if checked > 0:
            rate = accurate / checked
            assert rate >= 0.8, f"Accuracy rate too low: {rate:.1%}"

    def test_generated_files_schedule_name_extraction(self, synthetic_dir):
        """Verify schedule name extraction on generated files."""
        xlsx_files = list(synthetic_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")

        for xlsx_path in xlsx_files:
            with open(xlsx_path, 'rb') as f:
                wb = load_workbook_safe(f.read())

            name = get_schedule_name(wb, xlsx_path.name)

            # Name should not be empty
            assert name, f"{xlsx_path.name}: Empty schedule name"

            # Name should not be a formula
            assert not name.startswith('='), f"{xlsx_path.name}: Formula returned as name"


class TestPhase2VerificationSyntheticMutated:
    """Verification tests using synthetic mutated files."""

    @pytest.fixture
    def mutated_dir(self):
        """Get synthetic mutated files directory."""
        path = Path('synthetic_out/mutated')
        if not path.exists():
            pytest.skip("Mutated files not generated")
        return path

    def test_mutated_files_load_successfully(self, mutated_dir):
        """Verify all mutated files load without error."""
        xlsx_files = list(mutated_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No mutated xlsx files found")

        errors = []
        for xlsx_path in xlsx_files:
            try:
                with open(xlsx_path, 'rb') as f:
                    wb = load_workbook_safe(f.read())
            except WorkbookLoadError as e:
                errors.append(f"{xlsx_path.name}: {e}")

        assert not errors, f"Failed to load files:\n" + "\n".join(errors)

    def test_mutated_files_no_crash_on_detection(self, mutated_dir):
        """Verify schedule detection doesn't crash on mutated files."""
        xlsx_files = list(mutated_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No mutated xlsx files found")

        for xlsx_path in xlsx_files:
            with open(xlsx_path, 'rb') as f:
                wb = load_workbook_safe(f.read())

            # These operations should not crash
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # Find header (may or may not find one)
                header = find_header_row(ws)

                # Check if schedule (may or may not be)
                is_schedule = is_schedule_sheet(ws)

    def test_mutated_files_merged_cells_handled(self, mutated_dir):
        """Verify merged cell handling on mutated files."""
        xlsx_files = list(mutated_dir.glob('*.xlsx'))
        if not xlsx_files:
            pytest.skip("No mutated xlsx files found")

        for xlsx_path in xlsx_files:
            with open(xlsx_path, 'rb') as f:
                wb = load_workbook_safe(f.read())

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # Fill merged regions should not crash
                fill_merged_regions(ws)


class TestPhase2VerificationSummary:
    """Summary tests that verify all Phase 2 requirements at once."""

    def test_all_samples_load_successfully(self):
        """Verify all sample files load successfully."""
        samples = [
            'data/schedule_sample1.xlsx',
            'data/schedule_sample2.xlsx',
            'data/schedule_sample3.xlsx',
        ]

        for sample_path in samples:
            with open(sample_path, 'rb') as f:
                wb = load_workbook_safe(f.read())
            assert wb is not None, f"Failed to load {sample_path}"

    def test_phase2_requirements_matrix(self):
        """Comprehensive test of all Phase 2.5 requirements.

        Requirements from tasks.md:
        - Sample1: APARTMENTS sheet detected, header at row 4
        - Sample2: Cover Sheet skipped, Schedule + Sales Schedule detected, header at row 9
        - Sample3: Schedule sheet detected, header at row 10
        - Schedule name extracted correctly (not formula string)
        """
        results = {}

        # Sample 1
        with open('data/schedule_sample1.xlsx', 'rb') as f:
            wb1 = load_workbook_safe(f.read())

        sheets1 = get_schedule_sheets(wb1)
        results['sample1_apartments_detected'] = (
            len(sheets1) == 1 and
            sheets1[0][0] == 'APARTMENTS' and
            sheets1[0][2] == 4
        )
        results['sample1_header_row_4'] = sheets1[0][2] == 4 if sheets1 else False

        name1 = get_schedule_name(wb1, 'schedule_sample1.xlsx')
        results['sample1_schedule_name'] = '12006' in name1 and not name1.startswith('=')

        # Sample 2
        with open('data/schedule_sample2.xlsx', 'rb') as f:
            wb2 = load_workbook_safe(f.read())

        sheets2 = get_schedule_sheets(wb2)
        sheet_names2 = [s[0] for s in sheets2]

        results['sample2_cover_sheet_skipped'] = 'Cover Sheet' not in sheet_names2
        results['sample2_schedule_detected'] = 'Schedule' in sheet_names2
        results['sample2_sales_schedule_detected'] = any(
            'Sales Schedule' in s.strip() for s in sheet_names2
        )
        results['sample2_header_row_9'] = all(s[2] == 9 for s in sheets2)

        name2 = get_schedule_name(wb2, 'schedule_sample2.xlsx')
        results['sample2_schedule_name'] = (
            ('SCHEDULE' in name2 or 'INTERNAL FINISHES' in name2) and
            not name2.startswith('=')
        )

        # Sample 3
        with open('data/schedule_sample3.xlsx', 'rb') as f:
            wb3 = load_workbook_safe(f.read())

        sheets3 = get_schedule_sheets(wb3)
        results['sample3_schedule_detected'] = (
            len(sheets3) == 1 and
            sheets3[0][0] == 'Schedule'
        )
        results['sample3_header_row_10'] = sheets3[0][2] == 10 if sheets3 else False

        name3 = get_schedule_name(wb3, 'schedule_sample3.xlsx')
        results['sample3_schedule_name'] = not name3.startswith('=')

        # Print results summary
        print("\n" + "=" * 60)
        print("PHASE 2 VERIFICATION RESULTS")
        print("=" * 60)

        for key, value in results.items():
            status = "PASS" if value else "FAIL"
            print(f"  {key}: {status}")

        print("=" * 60)

        # Assert all requirements met
        failed = [k for k, v in results.items() if not v]
        assert not failed, f"Failed requirements: {failed}"


@pytest.mark.synthetic
class TestPhase2SyntheticRobustness:
    """Robustness tests using synthetic data (marked for separate execution)."""

    def test_all_synthetic_files_processable(self):
        """Verify all synthetic files can be processed end-to-end."""
        generated = Path('synthetic_out/generated')
        mutated = Path('synthetic_out/mutated')

        if not generated.exists() and not mutated.exists():
            pytest.skip("Synthetic files not generated")

        all_files = []
        if generated.exists():
            all_files.extend(generated.glob('*.xlsx'))
        if mutated.exists():
            all_files.extend(mutated.glob('*.xlsx'))

        if not all_files:
            pytest.skip("No synthetic files found")

        errors = []
        for xlsx_path in all_files:
            try:
                with open(xlsx_path, 'rb') as f:
                    wb = load_workbook_safe(f.read())

                # Test all Phase 2 operations
                name = get_schedule_name(wb, xlsx_path.name)
                sheets = get_schedule_sheets(wb)

                for sheet_name, ws, header_row in sheets:
                    fill_merged_regions(ws)
                    columns = get_header_columns(ws, header_row)

            except Exception as e:
                errors.append(f"{xlsx_path.name}: {type(e).__name__}: {e}")

        # Allow up to 5% failures for edge cases in mutated files
        if errors:
            error_rate = len(errors) / len(all_files)
            if error_rate > 0.05:
                pytest.fail(
                    f"Too many errors ({error_rate:.1%}):\n" +
                    "\n".join(errors[:10])
                )
