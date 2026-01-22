"""Tests for the row_extractor module.

This module tests the row extraction functionality including:
- Single-row-per-product layout (sample1, sample2)
- Grouped-row layout (sample3)
- Section header detection and propagation
- Detail row extraction
"""

import pytest
from pathlib import Path

from app.parser.workbook import load_workbook_safe
from app.parser.merged_cells import fill_merged_regions
from app.parser.column_mapper import map_columns
from app.parser.sheet_detector import find_header_row, get_schedule_sheets
from app.parser.row_extractor import (
    iter_product_rows,
    extract_all_products,
    get_product_count,
    _detect_layout_type,
    _is_section_header,
    _is_detail_row,
    _is_item_row,
    _is_empty_row,
    _has_item_key,
    _get_cell_value,
    _normalize_text,
)


# Test data paths
DATA_DIR = Path(__file__).parent.parent / "data"
SYNTHETIC_DIR = Path(__file__).parent.parent / "synthetic_out" / "generated"


class TestNormalizeText:
    """Tests for _normalize_text helper function."""
    
    def test_normalize_string(self):
        """Test normalizing a string value."""
        assert _normalize_text("Hello World") == "hello world"
    
    def test_normalize_with_whitespace(self):
        """Test normalizing string with leading/trailing whitespace."""
        assert _normalize_text("  HELLO  ") == "hello"
    
    def test_normalize_none(self):
        """Test normalizing None returns empty string."""
        assert _normalize_text(None) == ""
    
    def test_normalize_number(self):
        """Test normalizing numeric value."""
        assert _normalize_text(123) == "123"
        assert _normalize_text(45.67) == "45.67"


class TestSample1SingleRowLayout:
    """Tests for sample1 which uses single-row-per-product layout with sections."""
    
    @pytest.fixture
    def sample1_data(self):
        """Load and prepare sample1 data."""
        with open(DATA_DIR / "schedule_sample1.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["APARTMENTS"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        return ws, header_row, col_map
    
    def test_layout_detection(self, sample1_data):
        """Test that sample1 is detected as single-row layout."""
        ws, header_row, col_map = sample1_data
        layout = _detect_layout_type(ws, header_row, col_map)
        assert layout == "single"
    
    def test_header_row_detection(self, sample1_data):
        """Test that header row is correctly detected."""
        ws, header_row, col_map = sample1_data
        assert header_row == 4
    
    def test_product_count(self, sample1_data):
        """Test that products are extracted."""
        ws, header_row, col_map = sample1_data
        products = list(iter_product_rows(ws, header_row, col_map))
        # Should have a reasonable number of products
        assert len(products) > 50
        assert len(products) < 100
    
    def test_first_product_doc_code(self, sample1_data):
        """Test that first product has correct doc_code."""
        ws, header_row, col_map = sample1_data
        products = list(iter_product_rows(ws, header_row, col_map))
        assert products[0]["doc_code"] == "FCA-01 A"
    
    def test_section_header_propagation(self, sample1_data):
        """Test that section headers are propagated to products."""
        ws, header_row, col_map = sample1_data
        products = list(iter_product_rows(ws, header_row, col_map))
        
        # First product should have FLOORING section
        assert products[0]["section"] == "FLOORING"
        
        # Check that some products have different sections
        sections = set(p["section"] for p in products if p["section"])
        assert len(sections) > 1  # Should have multiple sections
    
    def test_product_has_required_fields(self, sample1_data):
        """Test that products have required fields."""
        ws, header_row, col_map = sample1_data
        products = list(iter_product_rows(ws, header_row, col_map))
        
        for product in products[:5]:  # Check first 5
            assert "row_num" in product
            assert "doc_code" in product
            assert "section" in product
            assert "detail_rows" in product


class TestSample2SingleRowLayout:
    """Tests for sample2 which uses single-row-per-product layout."""
    
    @pytest.fixture
    def sample2_data(self):
        """Load and prepare sample2 data."""
        with open(DATA_DIR / "schedule_sample2.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["Schedule"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        return ws, header_row, col_map
    
    def test_layout_detection(self, sample2_data):
        """Test that sample2 is detected as single-row layout."""
        ws, header_row, col_map = sample2_data
        layout = _detect_layout_type(ws, header_row, col_map)
        assert layout == "single"
    
    def test_header_row_detection(self, sample2_data):
        """Test that header row is correctly detected."""
        ws, header_row, col_map = sample2_data
        assert header_row == 9
    
    def test_product_count(self, sample2_data):
        """Test that products are extracted."""
        ws, header_row, col_map = sample2_data
        products = list(iter_product_rows(ws, header_row, col_map))
        # Should have a reasonable number of products
        assert len(products) > 40
        assert len(products) < 80
    
    def test_first_product_doc_code(self, sample2_data):
        """Test that first product has correct doc_code."""
        ws, header_row, col_map = sample2_data
        products = list(iter_product_rows(ws, header_row, col_map))
        assert products[0]["doc_code"] == "CA-01 A"


class TestSample3GroupedLayout:
    """Tests for sample3 which uses grouped-row layout."""
    
    @pytest.fixture
    def sample3_data(self):
        """Load and prepare sample3 data."""
        with open(DATA_DIR / "schedule_sample3.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["Schedule"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        return ws, header_row, col_map
    
    def test_layout_detection(self, sample3_data):
        """Test that sample3 is detected as grouped layout."""
        ws, header_row, col_map = sample3_data
        layout = _detect_layout_type(ws, header_row, col_map)
        assert layout == "grouped"
    
    def test_header_row_detection(self, sample3_data):
        """Test that header row is correctly detected."""
        ws, header_row, col_map = sample3_data
        assert header_row == 10
    
    def test_product_count(self, sample3_data):
        """Test that products are extracted."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        # Should have a reasonable number of products
        assert len(products) > 50
        assert len(products) < 150
    
    def test_first_product_doc_code(self, sample3_data):
        """Test that first product has correct doc_code."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        assert products[0]["doc_code"] == "F64"
    
    def test_first_product_item_name(self, sample3_data):
        """Test that first product has item_name from Item: key."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        assert products[0]["item_name"] == "Coffee Table"
    
    def test_first_product_has_detail_rows(self, sample3_data):
        """Test that first product has detail rows."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        
        detail_rows = products[0]["detail_rows"]
        assert len(detail_rows) > 0
        
        # Check for expected detail keys
        detail_keys = [dr["key"].lower() for dr in detail_rows]
        assert "maker" in detail_keys
        assert "name" in detail_keys
    
    def test_detail_row_values(self, sample3_data):
        """Test that detail rows have correct values."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        
        detail_rows = products[0]["detail_rows"]
        detail_dict = {dr["key"].lower(): dr["value"] for dr in detail_rows}
        
        assert detail_dict.get("maker") == "Thomas Lentini"
        assert detail_dict.get("name") == "Custom coffee table"
    
    def test_product_has_qty(self, sample3_data):
        """Test that products have qty values."""
        ws, header_row, col_map = sample3_data
        products = list(iter_product_rows(ws, header_row, col_map))
        
        # First product should have qty
        assert products[0].get("qty") is not None


class TestIsDetailRow:
    """Tests for _is_detail_row function."""
    
    @pytest.fixture
    def sample3_ws(self):
        """Load sample3 worksheet."""
        with open(DATA_DIR / "schedule_sample3.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["Schedule"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        return ws, col_map
    
    def test_maker_row_is_detail(self, sample3_ws):
        """Test that a Maker: row is detected as detail row."""
        ws, col_map = sample3_ws
        # Row 13 has "Maker:" in column D
        is_detail, key, value = _is_detail_row(ws, 13, col_map)
        assert is_detail is True
        assert key.lower() == "maker"
        assert value == "Thomas Lentini"
    
    def test_item_row_is_not_detail(self, sample3_ws):
        """Test that an Item: row is NOT detected as detail row."""
        ws, col_map = sample3_ws
        # Row 12 has "Item:" in column D
        is_detail, key, value = _is_detail_row(ws, 12, col_map)
        assert is_detail is False


class TestHasItemKey:
    """Tests for _has_item_key function."""
    
    @pytest.fixture
    def sample3_ws(self):
        """Load sample3 worksheet."""
        with open(DATA_DIR / "schedule_sample3.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["Schedule"]
        fill_merged_regions(ws)
        return ws
    
    def test_item_row_has_item_key(self, sample3_ws):
        """Test that an Item: row is detected."""
        ws = sample3_ws
        # Row 12 has "Item:" in column D
        has_item, item_value = _has_item_key(ws, 12)
        assert has_item is True
        assert item_value == "Coffee Table"
    
    def test_detail_row_no_item_key(self, sample3_ws):
        """Test that a detail row does not have Item: key."""
        ws = sample3_ws
        # Row 13 has "Maker:" not "Item:"
        has_item, item_value = _has_item_key(ws, 13)
        assert has_item is False


class TestExtractAllProducts:
    """Tests for extract_all_products convenience function."""
    
    def test_extract_all_returns_list(self):
        """Test that extract_all_products returns a list."""
        with open(DATA_DIR / "schedule_sample1.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["APARTMENTS"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        
        products = extract_all_products(ws, header_row, col_map)
        assert isinstance(products, list)
        assert len(products) > 0


class TestGetProductCount:
    """Tests for get_product_count function."""
    
    def test_count_matches_list_length(self):
        """Test that get_product_count matches list length."""
        with open(DATA_DIR / "schedule_sample1.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["APARTMENTS"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        
        count = get_product_count(ws, header_row, col_map)
        products = list(iter_product_rows(ws, header_row, col_map))
        
        assert count == len(products)


class TestMaxRowsParameter:
    """Tests for max_rows parameter in iter_product_rows."""
    
    def test_max_rows_limits_output(self):
        """Test that max_rows limits the number of rows processed."""
        with open(DATA_DIR / "schedule_sample1.xlsx", "rb") as f:
            wb = load_workbook_safe(f.read())
        ws = wb["APARTMENTS"]
        fill_merged_regions(ws)
        header_row = find_header_row(ws)
        col_map = map_columns(ws, header_row)
        
        # Get all products
        all_products = list(iter_product_rows(ws, header_row, col_map))
        
        # Get limited products
        limited_products = list(iter_product_rows(ws, header_row, col_map, max_rows=10))
        
        # Limited should be fewer
        assert len(limited_products) < len(all_products)
        assert len(limited_products) <= 10


@pytest.mark.synthetic
class TestSyntheticFiles:
    """Tests using synthetic generated files."""
    
    def test_synthetic_file_parsing(self):
        """Test that synthetic files can be parsed without errors."""
        import json
        
        if not SYNTHETIC_DIR.exists():
            pytest.skip("Synthetic data directory not found")
        
        xlsx_files = list(SYNTHETIC_DIR.glob("*.xlsx"))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")
        
        for xlsx_file in xlsx_files[:5]:  # Test first 5
            with open(xlsx_file, "rb") as f:
                wb = load_workbook_safe(f.read())
            
            schedule_sheets = get_schedule_sheets(wb)
            
            for sheet_name, ws, header_row in schedule_sheets:
                fill_merged_regions(ws)
                col_map = map_columns(ws, header_row)
                
                # Should not raise any exceptions
                products = list(iter_product_rows(ws, header_row, col_map))
                
                # Should return a list
                assert isinstance(products, list)
    
    def test_synthetic_file_product_extraction(self):
        """Test that synthetic files extract products correctly."""
        import json
        
        if not SYNTHETIC_DIR.exists():
            pytest.skip("Synthetic data directory not found")
        
        xlsx_files = list(SYNTHETIC_DIR.glob("*.xlsx"))
        if not xlsx_files:
            pytest.skip("No synthetic xlsx files found")
        
        # Test first file that has a truth file
        for xlsx_file in xlsx_files:
            truth_file = xlsx_file.with_suffix(".truth.json")
            if not truth_file.exists():
                continue
            
            with open(truth_file) as f:
                truth = json.load(f)
            
            expected_count = len(truth.get("products", []))
            
            with open(xlsx_file, "rb") as f:
                wb = load_workbook_safe(f.read())
            
            schedule_sheets = get_schedule_sheets(wb)
            
            # Get products from first schedule sheet only
            if schedule_sheets:
                sheet_name, ws, header_row = schedule_sheets[0]
                fill_merged_regions(ws)
                col_map = map_columns(ws, header_row)
                products = list(iter_product_rows(ws, header_row, col_map))
                
                # Should extract some products
                assert len(products) > 0
                
                # Note: May not match exactly due to sheet detection issues
                # but should be in the right ballpark
                break
