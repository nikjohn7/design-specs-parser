from app.parser.normalizers import parse_dimensions, parse_price


class TestParseDimensions:
    def test_none_input(self):
        assert parse_dimensions(None) == {"width": None, "length": None, "height": None}

    def test_explicit_width_metres(self):
        assert parse_dimensions("WIDTH: 3.66 METRES")["width"] == 3660

    def test_wxh_labeled_mm(self):
        dims = parse_dimensions("600 W X 600 H MM")
        assert dims["width"] == 600
        assert dims["height"] == 600

    def test_wxl_labeled_mm_with_noise(self):
        dims = parse_dimensions("SIZE - GRANDE BOARD - 220 W X 2200 L MM")
        assert dims["width"] == 220
        assert dims["length"] == 2200

    def test_sheet_size_unlabeled(self):
        dims = parse_dimensions("SHEET SIZE MAX: 5500 X 2800 MM")
        assert dims["width"] == 5500
        assert dims["length"] == 2800

    def test_thickness_maps_to_height(self):
        dims = parse_dimensions("THICKNESS: 10MM")
        assert dims["height"] == 10

    def test_cm_conversion(self):
        dims = parse_dimensions("WIDTH: 60 CM")
        assert dims["width"] == 600

    def test_three_part_unlabeled(self):
        dims = parse_dimensions("1200 X 800 X 330 MM")
        assert dims["width"] == 1200
        assert dims["length"] == 800
        assert dims["height"] == 330


class TestParsePrice:
    def test_none_input(self):
        assert parse_price(None) is None

    def test_empty_input(self):
        assert parse_price("") is None
        assert parse_price("   ") is None

    def test_non_numeric_tokens(self):
        assert parse_price("TBC") is None
        assert parse_price("POA") is None
        assert parse_price("N/A") is None
        assert parse_price("-") is None

    def test_dollar_amount(self):
        assert parse_price("$45.50") == 45.50
        assert parse_price("$25+GST") == 25.0
        assert parse_price("$25 +GST PER SQM") == 25.0

    def test_thousands_separator(self):
        assert parse_price("$1,234.56") == 1234.56

    def test_context_without_dollar(self):
        assert parse_price("RRP 199.99") == 199.99
        assert parse_price("Cost per unit: 25") == 25.0

    def test_avoids_unrelated_numbers(self):
        assert parse_price("SIZE: 600 X 600 MM") is None
