from app.parser.normalizers import parse_dimensions


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
