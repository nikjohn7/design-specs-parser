"""Tests for the field parser module.

Tests the key-value block parsing functionality including:
- parse_kv_block: Parse multi-line KEY: VALUE text
- normalize_key: Key alias normalization
- Various separator handling (:, -, =)
"""

import pytest
from app.parser.field_parser import (
    parse_kv_block,
    parse_kv_with_multivalue,
    normalize_key,
    extract_non_kv_lines,
    merge_kv_dicts,
    get_value,
    format_kv_as_details,
    has_kv_content,
)


class TestNormalizeKey:
    """Tests for normalize_key function."""

    def test_basic_normalization(self):
        """Test basic key normalization to uppercase."""
        assert normalize_key('colour') == 'COLOUR'
        assert normalize_key('COLOUR') == 'COLOUR'
        assert normalize_key('Colour') == 'COLOUR'

    def test_color_to_colour_alias(self):
        """Test that COLOR maps to COLOUR."""
        assert normalize_key('color') == 'COLOUR'
        assert normalize_key('COLOR') == 'COLOUR'
        assert normalize_key('Color') == 'COLOUR'

    def test_surface_to_finish_alias(self):
        """Test that SURFACE maps to FINISH."""
        assert normalize_key('surface') == 'FINISH'
        assert normalize_key('SURFACE') == 'FINISH'
        assert normalize_key('Surface Finish') == 'FINISH'

    def test_composition_to_material_alias(self):
        """Test that COMPOSITION maps to MATERIAL."""
        assert normalize_key('composition') == 'MATERIAL'
        assert normalize_key('COMPOSITION') == 'MATERIAL'
        assert normalize_key('species') == 'MATERIAL'

    def test_coir_colour_alias(self):
        """Test that COIR COLOUR maps to COLOUR."""
        assert normalize_key('coir colour') == 'COLOUR'
        assert normalize_key('COIR COLOR') == 'COLOUR'

    def test_dimension_aliases(self):
        """Test dimension key aliases."""
        assert normalize_key('w') == 'WIDTH'
        assert normalize_key('wide') == 'WIDTH'
        assert normalize_key('l') == 'LENGTH'
        assert normalize_key('len') == 'LENGTH'
        assert normalize_key('d') == 'DEPTH'
        assert normalize_key('h') == 'HEIGHT'
        assert normalize_key('ht') == 'HEIGHT'
        assert normalize_key('thk') == 'THICKNESS'

    def test_code_aliases(self):
        """Test code/reference aliases."""
        assert normalize_key('code') == 'CODE'
        assert normalize_key('ref') == 'CODE'
        assert normalize_key('reference') == 'CODE'
        assert normalize_key('sku') == 'CODE'

    def test_lead_time_alias(self):
        """Test lead time normalization."""
        assert normalize_key('lead time') == 'LEAD_TIME'
        assert normalize_key('leadtime') == 'LEAD_TIME'

    def test_none_input(self):
        """Test None input returns empty string."""
        assert normalize_key(None) == ''

    def test_empty_string(self):
        """Test empty string input."""
        assert normalize_key('') == ''

    def test_whitespace_stripping(self):
        """Test whitespace is stripped."""
        assert normalize_key('  colour  ') == 'COLOUR'
        assert normalize_key('\tfinish\n') == 'FINISH'

    def test_unknown_key(self):
        """Test unknown keys are returned uppercase as-is."""
        assert normalize_key('custom_field') == 'CUSTOM_FIELD'
        assert normalize_key('MySpecialKey') == 'MYSPECIALKEY'


class TestParseKVBlockColonSeparator:
    """Tests for parse_kv_block with colon separator."""

    def test_simple_kv_pairs(self):
        """Test simple KEY: VALUE parsing."""
        text = """PRODUCT: ICONIC
CODE: 50/2833
COLOUR: SILVER SHADOW"""
        result = parse_kv_block(text)
        assert result['PRODUCT'] == 'ICONIC'
        assert result['CODE'] == '50/2833'
        assert result['COLOUR'] == 'SILVER SHADOW'

    def test_no_space_after_colon(self):
        """Test KEY:VALUE without space."""
        text = """PRODUCT:ICONIC
CODE:50/2833"""
        result = parse_kv_block(text)
        assert result['PRODUCT'] == 'ICONIC'
        assert result['CODE'] == '50/2833'

    def test_multiline_block(self):
        """Test parsing actual sample1 format."""
        text = """PRODUCT: ICONIC
CODE: 50/2833
COLOUR: SILVER SHADOW
COMPOSITION: 80% WOOL 20% SYNTHETIC
STYLE: TWIST
WIDTH: 3.66 METRES"""
        result = parse_kv_block(text)
        assert result['PRODUCT'] == 'ICONIC'
        assert result['CODE'] == '50/2833'
        assert result['COLOUR'] == 'SILVER SHADOW'
        assert result['MATERIAL'] == '80% WOOL 20% SYNTHETIC'  # Normalized
        assert result['STYLE'] == 'TWIST'
        assert result['WIDTH'] == '3.66 METRES'

    def test_manufacturer_block(self):
        """Test parsing manufacturer info block."""
        text = """NAME: VICTORIA CARPETS
ADDRESS: 7-29 GLADSTONE ROAD
WEB: WWW.EXAMPLE.COM
CONTACT: JOHN DOE
PHONE: (03) 1234 5678"""
        result = parse_kv_block(text)
        assert result['NAME'] == 'VICTORIA CARPETS'
        assert result['ADDRESS'] == '7-29 GLADSTONE ROAD'
        assert result['WEB'] == 'WWW.EXAMPLE.COM'
        assert result['CONTACT'] == 'JOHN DOE'
        assert result['PHONE'] == '(03) 1234 5678'


class TestParseKVBlockDashSeparator:
    """Tests for parse_kv_block with dash separator."""

    def test_dash_with_spaces(self):
        """Test KEY - VALUE format (dash with spaces)."""
        text = """NAME - ELM VIEW
CODE - 50/1403
FINISH - NEPTUNE"""
        result = parse_kv_block(text)
        assert result['NAME'] == 'ELM VIEW'
        assert result['CODE'] == '50/1403'
        assert result['FINISH'] == 'NEPTUNE'

    def test_dash_without_leading_space(self):
        """Test KEY- VALUE format (dash without leading space)."""
        text = """FINISH- MATT
DIMENSIONS- 600X600 MM"""
        result = parse_kv_block(text)
        assert result['FINISH'] == 'MATT'
        assert result['SIZE'] == '600X600 MM'  # DIMENSIONS normalized to SIZE


class TestParseKVBlockEqualsSeparator:
    """Tests for parse_kv_block with equals separator."""

    def test_equals_separator(self):
        """Test KEY = VALUE format."""
        text = """COLOUR = Cool Grey
FINISH = Polished
ITEM = Dishwasher
MATERIAL = Polyurethane
CODE = A008"""
        result = parse_kv_block(text)
        assert result['COLOUR'] == 'Cool Grey'
        assert result['FINISH'] == 'Polished'
        assert result['ITEM'] == 'Dishwasher'
        assert result['MATERIAL'] == 'Polyurethane'
        assert result['CODE'] == 'A008'

    def test_equals_no_spaces(self):
        """Test KEY=VALUE format without spaces."""
        text = """COLOR=Red
SIZE=Large"""
        result = parse_kv_block(text)
        assert result['COLOUR'] == 'Red'  # COLOR normalized to COLOUR
        assert result['SIZE'] == 'Large'


class TestParseKVBlockMixedFormats:
    """Tests for parse_kv_block with mixed separator formats."""

    def test_mixed_separators(self):
        """Test mixed colon, dash, and equals in same block."""
        text = """PRODUCT: Cabinet Door
FINISH - Brushed Brass
COLOUR = Terracotta
CODE: J002"""
        result = parse_kv_block(text)
        assert result['PRODUCT'] == 'Cabinet Door'
        assert result['FINISH'] == 'Brushed Brass'
        assert result['COLOUR'] == 'Terracotta'
        assert result['CODE'] == 'J002'

    def test_sample2_row14_format(self):
        """Test Sample2 Row 14 mixed format."""
        text = """NAME - BLINK
COLOUR - BLANCO
FINISH- MATT
DIMENSIONS- 600X600 MM"""
        result = parse_kv_block(text)
        assert result['NAME'] == 'BLINK'
        assert result['COLOUR'] == 'BLANCO'
        assert result['FINISH'] == 'MATT'
        assert result['SIZE'] == '600X600 MM'


class TestParseKVBlockEdgeCases:
    """Tests for parse_kv_block edge cases."""

    def test_none_input(self):
        """Test None input returns empty dict."""
        assert parse_kv_block(None) == {}

    def test_empty_string(self):
        """Test empty string returns empty dict."""
        assert parse_kv_block('') == {}

    def test_whitespace_only(self):
        """Test whitespace-only string."""
        assert parse_kv_block('   \n\t\n   ') == {}

    def test_non_kv_lines_skipped(self):
        """Test that non-KV lines are skipped."""
        text = """PRODUCT: ICONIC
This is a note without a key
CODE: 123
Another random line"""
        result = parse_kv_block(text)
        assert len(result) == 2
        assert result['PRODUCT'] == 'ICONIC'
        assert result['CODE'] == '123'

    def test_first_occurrence_wins(self):
        """Test that first occurrence of a key is kept."""
        text = """COLOUR: RED
COLOUR: BLUE
COLOUR: GREEN"""
        result = parse_kv_block(text)
        assert result['COLOUR'] == 'RED'

    def test_colon_in_value(self):
        """Test that colons in values are preserved."""
        text = "TIME: 10:30 AM"
        result = parse_kv_block(text)
        assert result['TIME'] == '10:30 AM'

    def test_numeric_only_value(self):
        """Test numeric values."""
        text = """WIDTH: 600
HEIGHT: 400"""
        result = parse_kv_block(text)
        assert result['WIDTH'] == '600'
        assert result['HEIGHT'] == '400'

    def test_integer_input(self):
        """Test non-string input is converted."""
        result = parse_kv_block(123)
        assert result == {}  # Numeric alone doesn't parse as KV

    def test_long_key_rejected(self):
        """Test that overly long keys are rejected."""
        text = "THIS IS A VERY LONG KEY THAT SHOULD NOT MATCH: value"
        result = parse_kv_block(text)
        assert len(result) == 0

    def test_key_starting_with_number_rejected(self):
        """Test that keys starting with numbers are rejected."""
        text = "123ABC: value"
        result = parse_kv_block(text)
        assert len(result) == 0


class TestParseKVWithMultivalue:
    """Tests for parse_kv_with_multivalue function."""

    def test_multiple_values_same_key(self):
        """Test collecting multiple values for same key."""
        text = """NOTES: First note
NOTES: Second note
NOTES: Third note"""
        result = parse_kv_with_multivalue(text)
        assert result['NOTES'] == ['First note', 'Second note', 'Third note']

    def test_mixed_single_and_multi(self):
        """Test mix of single and multiple value keys."""
        text = """PRODUCT: Widget
NOTES: Note 1
CODE: ABC123
NOTES: Note 2"""
        result = parse_kv_with_multivalue(text)
        assert result['PRODUCT'] == ['Widget']
        assert result['NOTES'] == ['Note 1', 'Note 2']
        assert result['CODE'] == ['ABC123']

    def test_empty_input(self):
        """Test empty input returns empty dict."""
        assert parse_kv_with_multivalue(None) == {}
        assert parse_kv_with_multivalue('') == {}


class TestExtractNonKVLines:
    """Tests for extract_non_kv_lines function."""

    def test_extracts_non_kv_lines(self):
        """Test extraction of non-KV lines."""
        text = """PRODUCT: Widget
This is a note line
CODE: 123
Another plain text line
Install per manufacturer specs."""
        result = extract_non_kv_lines(text)
        assert 'This is a note line' in result
        assert 'Another plain text line' in result
        assert 'Install per manufacturer specs.' in result
        assert len(result) == 3

    def test_all_kv_lines(self):
        """Test with all KV lines returns empty list."""
        text = """PRODUCT: Widget
CODE: 123
COLOUR: Red"""
        result = extract_non_kv_lines(text)
        assert result == []

    def test_empty_lines_skipped(self):
        """Test that empty lines are skipped."""
        text = """
PRODUCT: Widget

Some text

"""
        result = extract_non_kv_lines(text)
        assert len(result) == 1
        assert result[0] == 'Some text'


class TestMergeKVDicts:
    """Tests for merge_kv_dicts function."""

    def test_basic_merge(self):
        """Test basic dictionary merge."""
        d1 = {'A': '1', 'B': '2'}
        d2 = {'C': '3', 'D': '4'}
        result = merge_kv_dicts(d1, d2)
        assert result == {'A': '1', 'B': '2', 'C': '3', 'D': '4'}

    def test_first_occurrence_wins(self):
        """Test that first dict value wins on conflict."""
        d1 = {'COLOUR': 'RED'}
        d2 = {'COLOUR': 'BLUE', 'SIZE': 'LARGE'}
        result = merge_kv_dicts(d1, d2)
        assert result['COLOUR'] == 'RED'
        assert result['SIZE'] == 'LARGE'

    def test_multiple_dicts(self):
        """Test merging multiple dictionaries."""
        d1 = {'A': '1'}
        d2 = {'B': '2', 'A': 'X'}
        d3 = {'C': '3', 'B': 'Y'}
        result = merge_kv_dicts(d1, d2, d3)
        assert result == {'A': '1', 'B': '2', 'C': '3'}

    def test_empty_dicts(self):
        """Test with empty dictionaries."""
        result = merge_kv_dicts({}, {}, {})
        assert result == {}

    def test_none_dicts_handled(self):
        """Test that None dicts are handled gracefully."""
        d1 = {'A': '1'}
        result = merge_kv_dicts(d1, None, {})
        assert result == {'A': '1'}


class TestGetValue:
    """Tests for get_value function."""

    def test_first_key_found(self):
        """Test returns value for first matching key."""
        kv = {'NAME': 'Widget', 'CODE': '123'}
        result = get_value(kv, 'PRODUCT', 'NAME', 'ITEM')
        assert result == 'Widget'

    def test_multiple_keys_priority(self):
        """Test key priority order."""
        kv = {'NAME': 'Widget', 'PRODUCT': 'Super Widget'}
        result = get_value(kv, 'PRODUCT', 'NAME')
        assert result == 'Super Widget'

    def test_no_match_returns_default(self):
        """Test returns default when no key matches."""
        kv = {'CODE': '123'}
        result = get_value(kv, 'PRODUCT', 'NAME', default='Unknown')
        assert result == 'Unknown'

    def test_default_none(self):
        """Test default is None when not specified."""
        kv = {'CODE': '123'}
        result = get_value(kv, 'PRODUCT')
        assert result is None

    def test_key_normalization(self):
        """Test that keys are normalized during lookup."""
        kv = {'COLOUR': 'Red'}
        result = get_value(kv, 'color')
        assert result == 'Red'

    def test_empty_dict(self):
        """Test with empty dict returns default."""
        result = get_value({}, 'KEY', default='Default')
        assert result == 'Default'

    def test_none_dict(self):
        """Test with None dict returns default."""
        result = get_value(None, 'KEY', default='Default')
        assert result == 'Default'


class TestFormatKVAsDetails:
    """Tests for format_kv_as_details function."""

    def test_basic_formatting(self):
        """Test basic pipe-separated formatting."""
        kv = {'PRODUCT': 'Widget', 'CODE': '123', 'COLOUR': 'Red'}
        result = format_kv_as_details(kv)
        assert 'PRODUCT: Widget' in result
        assert 'CODE: 123' in result
        assert 'COLOUR: Red' in result
        assert ' | ' in result

    def test_exclude_keys(self):
        """Test exclusion of specific keys."""
        kv = {'PRODUCT': 'Widget', 'CODE': '123', 'COLOUR': 'Red'}
        result = format_kv_as_details(kv, exclude_keys={'CODE'})
        assert 'CODE' not in result
        assert 'PRODUCT: Widget' in result

    def test_exclude_key_normalization(self):
        """Test that exclude keys are normalized."""
        kv = {'COLOUR': 'Red', 'SIZE': 'Large'}
        result = format_kv_as_details(kv, exclude_keys={'color'})
        assert 'COLOUR' not in result
        assert 'SIZE: Large' in result

    def test_empty_dict(self):
        """Test empty dict returns None."""
        assert format_kv_as_details({}) is None

    def test_all_keys_excluded(self):
        """Test all keys excluded returns None."""
        kv = {'PRODUCT': 'Widget'}
        result = format_kv_as_details(kv, exclude_keys={'PRODUCT'})
        assert result is None


class TestHasKVContent:
    """Tests for has_kv_content function."""

    def test_with_kv_content(self):
        """Test detects KV content."""
        assert has_kv_content('PRODUCT: Widget')
        assert has_kv_content('NAME - VALUE')
        assert has_kv_content('KEY = VALUE')

    def test_without_kv_content(self):
        """Test no false positives for plain text."""
        assert not has_kv_content('Just plain text')
        assert not has_kv_content('No key value pairs here')

    def test_empty_input(self):
        """Test empty input returns False."""
        assert not has_kv_content(None)
        assert not has_kv_content('')


class TestRealWorldSamples:
    """Tests with real-world sample data patterns."""

    def test_sample1_specs_format(self):
        """Test parsing Sample1 specification format."""
        text = """PRODUCT: ICONIC
CODE: 50/2833
COLOUR: SILVER SHADOW
COMPOSITION: 80% WOOL 20% SYNTHETIC
STYLE: TWIST
WIDTH: 3.66 METRES
CARPET THICKNESS: 11 MM
PILE HEIGHT: 9 MM
PILE WEIGHT: 1356 GM PER SQM
FIRE RATING: AS/ISO 9239.1:2003
INSTALLATION: DUAL BOND"""
        result = parse_kv_block(text)
        assert result['PRODUCT'] == 'ICONIC'
        assert result['COLOUR'] == 'SILVER SHADOW'
        assert result['MATERIAL'] == '80% WOOL 20% SYNTHETIC'  # Normalized
        assert result['WIDTH'] == '3.66 METRES'
        assert result['CARPET_THICKNESS'] == '11 MM'
        assert result['INSTALLATION'] == 'DUAL BOND'

    def test_sample2_specs_format(self):
        """Test parsing Sample2 specification format (dash separator)."""
        text = """NAME - ELM VIEW
CODE - 50/1403
FINISH - NEPTUNE
COMPOSITION - 80% WOOL 20% SYNTHETIC
PILE HEIGHT - 10 MM
TOTAL WEIGHT - 1200 GM PER SQM"""
        result = parse_kv_block(text)
        assert result['NAME'] == 'ELM VIEW'
        assert result['CODE'] == '50/1403'
        assert result['FINISH'] == 'NEPTUNE'
        assert result['MATERIAL'] == '80% WOOL 20% SYNTHETIC'  # Normalized

    def test_sample2_row14_mixed_format(self):
        """Test Sample2 Row 14 with mixed dash formats."""
        text = """NAME - BLINK
COLOUR - BLANCO
FINISH- MATT
DIMENSIONS- 600X600 MM"""
        result = parse_kv_block(text)
        assert result['NAME'] == 'BLINK'
        assert result['COLOUR'] == 'BLANCO'
        assert result['FINISH'] == 'MATT'
        assert result['SIZE'] == '600X600 MM'

    def test_synthetic_equals_format(self):
        """Test synthetic data with equals separator."""
        text = """COLOUR = Cool Grey
FINISH = Polished
ITEM = Dishwasher
MATERIAL = Polyurethane
CODE = A008
Install per manufacturer specification."""
        result = parse_kv_block(text)
        assert result['COLOUR'] == 'Cool Grey'
        assert result['FINISH'] == 'Polished'
        assert result['ITEM'] == 'Dishwasher'
        assert result['MATERIAL'] == 'Polyurethane'
        assert result['CODE'] == 'A008'
        # Non-KV line should be skipped
        assert 'Install per manufacturer specification' not in str(result.values())
