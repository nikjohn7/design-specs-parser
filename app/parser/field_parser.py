"""Field parsing utilities for Excel schedule parsing.

This module provides functionality to parse key-value blocks from specification
text and normalize field keys to canonical names.

Key functions:
- parse_kv_block: Parse multi-line KEY: VALUE text into a dictionary
- normalize_key: Normalize and alias field keys (COLOR → COLOUR)
"""

import re
from typing import Any


# Key aliases mapping variations to canonical names
# All keys should be uppercase
KEY_ALIASES: dict[str, str] = {
    # Product name aliases
    'PRODUCT': 'PRODUCT',
    'NAME': 'NAME',
    'ITEM': 'ITEM',
    'RANGE': 'RANGE',

    # Color aliases
    'COLOR': 'COLOUR',
    'COLOUR': 'COLOUR',
    'COIR COLOUR': 'COLOUR',
    'COIR COLOR': 'COLOUR',

    # Finish aliases
    'FINISH': 'FINISH',
    'SURFACE': 'FINISH',
    'SURFACE FINISH': 'FINISH',

    # Material aliases
    'MATERIAL': 'MATERIAL',
    'COMPOSITION': 'MATERIAL',
    'SPECIES': 'MATERIAL',

    # Dimension aliases
    'WIDTH': 'WIDTH',
    'W': 'WIDTH',
    'WIDE': 'WIDTH',
    'LENGTH': 'LENGTH',
    'L': 'LENGTH',
    'LEN': 'LENGTH',
    'LONG': 'LENGTH',
    'DEPTH': 'DEPTH',
    'D': 'DEPTH',
    'HEIGHT': 'HEIGHT',
    'H': 'HEIGHT',
    'HT': 'HEIGHT',
    'THICKNESS': 'THICKNESS',
    'THK': 'THICKNESS',

    # Size/dimensions aliases
    'SIZE': 'SIZE',
    'DIMENSIONS': 'SIZE',
    'DIMS': 'SIZE',
    'DIM': 'SIZE',
    'SHEET SIZE': 'SIZE',
    'SHEET SIZE MAX': 'SIZE',

    # Brand/manufacturer aliases
    'MAKER': 'MAKER',
    'BRAND': 'BRAND',
    'MANUFACTURER': 'MANUFACTURER',
    'SUPPLIER': 'SUPPLIER',

    # Code aliases
    'CODE': 'CODE',
    'REF': 'CODE',
    'REFERENCE': 'CODE',
    'PRODUCT CODE': 'CODE',
    'ITEM CODE': 'CODE',
    'SKU': 'CODE',

    # Other common fields
    'STYLE': 'STYLE',
    'LEAD TIME': 'LEAD_TIME',
    'LEADTIME': 'LEAD_TIME',
    'NOTES': 'NOTES',
    'NOTE': 'NOTES',
    'COMMENTS': 'NOTES',
    'COMMENT': 'NOTES',

    # Carpet/flooring specific
    'CARPET THICKNESS': 'CARPET_THICKNESS',
    'PILE HEIGHT': 'PILE_HEIGHT',
    'PILE WEIGHT': 'PILE_WEIGHT',
    'INSTALLATION': 'INSTALLATION',

    # Quantity
    'QTY': 'QTY',
    'QUANTITY': 'QTY',
}


# Patterns for detecting key-value separators
# Order matters: try more specific patterns first
KV_PATTERNS = [
    # Pattern 1: KEY: VALUE (colon separator)
    # Matches "COLOUR: SILVER SHADOW" or "COLOUR:SILVER SHADOW"
    re.compile(r'^([A-Z][A-Z0-9\s/&\-]*?)\s*:\s*(.+)$', re.IGNORECASE),

    # Pattern 2: KEY - VALUE (dash separator with spaces)
    # Matches "FINISH - MATT" or "NAME - VICTORIA CARPETS"
    re.compile(r'^([A-Z][A-Z0-9\s/&]*?)\s+-\s+(.+)$', re.IGNORECASE),

    # Pattern 3: KEY- VALUE (dash separator, no leading space)
    # Matches "FINISH- MATT" (note no space before dash)
    re.compile(r'^([A-Z][A-Z0-9\s/&]*?)-\s+(.+)$', re.IGNORECASE),

    # Pattern 4: KEY = VALUE (equals separator)
    # Matches "COLOR = Charcoal"
    re.compile(r'^([A-Z][A-Z0-9\s/&\-]*?)\s*=\s*(.+)$', re.IGNORECASE),
]


def normalize_key(key: str | None) -> str:
    """Normalize a field key to its canonical form.

    Performs the following normalizations:
    - Strip leading/trailing whitespace
    - Convert to uppercase
    - Map aliases to canonical names (e.g., COLOR → COLOUR)

    Args:
        key: Raw key string to normalize

    Returns:
        Normalized and canonicalized key string, or empty string if input is None/empty

    Examples:
        >>> normalize_key('color')
        'COLOUR'
        >>> normalize_key('Surface')
        'FINISH'
        >>> normalize_key('COMPOSITION')
        'MATERIAL'
    """
    if not key:
        return ''

    # Strip and uppercase
    normalized = key.strip().upper()

    # Look up alias
    return KEY_ALIASES.get(normalized, normalized)


def _parse_line(line: str) -> tuple[str | None, str | None]:
    """Parse a single line for key-value content.

    Attempts to match the line against known KV patterns.
    Returns None, None if no pattern matches.

    Args:
        line: A single line of text to parse

    Returns:
        Tuple of (normalized_key, value) or (None, None) if no match
    """
    line = line.strip()

    if not line:
        return None, None

    # Try each pattern in order
    for pattern in KV_PATTERNS:
        match = pattern.match(line)
        if match:
            raw_key = match.group(1).strip()
            value = match.group(2).strip()

            # Skip if key is too long (likely not a real key)
            if len(raw_key) > 30:
                continue

            # Skip if key contains numbers at start (likely a code or measurement)
            if raw_key and raw_key[0].isdigit():
                continue

            # Normalize the key
            normalized_key = normalize_key(raw_key)

            return normalized_key, value

    return None, None


def parse_kv_block(text: str | None) -> dict[str, str]:
    """Parse a multi-line text block containing KEY: VALUE pairs.

    This function handles various separator styles commonly found in
    interior design specification documents:
    - KEY: VALUE (colon separator)
    - KEY - VALUE (dash separator)
    - KEY- VALUE (dash with no leading space)
    - KEY = VALUE (equals separator)

    Keys are normalized and aliased to canonical forms (e.g., COLOR → COLOUR).
    Lines that don't match any pattern are skipped.

    Args:
        text: Multi-line text block containing key-value pairs,
              or None/empty string

    Returns:
        Dictionary mapping normalized keys to their values.
        Keys are uppercase canonical names.
        Returns empty dict if input is None or empty.

    Examples:
        >>> text = '''PRODUCT: ICONIC
        ... CODE: 50/2833
        ... COLOUR: SILVER SHADOW'''
        >>> parse_kv_block(text)
        {'PRODUCT': 'ICONIC', 'CODE': '50/2833', 'COLOUR': 'SILVER SHADOW'}

        >>> text = '''NAME - ELM VIEW
        ... FINISH - NEPTUNE
        ... COMPOSITION - 80% WOOL'''
        >>> parse_kv_block(text)
        {'NAME': 'ELM VIEW', 'FINISH': 'NEPTUNE', 'MATERIAL': '80% WOOL'}
    """
    if not text:
        return {}

    if not isinstance(text, str):
        text = str(text)

    result: dict[str, str] = {}

    # Split on newlines
    lines = text.split('\n')

    for line in lines:
        key, value = _parse_line(line)

        if key and value:
            # Keep first occurrence of each key (don't overwrite)
            if key not in result:
                result[key] = value

    return result


def parse_kv_with_multivalue(text: str | None) -> dict[str, list[str]]:
    """Parse a multi-line text block, collecting multiple values per key.

    Similar to parse_kv_block but returns a list of values for each key,
    allowing multiple values to be captured when the same key appears
    multiple times.

    Args:
        text: Multi-line text block containing key-value pairs

    Returns:
        Dictionary mapping normalized keys to lists of values.

    Example:
        >>> text = '''NOTES: First note
        ... NOTES: Second note'''
        >>> parse_kv_with_multivalue(text)
        {'NOTES': ['First note', 'Second note']}
    """
    if not text:
        return {}

    if not isinstance(text, str):
        text = str(text)

    result: dict[str, list[str]] = {}

    lines = text.split('\n')

    for line in lines:
        key, value = _parse_line(line)

        if key and value:
            if key not in result:
                result[key] = []
            result[key].append(value)

    return result


def extract_non_kv_lines(text: str | None) -> list[str]:
    """Extract lines that don't match any key-value pattern.

    Useful for capturing free-form notes or descriptions that don't
    follow the KEY: VALUE format.

    Args:
        text: Multi-line text block

    Returns:
        List of non-empty lines that don't match KV patterns
    """
    if not text:
        return []

    if not isinstance(text, str):
        text = str(text)

    non_kv_lines: list[str] = []

    lines = text.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        key, _ = _parse_line(line)

        if key is None:
            non_kv_lines.append(line)

    return non_kv_lines


def merge_kv_dicts(*dicts: dict[str, str]) -> dict[str, str]:
    """Merge multiple key-value dictionaries, keeping first occurrence.

    When the same key appears in multiple dicts, the value from the
    first dict (in argument order) is kept.

    Args:
        *dicts: Variable number of dictionaries to merge

    Returns:
        Merged dictionary with first-occurrence values

    Example:
        >>> d1 = {'COLOUR': 'RED', 'FINISH': 'MATT'}
        >>> d2 = {'COLOUR': 'BLUE', 'MATERIAL': 'WOOD'}
        >>> merge_kv_dicts(d1, d2)
        {'COLOUR': 'RED', 'FINISH': 'MATT', 'MATERIAL': 'WOOD'}
    """
    result: dict[str, str] = {}

    for d in dicts:
        if not d:
            continue
        for key, value in d.items():
            if key not in result:
                result[key] = value

    return result


def get_value(kv_dict: dict[str, str], *keys: str, default: str | None = None) -> str | None:
    """Get a value from a KV dict, trying multiple keys in order.

    Useful when a field might be stored under different key names.

    Args:
        kv_dict: Dictionary to search
        *keys: Keys to try in order (will be normalized)
        default: Default value if no key found

    Returns:
        First found value, or default if none found

    Example:
        >>> kv = {'NAME': 'ICONIC', 'CODE': '123'}
        >>> get_value(kv, 'PRODUCT', 'NAME', 'ITEM')
        'ICONIC'
    """
    if not kv_dict:
        return default

    for key in keys:
        normalized = normalize_key(key)
        if normalized in kv_dict:
            return kv_dict[normalized]

    return default


def format_kv_as_details(kv_dict: dict[str, str], exclude_keys: set[str] | None = None) -> str | None:
    """Format a KV dictionary as a pipe-separated details string.

    Produces a string like "KEY1: VALUE1 | KEY2: VALUE2 | ..." suitable
    for the product_details field.

    Args:
        kv_dict: Dictionary of key-value pairs
        exclude_keys: Keys to exclude from the output

    Returns:
        Pipe-separated string, or None if dict is empty after exclusions

    Example:
        >>> kv = {'PRODUCT': 'ICONIC', 'CODE': '123', 'STYLE': 'TWIST'}
        >>> format_kv_as_details(kv, exclude_keys={'PRODUCT'})
        'CODE: 123 | STYLE: TWIST'
    """
    if not kv_dict:
        return None

    exclude_keys = exclude_keys or set()

    # Normalize exclude keys
    normalized_exclude = {normalize_key(k) for k in exclude_keys}

    parts = []
    for key, value in kv_dict.items():
        if key in normalized_exclude:
            continue
        parts.append(f"{key}: {value}")

    if not parts:
        return None

    return ' | '.join(parts)


# Convenience function for checking if text contains KV patterns
def has_kv_content(text: str | None) -> bool:
    """Check if text contains any key-value patterns.

    Args:
        text: Text to check

    Returns:
        True if at least one KV pattern is found
    """
    if not text:
        return False

    kv = parse_kv_block(text)
    return len(kv) > 0
