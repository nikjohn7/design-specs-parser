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


def _coerce_nonempty_str(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        v = value.strip()
        return v or None
    v = str(value).strip()
    return v or None


def _parse_qty(value: Any) -> int | None:
    """Parse quantity from a worksheet value.

    Accepts ints/floats (normalizes 1.0 -> 1) and strings like "2", "2.0", "2 pcs".
    """
    if value is None or value is False:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, int):
        return value if value >= 0 else None

    if isinstance(value, float):
        if value < 0:
            return None
        if value.is_integer():
            return int(value)
        # Very occasionally Excel stores quantities as floats; only accept if very close to an int
        rounded = round(value)
        if abs(value - rounded) < 1e-6:
            return int(rounded)
        return None

    text = _coerce_nonempty_str(value)
    if not text:
        return None

    match = re.search(r'(\d+(?:\.\d+)?)', text.replace(',', ''))
    if not match:
        return None

    try:
        number = float(match.group(1))
    except ValueError:
        return None

    if number < 0:
        return None
    if number.is_integer():
        return int(number)
    rounded = round(number)
    if abs(number - rounded) < 1e-6:
        return int(rounded)
    return None


def _parse_numeric_price(value: Any) -> float | None:
    """Parse a numeric unit price from a worksheet value.

    Task 3.4 scope: prefer numeric price columns when present. This helper is
    intentionally conservative; full text price parsing is implemented in 3.6.
    """
    if value is None or value is False:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        v = float(value)
        return v if v >= 0 else None

    text = _coerce_nonempty_str(value)
    if not text:
        return None

    # Strip common currency formatting and parse if the result is purely numeric-ish.
    cleaned = text.strip().replace(',', '')
    cleaned = re.sub(r'^\$', '', cleaned)
    if not re.fullmatch(r'\d+(?:\.\d+)?', cleaned):
        return None

    try:
        v = float(cleaned)
    except ValueError:
        return None
    return v if v >= 0 else None


def _build_product_description(section: str | None, item_location: str | None) -> str | None:
    section = _coerce_nonempty_str(section)
    item_location = _coerce_nonempty_str(item_location)

    if section and item_location:
        return f"{section} | {item_location}"
    if section:
        return section
    return item_location


def _normalize_detail_rows(detail_rows: Any) -> dict[str, str]:
    """Normalize grouped-layout detail rows into a KV dict.

    Expected detail_rows format from row_extractor:
      [{'key': 'maker', 'value': 'Acme'}, ...]
    """
    if not detail_rows or not isinstance(detail_rows, list):
        return {}

    kv: dict[str, str] = {}
    for item in detail_rows:
        if not isinstance(item, dict):
            continue
        key = normalize_key(_coerce_nonempty_str(item.get('key')))
        value = _coerce_nonempty_str(item.get('value'))
        if not key or not value:
            continue
        if key not in kv:
            kv[key] = value
    return kv


def extract_product_fields(
    row_data: dict[str, Any],
    kv_specs: dict[str, str] | None,
    kv_manufacturer: dict[str, str] | None,
):
    """Extract a Product model from raw row data and parsed KV blocks.

    This function maps the most common schedule fields into the Product schema:
    - PRODUCT/NAME/RANGE -> product_name (priority order, with grouped-row overrides)
    - COMPOSITION/MATERIAL/SPECIES -> material (normalized by parse_kv_block)
    - COLOUR/COLOR -> colour (normalized to COLOUR)
    - FINISH -> finish
    - qty from quantity columns when present (normalize 1.0 -> 1)
    - rrp from numeric price columns when present
    - product_description from section + item_location
    - product_details from remaining KV pairs

    For grouped rows (sample3 style), detail rows take precedence:
    - Maker: -> brand
    - Name: -> product_name
    """
    from app.core.models import Product
    from app.parser.normalizers import parse_dimensions, parse_price

    kv_specs = kv_specs or {}
    kv_manufacturer = kv_manufacturer or {}

    detail_kv = _normalize_detail_rows(row_data.get('detail_rows'))

    def _parse_mm_cell(value: Any) -> int | None:
        if value is None or value is False:
            return None
        if isinstance(value, bool):
            return None
        if isinstance(value, int):
            return value if value >= 0 else None
        if isinstance(value, float):
            if value < 0:
                return None
            if value.is_integer():
                return int(value)
            rounded = round(value)
            if abs(value - rounded) < 1e-6:
                return int(rounded)
            return None
        text = _coerce_nonempty_str(value)
        if not text:
            return None
        match = re.search(r'(\d+(?:[.,]\d+)?)', text.replace(',', ''))
        if not match:
            return None
        try:
            number = float(match.group(1))
        except ValueError:
            return None
        if number < 0:
            return None
        if number.is_integer():
            return int(number)
        rounded = round(number)
        if abs(number - rounded) < 1e-6:
            return int(rounded)
        return None

    # Product name: grouped rows override everything else
    product_name = (
        get_value(detail_kv, 'NAME')
        or _coerce_nonempty_str(row_data.get('product_name'))
        or _coerce_nonempty_str(row_data.get('item_name'))
        or get_value(kv_specs, 'PRODUCT', 'NAME', 'RANGE')
    )

    doc_code = _coerce_nonempty_str(row_data.get('doc_code'))
    if not doc_code:
        combined = _coerce_nonempty_str(row_data.get('product_name'))
        if combined:
            for delim in (" - ", " – ", " — ", ": "):
                if delim not in combined:
                    continue
                left, right = combined.split(delim, 1)
                left = left.strip()
                right = right.strip()
                if not left or not right:
                    continue
                if len(left) > 30:
                    continue
                doc_code = left
                if product_name == combined:
                    product_name = right
                break
    if not doc_code:
        image_value = _coerce_nonempty_str(row_data.get('image'))
        if image_value:
            match = re.search(r'/([^/]+?)\.(?:jpg|jpeg|png|gif|webp)\b', image_value, re.IGNORECASE)
            if match:
                candidate = match.group(1).strip()
                if candidate and len(candidate) <= 30 and " " not in candidate:
                    doc_code = candidate

    # Brand: grouped Maker overrides, else manufacturer block NAME, else explicit keys.
    brand = (
        get_value(detail_kv, 'MAKER', 'BRAND', 'MANUFACTURER', 'SUPPLIER')
        or get_value(kv_manufacturer, 'NAME')
        or (_coerce_nonempty_str(row_data.get('manufacturer')) if not kv_manufacturer else None)
        or get_value(kv_specs, 'MAKER', 'BRAND', 'MANUFACTURER', 'SUPPLIER')
    )

    # Other scalar fields from specs/detail
    colour = get_value(detail_kv, 'COLOUR') or get_value(kv_specs, 'COLOUR') or _coerce_nonempty_str(row_data.get('colour'))
    finish = get_value(detail_kv, 'FINISH') or get_value(kv_specs, 'FINISH') or _coerce_nonempty_str(row_data.get('finish'))
    material = get_value(detail_kv, 'MATERIAL') or get_value(kv_specs, 'MATERIAL') or _coerce_nonempty_str(row_data.get('material'))

    def build_dimension_text(kv: dict[str, str]) -> str:
        lines: list[str] = []
        for key in ("WIDTH", "LENGTH", "HEIGHT", "DEPTH", "THICKNESS", "SIZE"):
            value = get_value(kv, key)
            if not value:
                continue
            if re.search(rf'^\s*{re.escape(key)}\b', value, re.IGNORECASE):
                lines.append(value)
            else:
                lines.append(f"{key}: {value}")
        return "\n".join(lines)

    detail_dim_text = build_dimension_text(detail_kv)
    specs_dim_text = build_dimension_text(kv_specs)

    detail_dims = parse_dimensions(detail_dim_text) if detail_dim_text else {"width": None, "length": None, "height": None}
    specs_dims = parse_dimensions(specs_dim_text) if specs_dim_text else {"width": None, "length": None, "height": None}
    width = _parse_mm_cell(row_data.get('width')) or detail_dims.get("width") or specs_dims.get("width")
    length = _parse_mm_cell(row_data.get('length')) or detail_dims.get("length") or specs_dims.get("length")
    height = _parse_mm_cell(row_data.get('height')) or detail_dims.get("height") or specs_dims.get("height")

    qty = _parse_qty(row_data.get('qty'))
    rrp = _parse_numeric_price(row_data.get('cost'))
    if rrp is None:
        # Some schedules store price as free-text in the cost column or within the specs/notes block.
        # Keep this as a fallback so numeric columns (sample3) always win.
        candidate_texts: list[Any] = [
            row_data.get('cost'),
            row_data.get('specs'),
            row_data.get('notes'),
            kv_specs.get('PRICE'),
            kv_specs.get('COST'),
            kv_specs.get('RRP'),
        ]
        for candidate in candidate_texts:
            rrp = parse_price(candidate if isinstance(candidate, str) else (str(candidate) if candidate is not None else None))
            if rrp is not None:
                break

    product_description = _build_product_description(
        section=_coerce_nonempty_str(row_data.get('section')),
        item_location=_coerce_nonempty_str(row_data.get('item_location')),
    )

    # Build product_details from remaining KV pairs.
    used_keys: set[str] = set()
    if product_name:
        # We might have sourced from any of these.
        used_keys.update({'PRODUCT', 'NAME', 'RANGE'})
    if brand:
        used_keys.update({'MAKER', 'BRAND', 'MANUFACTURER', 'SUPPLIER', 'NAME'})
    if colour:
        used_keys.add('COLOUR')
    if finish:
        used_keys.add('FINISH')
    if material:
        used_keys.add('MATERIAL')
    if width is not None:
        used_keys.add('WIDTH')
    if length is not None:
        used_keys.add('LENGTH')
    if height is not None:
        used_keys.update({'HEIGHT', 'DEPTH', 'THICKNESS'})
    if detail_dim_text or specs_dim_text:
        used_keys.add('SIZE')

    details_parts: list[str] = []
    specs_details = format_kv_as_details(kv_specs, exclude_keys=used_keys)
    if specs_details:
        details_parts.append(specs_details)

    manufacturer_exclude = set(used_keys)
    manufacturer_exclude.add('NAME')  # avoid duplicating the brand in details
    manufacturer_details = format_kv_as_details(kv_manufacturer, exclude_keys=manufacturer_exclude)
    if manufacturer_details:
        details_parts.append(manufacturer_details)

    detail_details = format_kv_as_details(detail_kv, exclude_keys=used_keys)
    if detail_details:
        details_parts.append(detail_details)

    product_details = ' | '.join(details_parts) if details_parts else None

    return Product(
        doc_code=doc_code,
        product_name=product_name,
        brand=brand,
        colour=colour,
        finish=finish,
        material=material,
        width=width,
        length=length,
        height=height,
        qty=qty,
        rrp=rrp,
        feature_image=None,
        product_description=product_description,
        product_details=product_details,
    )
