"""Normalization utilities for schedule parsing.

This module provides parsing/normalization helpers for messy schedule text.
"""

from __future__ import annotations

import re


_UNIT_PATTERN = re.compile(
    r'^(?P<num>\d+(?:[.,]\d+)?)\s*(?P<unit>mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|m|met(?:er|re)s?)?$',
    re.IGNORECASE,
)


def _to_mm(value: str, unit: str | None) -> int | None:
    value = value.strip().replace(',', '.')
    try:
        number = float(value)
    except ValueError:
        return None

    if number < 0:
        return None

    if not unit:
        return int(round(number))

    unit_norm = unit.strip().lower()
    if unit_norm in {"mm", "millimeter", "millimeters", "millimetre", "millimetres"}:
        return int(round(number))
    if unit_norm in {"cm", "centimeter", "centimeters", "centimetre", "centimetres"}:
        return int(round(number * 10))
    if unit_norm in {"m", "meter", "meters", "metre", "metres"}:
        return int(round(number * 1000))

    return None


def _parse_number_with_unit(text: str) -> int | None:
    text = text.strip()
    if not text:
        return None

    # Handle glued forms like "10MM"
    glued = re.match(r'^(\d+(?:[.,]\d+)?)(mm|cm|m)$', text, re.IGNORECASE)
    if glued:
        return _to_mm(glued.group(1), glued.group(2))

    match = _UNIT_PATTERN.match(text)
    if not match:
        # Try to salvage a number+unit from within larger strings.
        inner = re.search(
            r'(\d+(?:[.,]\d+)?)\s*(mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|m|met(?:er|re)s?)\b',
            text,
            re.IGNORECASE,
        )
        if inner:
            return _to_mm(inner.group(1), inner.group(2))
        inner_num = re.search(r'(\d+(?:[.,]\d+)?)', text)
        if inner_num:
            return _to_mm(inner_num.group(1), None)
        return None

    return _to_mm(match.group("num"), match.group("unit"))


def parse_dimensions(text: str | None) -> dict[str, int | None]:
    """Parse dimension text into width/length/height (mm).

    Supports the patterns listed in TASKS.md (3.5):
    1) Explicit keys: WIDTH/LENGTH/HEIGHT/DEPTH/THICKNESS
    2) Labeled WxH / WxL blocks: "600 W X 600 H MM"
    3) Unlabeled sheet-size: "5500 X 2800 MM" (and 3-part "A X B X C MM")

    Unit conversion:
      - m/metre(s)/meter(s) → mm (×1000)
      - cm/centimetre(s)/centimeter(s) → mm (×10)
      - mm/millimetre(s)/millimeter(s) → mm

    Args:
        text: Raw dimension text (may contain additional words)

    Returns:
        Dict with keys: width, length, height (int mm or None)
    """
    result: dict[str, int | None] = {"width": None, "length": None, "height": None}
    if not text:
        return result

    normalized = str(text).replace("×", "X")

    # Pattern 1: explicit keys
    explicit: dict[str, int | None] = {}
    for key in ("WIDTH", "LENGTH", "HEIGHT", "DEPTH", "THICKNESS"):
        match = re.search(
            rf'\b{key}\b\s*[:=\-]?\s*([0-9]+(?:[.,][0-9]+)?\s*(?:mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|m|met(?:er|re)s?)?)',
            normalized,
            re.IGNORECASE,
        )
        if not match:
            continue
        value_mm = _parse_number_with_unit(match.group(1))
        if value_mm is None:
            continue
        explicit[key.upper()] = value_mm

    if "WIDTH" in explicit:
        result["width"] = explicit["WIDTH"]
    if "LENGTH" in explicit:
        result["length"] = explicit["LENGTH"]

    if "HEIGHT" in explicit:
        result["height"] = explicit["HEIGHT"]
    elif "DEPTH" in explicit:
        result["height"] = explicit["DEPTH"]
    elif "THICKNESS" in explicit:
        result["height"] = explicit["THICKNESS"]

    # Pattern 2: labeled blocks like "220 W X 2200 L MM" (optionally 3-part)
    # Prefer this only for missing fields so explicit keys win.
    labeled = re.search(
        r'(\d+(?:[.,]\d+)?)\s*([WLHDT])\s*X\s*(\d+(?:[.,]\d+)?)\s*([WLHDT])(?:\s*X\s*(\d+(?:[.,]\d+)?)\s*([WLHDT]))?\s*(mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|m|met(?:er|re)s?)?\b',
        normalized,
        re.IGNORECASE,
    )
    if labeled:
        a_num, a_label, b_num, b_label = labeled.group(1), labeled.group(2), labeled.group(3), labeled.group(4)
        c_num, c_label = labeled.group(5), labeled.group(6)
        unit = labeled.group(7)

        def set_labeled(num: str, label: str) -> None:
            mm = _to_mm(num, unit)
            if mm is None:
                return
            label_norm = label.upper()
            if label_norm == "W" and result["width"] is None:
                result["width"] = mm
            elif label_norm == "L" and result["length"] is None:
                result["length"] = mm
            elif label_norm in {"H", "D", "T"} and result["height"] is None:
                result["height"] = mm

        set_labeled(a_num, a_label)
        set_labeled(b_num, b_label)
        if c_num and c_label:
            set_labeled(c_num, c_label)

    # Pattern 3: unlabeled "A X B (X C) MM"
    if result["width"] is None or result["length"] is None or result["height"] is None:
        unlabeled = re.search(
            r'(\d+(?:[.,]\d+)?)\s*X\s*(\d+(?:[.,]\d+)?)(?:\s*X\s*(\d+(?:[.,]\d+)?))?\s*(mm|millimet(?:er|re)s?|cm|centimet(?:er|re)s?|m|met(?:er|re)s?)\b',
            normalized,
            re.IGNORECASE,
        )
        if unlabeled:
            a_mm = _to_mm(unlabeled.group(1), unlabeled.group(4))
            b_mm = _to_mm(unlabeled.group(2), unlabeled.group(4))
            c_mm = _to_mm(unlabeled.group(3), unlabeled.group(4)) if unlabeled.group(3) else None
            if result["width"] is None:
                result["width"] = a_mm
            if result["length"] is None:
                result["length"] = b_mm
            if result["height"] is None and c_mm is not None:
                result["height"] = c_mm

    return result


_NON_NUMERIC_PRICE_PATTERN = re.compile(
    r'^\s*(?:tbc|tba|poa|n/?a|na|nil|-\s*)\s*$',
    re.IGNORECASE,
)

# First preference: explicit currency marker like "$25+GST" or "$45.50 PER SQM"
_DOLLAR_AMOUNT_PATTERN = re.compile(
    r'\$\s*(?P<num>\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)',
    re.IGNORECASE,
)

# Fallback: amount near a price context word (RRP/PRICE/COST) when "$" is absent.
_CONTEXT_AMOUNT_PATTERN = re.compile(
    r'\b(?:rrp|price|cost|unit\s*cost|rate)\b[^\d$]{0,20}(?P<num>\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)',
    re.IGNORECASE,
)


def parse_price(text: str | None) -> float | None:
    """Parse a unit price from messy schedule text.

    Task 3.6 scope:
    - Extract numeric value from patterns like "$45.50", "$25+GST", "$X PER SQM".
    - Ignore non-numeric tokens like "TBC", "POA", "N/A" (return None).
    - Handle empty/None input (return None).

    Notes:
    - Prefer using numeric price columns when available; this function is for
      free-text cases where the sheet stores prices as strings.
    - This function is intentionally conservative to avoid mis-parsing unrelated
      numbers (e.g., dimensions, phone numbers).
    """
    if text is None:
        return None

    raw = str(text).strip()
    if not raw:
        return None

    if _NON_NUMERIC_PRICE_PATTERN.match(raw):
        return None

    match = _DOLLAR_AMOUNT_PATTERN.search(raw)
    if not match:
        match = _CONTEXT_AMOUNT_PATTERN.search(raw)
    if not match:
        return None

    num_text = match.group("num").replace(",", "").strip()
    try:
        value = float(num_text)
    except ValueError:
        return None

    if value < 0:
        return None

    return value
