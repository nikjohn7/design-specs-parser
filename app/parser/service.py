"""Schedule parser service module.

Provides the ScheduleParser OOP service and ScheduleParserConfig dataclass
for configuring parser behavior, separating runtime config from app-level settings.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Literal

if TYPE_CHECKING:
    from openpyxl import Workbook

from app.core.models import ParseResponse, Product
from app.parser.llm_client import BaseLLMClient, NoopLLMClient
from app.parser.llm_extractor import ExtractionContext, ProductExtractor


@dataclass
class ScheduleParserConfig:
    """Configuration for the ScheduleParser service.

    Allows different parser configurations per request if needed,
    independent of global application settings.
    """

    use_llm: bool = False
    llm_mode: Literal["fallback", "refine"] = "fallback"
    llm_min_missing_fields: int = 3
    llm_batch_size: int = 5
    extract_images: bool = False


class ScheduleParser:
    """OOP service encapsulating the entire parsing pipeline.

    Usage:
        config = ScheduleParserConfig(use_llm=True, llm_mode="fallback")
        parser = ScheduleParser(config, llm_client=DeepInfraLLMClient(...))
        result = parser.parse_workbook(wb, filename="schedule.xlsx")

    When LLM is disabled (default), heuristic-only extraction is used.
    When LLM is enabled, the ProductExtractor layer enhances extraction
    for sparse products.
    """

    # Columns that indicate a sheet contains schedule data
    _SCHEDULE_SUPPORTING_COLS = {
        "item_location", "specs", "manufacturer", "notes", "qty", "cost", "product_name"
    }

    def __init__(
        self,
        config: ScheduleParserConfig | None = None,
        llm_client: BaseLLMClient | None = None,
    ):
        """Initialize the parser service.

        Args:
            config: Parser configuration. Uses defaults if not provided.
            llm_client: LLM client for enhanced extraction. Uses NoopLLMClient
                       (no-op) if not provided or if use_llm is False.
        """
        self.config = config or ScheduleParserConfig()
        self._llm_client = llm_client or NoopLLMClient()

        # Create the product extractor for LLM-enhanced extraction
        self._extractor = ProductExtractor(
            llm_client=self._llm_client,
            mode=self.config.llm_mode,
            min_missing_fields=self.config.llm_min_missing_fields,
            batch_size=self.config.llm_batch_size,
        )

    def parse_workbook(self, wb: "Workbook", filename: str) -> ParseResponse:
        """Parse a workbook and return structured product data.

        Orchestrates the full parsing pipeline:
          - Determine schedule name
          - Iterate schedule-like sheets
          - Fill merged cells
          - Detect header row and map columns
          - Extract raw product rows and parse into Product models
          - De-duplicate products by doc_code

        Args:
            wb: Loaded openpyxl Workbook object.
            filename: Original filename (used for schedule_name fallback).

        Returns:
            ParseResponse with schedule_name and extracted products.
        """
        from app.parser.column_mapper import map_columns
        from app.parser.field_parser import extract_product_fields, parse_kv_block
        from app.parser.merged_cells import fill_merged_regions
        from app.parser.row_extractor import iter_product_rows
        from app.parser.sheet_detector import find_header_row
        from app.parser.workbook import get_schedule_name

        schedule_name = get_schedule_name(wb, filename)

        # Collect products with their raw text and context for batch LLM extraction
        extraction_items: list[tuple[Product, str, ExtractionContext | None]] = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Fill merged regions so all cells have values
            try:
                fill_merged_regions(ws)
            except Exception:
                pass

            # Find header row; skip sheet if not found
            header_row = find_header_row(ws)
            if header_row is None:
                continue

            # Map columns to canonical names
            col_map = map_columns(ws, header_row=header_row)

            # Skip sheets without doc_code or product_name
            if "doc_code" not in col_map and "product_name" not in col_map:
                continue

            # Skip sheets without supporting schedule columns
            if not (set(col_map.keys()) & self._SCHEDULE_SUPPORTING_COLS):
                continue

            # Extract products from this sheet
            row_index = header_row
            for row_data in iter_product_rows(ws, header_row=header_row, col_map=col_map):
                row_index += 1

                if self._looks_like_repeated_header_row(row_data):
                    continue

                kv_specs = parse_kv_block(row_data.get("specs"))
                kv_manufacturer = parse_kv_block(row_data.get("manufacturer"))

                product = extract_product_fields(row_data, kv_specs, kv_manufacturer)

                # Skip empty products
                if not self._has_meaningful_data(product):
                    continue

                # Build raw text for LLM from row data
                raw_text = self._build_raw_text(row_data)
                context = ExtractionContext(
                    sheet_name=sheet_name,
                    row_index=row_index,
                    section=product.product_description,
                )

                extraction_items.append((product, raw_text, context))

        # Apply LLM extraction (batch) if enabled
        if self.config.use_llm and extraction_items:
            products = self._extractor.extract_batch(extraction_items)
        else:
            products = [item[0] for item in extraction_items]

        # De-duplicate by doc_code
        products = self._dedupe_products_by_doc_code(products)

        return ParseResponse(schedule_name=schedule_name, products=products)

    def _has_meaningful_data(self, product: Product) -> bool:
        """Check if product has at least one meaningful field populated."""
        return any((
            product.doc_code,
            product.product_name,
            product.brand,
            product.colour,
            product.finish,
            product.material,
            product.product_description,
            product.product_details,
        ))

    def _build_raw_text(self, row_data: dict[str, Any]) -> str:
        """Build raw text from row data for LLM analysis.

        Concatenates relevant cell values into a single string that
        the LLM can use to extract product information.
        """
        # Columns that contain useful text for LLM extraction
        text_columns = [
            "item_location",
            "specs",
            "manufacturer",
            "notes",
            "product_name",
        ]

        parts = []
        for col in text_columns:
            value = row_data.get(col)
            if value is not None:
                text = str(value).strip()
                if text:
                    parts.append(text)

        return " | ".join(parts)

    def _looks_like_repeated_header_row(self, row_data: dict[str, Any]) -> bool:
        """Detect header rows repeated mid-sheet (e.g., after page breaks)."""
        def norm(value: Any) -> str:
            if value is None:
                return ""
            return str(value).strip().lower()

        doc_code = norm(row_data.get("doc_code"))
        if not doc_code:
            return False

        # Known header-like values for doc_code column
        doc_code_headers = {
            "spec code", "doc code", "drawing code", "code", "ref", "ref no",
            "reference", "id", "sku", "item code", "product code",
        }
        if doc_code not in doc_code_headers:
            return False

        # Check for additional header-like cells
        header_patterns: dict[str, set[str]] = {
            "item_location": {"item & location", "item and location", "area", "room", "location", "description"},
            "specs": {"specification", "specifications", "specs", "notes/comments", "details", "spec"},
            "manufacturer": {"manufacturer", "supplier", "brand", "vendor", "maker", "manufacturer / supplier"},
            "notes": {"notes", "comments", "remarks"},
            "qty": {"qty", "quantity", "units", "no.", "no"},
            "cost": {"cost", "rrp", "price", "indicative cost", "cost per unit", "unit price", "unit cost", "$"},
        }

        header_count = sum(
            1 for key, values in header_patterns.items()
            if norm(row_data.get(key)) in values
        )

        return header_count >= 2

    def _dedupe_products_by_doc_code(self, products: list[Product]) -> list[Product]:
        """De-duplicate products by doc_code, keeping first occurrence.

        Products with None/empty doc_code are always kept.
        """
        seen: set[str] = set()
        deduped: list[Product] = []

        for product in products:
            doc_code_key = self._normalize_doc_code(product.doc_code)
            if doc_code_key is None:
                deduped.append(product)
                continue
            if doc_code_key in seen:
                continue
            seen.add(doc_code_key)
            deduped.append(product)

        return deduped

    def _normalize_doc_code(self, doc_code: str | None) -> str | None:
        """Normalize doc_code for deduplication."""
        if doc_code is None:
            return None
        normalized = doc_code.strip()
        return normalized or None
