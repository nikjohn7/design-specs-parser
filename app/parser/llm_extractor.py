"""LLM-based product extraction with heuristic fallback.

This module provides the ProductExtractor class which orchestrates
extraction using heuristics and optionally calls an LLM to fill in
missing fields for sparse products.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TYPE_CHECKING

from app.core.models import Product
from app.parser.llm_client import BaseLLMClient, LLMProductPatch, NoopLLMClient

if TYPE_CHECKING:
    pass


@dataclass
class ExtractionContext:
    """Metadata about where a product was extracted from.

    This context is passed to the LLM to help it understand
    the source of the raw text and make better extraction decisions.
    """

    sheet_name: str
    row_index: int
    section: str | None = None


class ProductExtractor:
    """Orchestrates product field extraction using heuristics + optional LLM.

    Supports two modes:
    - fallback: Use heuristics first, call LLM only if product is too sparse
    - refine: Always call LLM to validate/enhance heuristic results

    Example:
        extractor = ProductExtractor(
            llm_client=my_client,
            mode="fallback",
            min_missing_fields=3,
        )
        enhanced_product = extractor.extract(heuristic_product, raw_text, context)
    """

    def __init__(
        self,
        llm_client: BaseLLMClient | None = None,
        mode: Literal["fallback", "refine"] = "fallback",
        min_missing_fields: int = 3,
        batch_size: int = 5,
    ) -> None:
        """Initialize the extractor.

        Args:
            llm_client: LLM client to use for extraction. Defaults to NoopLLMClient.
            mode: Extraction mode - "fallback" or "refine".
            min_missing_fields: Threshold for calling LLM in fallback mode.
            batch_size: Number of products per batch LLM call.
        """
        self.llm_client = llm_client or NoopLLMClient()
        self.mode = mode
        self.min_missing_fields = min_missing_fields
        self.batch_size = batch_size

    def extract(
        self,
        heuristic_product: Product,
        raw_text: str,
        context: ExtractionContext | None = None,
    ) -> Product:
        """Extract/enhance a single product.

        In fallback mode, the LLM is only called if the product has
        too many missing fields (>= min_missing_fields threshold).
        In refine mode, the LLM is always called.

        Args:
            heuristic_product: Product already extracted via heuristics.
            raw_text: Raw text source for LLM to analyze.
            context: Optional metadata about the extraction source.

        Returns:
            Enhanced Product with LLM-filled fields merged in.
        """
        if self.mode == "refine":
            patch = self._get_llm_patch(raw_text, context)
            return self._merge_patch(heuristic_product, patch, prefer_llm=False)

        # Fallback mode: only call LLM if too many fields missing
        missing_count = self._count_missing_fields(heuristic_product)
        if missing_count >= self.min_missing_fields:
            patch = self._get_llm_patch(raw_text, context)
            return self._merge_patch(heuristic_product, patch, prefer_llm=True)

        return heuristic_product

    def extract_batch(
        self,
        items: list[tuple[Product, str, ExtractionContext | None]],
    ) -> list[Product]:
        """Extract/enhance multiple products using batched LLM calls.

        Args:
            items: List of (heuristic_product, raw_text, context) tuples.

        Returns:
            List of enhanced Products in the same order as input.
        """
        if not items:
            return []

        # Determine which items need LLM based on mode
        if self.mode == "refine":
            needs_llm = list(range(len(items)))
        else:
            needs_llm = [
                i
                for i, (prod, _, _) in enumerate(items)
                if self._count_missing_fields(prod) >= self.min_missing_fields
            ]

        # If no items need LLM, return products as-is
        if not needs_llm:
            return [prod for prod, _, _ in items]

        # Prepare LLM inputs for items that need enhancement
        llm_inputs = [
            (items[i][1], self._context_to_dict(items[i][2])) for i in needs_llm
        ]

        # Batch LLM calls according to batch_size
        patches: list[LLMProductPatch] = []
        for batch_start in range(0, len(llm_inputs), self.batch_size):
            batch = llm_inputs[batch_start : batch_start + self.batch_size]
            batch_patches = self.llm_client.extract_batch(batch)
            patches.extend(batch_patches)

        # Build results by merging patches where needed
        results: list[Product] = []
        patch_idx = 0
        for i, (prod, _, _) in enumerate(items):
            if i in needs_llm:
                patch = (
                    patches[patch_idx] if patch_idx < len(patches) else LLMProductPatch()
                )
                patch_idx += 1
                # In fallback mode, LLM fills gaps (prefer_llm=True means fill only)
                prefer_llm = self.mode == "fallback"
                results.append(self._merge_patch(prod, patch, prefer_llm=prefer_llm))
            else:
                results.append(prod)

        return results

    def _get_llm_patch(
        self,
        raw_text: str,
        context: ExtractionContext | None,
    ) -> LLMProductPatch:
        """Call LLM to extract a patch from raw text."""
        ctx_dict = self._context_to_dict(context)
        return self.llm_client.extract_product_patch(raw_text, ctx_dict)

    def _context_to_dict(self, context: ExtractionContext | None) -> dict | None:
        """Convert ExtractionContext to a dict for the LLM client."""
        if not context:
            return None
        return {
            "sheet": context.sheet_name,
            "row": context.row_index,
            "section": context.section,
        }

    def _count_missing_fields(self, product: Product) -> int:
        """Count how many key fields are missing (None) on a product.

        Key fields are the descriptive fields that help identify a product:
        product_name, brand, colour, finish, material.
        """
        key_fields = [
            product.product_name,
            product.brand,
            product.colour,
            product.finish,
            product.material,
        ]
        return sum(1 for f in key_fields if f is None)

    def _merge_patch(
        self,
        product: Product,
        patch: LLMProductPatch,
        prefer_llm: bool = False,
    ) -> Product:
        """Merge an LLM patch into a product.

        Args:
            product: Base product from heuristic extraction.
            patch: Partial fields from LLM.
            prefer_llm: If True, LLM values only fill gaps (heuristic wins).
                       If False, also only fills gaps (refine mode).

        Returns:
            New Product with merged fields.
        """
        data = product.model_dump()
        patch_data = patch.model_dump(exclude_none=True)

        for field, value in patch_data.items():
            # Only fill if heuristic value is None
            if data.get(field) is None:
                data[field] = value

        return Product.model_validate(data)
