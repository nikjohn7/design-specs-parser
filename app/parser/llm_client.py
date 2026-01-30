"""LLM client abstractions for the schedule parser.

This module defines the data model used for partial product updates
returned by an LLM, and the abstract base class that concrete LLM
clients must implement.
"""

from __future__ import annotations

from abc import ABC, abstractmethod

from pydantic import BaseModel


class LLMProductPatch(BaseModel):
    """Partial product fields returned by LLM extraction.

    All fields are optional so that patches can be merged into
    an existing Product instance, only filling values that are
    present in the LLM output.
    """

    product_name: str | None = None
    brand: str | None = None
    colour: str | None = None
    finish: str | None = None
    material: str | None = None
    width: int | None = None
    length: int | None = None
    height: int | None = None
    qty: int | None = None
    rrp: float | None = None


class BaseLLMClient(ABC):
    """Abstract base class for LLM-backed product extractors.

    Concrete implementations can call different providers (or none at all)
    but must return an ``LLMProductPatch`` instance.
    """

    @abstractmethod
    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None,
    ) -> LLMProductPatch:
        """Extract a partial product representation from raw text.

        Args:
            raw_text: Source text describing a product (e.g. a row or cell).
            context: Optional structured context about where the text came
                from, such as sheet name or row index.
        """

    def extract_batch(
        self,
        items: list[tuple[str, dict | None]],
    ) -> list[LLMProductPatch]:
        """Extract patches for multiple items.

        The default implementation simply iterates and calls
        :meth:`extract_product_patch` for each item. Subclasses can
        override this method to perform real batch calls when the
        underlying provider supports it.
        """
        return [self.extract_product_patch(text, ctx) for text, ctx in items]
