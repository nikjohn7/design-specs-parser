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


class NoopLLMClient(BaseLLMClient):
    """LLM client implementation that always returns an empty patch.

    This is the default client used when LLM integration is disabled.
    It allows the rest of the pipeline to depend on the LLM interface
    without making any external calls or changing behavior.
    """

    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None,
    ) -> LLMProductPatch:
        """Return an empty patch regardless of input."""
        return LLMProductPatch()


class FakeLLMClient(BaseLLMClient):
    """Test LLM client that returns predictable, configurable patches.

    The client is configured with a mapping from substring -> patch. On
    each call, it records the request and returns the first patch whose
    key is found as a substring in the input text. If no key matches,
    an empty patch is returned.
    """

    def __init__(self, responses: dict[str, LLMProductPatch] | None = None) -> None:
        # Simple substring-to-patch mapping used for deterministic tests.
        self.responses: dict[str, LLMProductPatch] = responses or {}
        # Record of calls made: (raw_text, context) tuples.
        self.calls: list[tuple[str, dict | None]] = []

    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None,
    ) -> LLMProductPatch:
        """Return a configured patch and record the call.

        Args:
            raw_text: Free-form product description text.
            context: Optional structured context, passed through for
                observability but not used in matching.
        """
        # Track every call for assertions in tests.
        self.calls.append((raw_text, context))

        # Find the first configured key that appears in the input text.
        for key, patch in self.responses.items():
            if key in raw_text:
                return patch

        # No configured response matched; return an empty patch.
        return LLMProductPatch()


def build_llm_client(settings) -> BaseLLMClient:
    """Factory function to create appropriate LLM client from settings.

    Args:
        settings: Application settings object with use_llm, llm_provider,
            and credential fields.

    Returns:
        A BaseLLMClient instance appropriate for the current configuration.
    """
    import logging

    logger = logging.getLogger(__name__)

    if not settings.use_llm:
        return NoopLLMClient()

    if settings.llm_provider == "deepinfra":
        # DeepInfraLLMClient will be implemented in Task 4.1.
        # For now, return NoopLLMClient as a placeholder.
        logger.info("DeepInfra provider selected but not yet implemented; using NoopLLMClient")
        return NoopLLMClient()

    logger.warning(f"Unknown LLM provider: {settings.llm_provider}, using NoopLLMClient")
    return NoopLLMClient()
