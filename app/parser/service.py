"""Schedule parser service module.

Provides the ScheduleParser OOP service and ScheduleParserConfig dataclass
for configuring parser behavior, separating runtime config from app-level settings.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from openpyxl import Workbook

    from app.core.models import ParseResponse


@dataclass
class ScheduleParserConfig:
    """Configuration for the ScheduleParser service.

    Allows different parser configurations per request if needed,
    independent of global application settings.
    """

    use_llm: bool = False
    llm_mode: Literal["fallback", "refine"] = "fallback"
    llm_min_missing_fields: int = 3
    extract_images: bool = False


class ScheduleParser:
    """OOP service encapsulating the entire parsing pipeline.

    Usage:
        config = ScheduleParserConfig(use_llm=True, llm_mode="fallback")
        parser = ScheduleParser(config, llm_client=DeepInfraLLMClient(...))
        result = parser.parse_workbook(wb, filename="schedule.xlsx")

    When LLM is disabled (default), this delegates to the heuristic-only
    parser in workbook.py. When LLM is enabled, the ProductExtractor
    layer enhances extraction for sparse products.
    """

    def __init__(
        self,
        config: ScheduleParserConfig | None = None,
        llm_client: object | None = None,  # BaseLLMClient once llm_client.py exists
    ):
        """Initialize the parser service.

        Args:
            config: Parser configuration. Uses defaults if not provided.
            llm_client: LLM client for enhanced extraction. Uses NoopLLMClient
                       (no-op) if not provided or if use_llm is False.
        """
        self.config = config or ScheduleParserConfig()
        self._llm_client = llm_client

    def parse_workbook(self, wb: Workbook, filename: str) -> ParseResponse:
        """Parse a workbook and return structured product data.

        Args:
            wb: Loaded openpyxl Workbook object.
            filename: Original filename (used for schedule_name fallback).

        Returns:
            ParseResponse with schedule_name and extracted products.
        """
        # Import here to avoid circular imports
        from app.parser.workbook import parse_workbook as heuristic_parse

        # For now, delegate to the heuristic parser.
        # LLM enhancement will be integrated in Step 4-5.
        return heuristic_parse(
            wb,
            filename=filename,
            extract_images=self.config.extract_images,
        )
