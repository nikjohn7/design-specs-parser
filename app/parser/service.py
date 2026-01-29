"""Schedule parser service module.

Provides the ScheduleParserConfig dataclass for configuring parser behavior,
separating runtime parser config from app-level settings.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


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
