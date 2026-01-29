"""Application settings.

Centralizes configuration (especially LLM toggles) so the rest of the app can
depend on a single settings object rather than scattered env reads.
"""

from __future__ import annotations

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application configuration loaded from environment variables and .env."""

    # LLM Configuration
    use_llm: bool = False
    llm_mode: str = "fallback"  # "fallback" | "refine"
    llm_provider: str = "deepinfra"
    llm_model: str = "openai/gpt-oss-120b"

    # Provider credentials
    deepinfra_api_key: str | None = None

    # LLM behavior thresholds
    llm_min_missing_fields: int = 3
    llm_batch_size: int = 5

    # Load variables from a local .env file when present.
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
    )


# Singleton settings instance used throughout the application.
settings = Settings()
