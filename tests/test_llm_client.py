"""Unit tests for LLM client abstractions."""

from dataclasses import dataclass

import pytest

from app.parser.llm_client import (
    BaseLLMClient,
    FakeLLMClient,
    LLMProductPatch,
    NoopLLMClient,
    build_llm_client,
)


class TestLLMProductPatch:
    """Tests for the LLMProductPatch model."""

    def test_empty_patch_has_all_none(self):
        patch = LLMProductPatch()
        assert patch.product_name is None
        assert patch.brand is None
        assert patch.colour is None
        assert patch.finish is None
        assert patch.material is None
        assert patch.width is None
        assert patch.length is None
        assert patch.height is None
        assert patch.qty is None
        assert patch.rrp is None

    def test_patch_with_partial_fields(self):
        patch = LLMProductPatch(brand="TestBrand", qty=5)
        assert patch.brand == "TestBrand"
        assert patch.qty == 5
        assert patch.product_name is None


class TestNoopLLMClient:
    """Tests for NoopLLMClient."""

    def test_returns_empty_patch(self):
        client = NoopLLMClient()
        patch = client.extract_product_patch("Some product text here")
        assert patch == LLMProductPatch()

    def test_returns_empty_patch_with_context(self):
        client = NoopLLMClient()
        patch = client.extract_product_patch(
            "Product description",
            context={"sheet": "Schedule", "row": 10},
        )
        assert patch == LLMProductPatch()

    def test_extract_batch_returns_empty_patches(self):
        client = NoopLLMClient()
        items = [
            ("Product 1", None),
            ("Product 2", {"sheet": "A"}),
            ("Product 3", {"row": 5}),
        ]
        patches = client.extract_batch(items)
        assert len(patches) == 3
        assert all(p == LLMProductPatch() for p in patches)


class TestFakeLLMClient:
    """Tests for FakeLLMClient."""

    def test_returns_configured_response(self):
        client = FakeLLMClient(
            responses={"ICONIC": LLMProductPatch(product_name="Iconic Carpet")}
        )
        patch = client.extract_product_patch("Product: ICONIC carpet tile")
        assert patch.product_name == "Iconic Carpet"

    def test_records_calls(self):
        client = FakeLLMClient(
            responses={"MATCH": LLMProductPatch(brand="TestBrand")}
        )
        client.extract_product_patch("First call with MATCH")
        client.extract_product_patch("Second call", context={"sheet": "A"})

        assert len(client.calls) == 2
        assert client.calls[0] == ("First call with MATCH", None)
        assert client.calls[1] == ("Second call", {"sheet": "A"})

    def test_returns_empty_for_unmatched_input(self):
        client = FakeLLMClient(
            responses={"SPECIAL": LLMProductPatch(brand="SpecialBrand")}
        )
        patch = client.extract_product_patch("Regular product text")
        assert patch == LLMProductPatch()

    def test_matches_first_key_found(self):
        client = FakeLLMClient(
            responses={
                "AAA": LLMProductPatch(brand="BrandA"),
                "BBB": LLMProductPatch(brand="BrandB"),
            }
        )
        # Text contains both keys; first configured one should match.
        patch = client.extract_product_patch("Text with AAA and BBB")
        assert patch.brand == "BrandA"

    def test_extract_batch_with_mixed_matches(self):
        client = FakeLLMClient(
            responses={
                "CARPET": LLMProductPatch(product_name="Carpet Product"),
                "TILE": LLMProductPatch(product_name="Tile Product"),
            }
        )
        items = [
            ("CARPET description", None),
            ("Unknown product", None),
            ("TILE specs", {"row": 3}),
        ]
        patches = client.extract_batch(items)

        assert len(patches) == 3
        assert patches[0].product_name == "Carpet Product"
        assert patches[1] == LLMProductPatch()
        assert patches[2].product_name == "Tile Product"
        assert len(client.calls) == 3

    def test_empty_responses_dict(self):
        client = FakeLLMClient()
        patch = client.extract_product_patch("Any text")
        assert patch == LLMProductPatch()
        assert len(client.calls) == 1


class TestBuildLLMClient:
    """Tests for the build_llm_client factory."""

    @dataclass
    class MockSettings:
        """Simple mock settings object for testing."""

        use_llm: bool = False
        llm_provider: str = "deepinfra"
        deepinfra_api_key: str | None = None
        llm_model: str = "test-model"

    def test_returns_noop_when_llm_disabled(self):
        settings = self.MockSettings(use_llm=False)
        client = build_llm_client(settings)
        assert isinstance(client, NoopLLMClient)

    def test_returns_noop_for_deepinfra_placeholder(self):
        # DeepInfra not yet implemented, so factory returns NoopLLMClient.
        settings = self.MockSettings(use_llm=True, llm_provider="deepinfra")
        client = build_llm_client(settings)
        assert isinstance(client, NoopLLMClient)

    def test_returns_noop_for_unknown_provider(self):
        settings = self.MockSettings(use_llm=True, llm_provider="unknown_provider")
        client = build_llm_client(settings)
        assert isinstance(client, NoopLLMClient)

    def test_factory_returns_base_llm_client_subclass(self):
        settings = self.MockSettings(use_llm=False)
        client = build_llm_client(settings)
        assert isinstance(client, BaseLLMClient)
