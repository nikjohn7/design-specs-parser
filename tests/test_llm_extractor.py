"""Unit tests for the LLM extractor module.

Tests the ProductExtractor class using FakeLLMClient for deterministic behavior.
"""

import pytest

from app.core.models import Product
from app.parser.llm_client import FakeLLMClient, LLMProductPatch, NoopLLMClient
from app.parser.llm_extractor import ExtractionContext, ProductExtractor


class TestExtractionContext:
    """Tests for the ExtractionContext dataclass."""

    def test_required_fields(self):
        ctx = ExtractionContext(sheet_name="Schedule", row_index=10)
        assert ctx.sheet_name == "Schedule"
        assert ctx.row_index == 10
        assert ctx.section is None

    def test_optional_section(self):
        ctx = ExtractionContext(sheet_name="Schedule", row_index=10, section="FLOORING")
        assert ctx.section == "FLOORING"


class TestProductExtractorFallbackMode:
    """Tests for ProductExtractor in fallback mode."""

    def test_skips_llm_for_rich_products(self):
        """Products with few missing fields should not trigger LLM."""
        fake_client = FakeLLMClient(
            responses={"test": LLMProductPatch(product_name="LLM Name")}
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
        )

        # Product with only 2 missing fields (finish, material)
        rich_product = Product(
            doc_code="X1",
            product_name="Existing Name",
            brand="Existing Brand",
            colour="Blue",
        )

        result = extractor.extract(rich_product, "test product text")

        # Should not call LLM
        assert len(fake_client.calls) == 0
        # Original values preserved
        assert result.product_name == "Existing Name"
        assert result.brand == "Existing Brand"

    def test_calls_llm_for_sparse_products(self):
        """Products with many missing fields should trigger LLM."""
        fake_client = FakeLLMClient(
            responses={"sparse": LLMProductPatch(product_name="LLM Filled", brand="LLM Brand")}
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
        )

        # Product with 5 missing key fields (all None)
        sparse_product = Product(doc_code="X2")

        result = extractor.extract(sparse_product, "sparse product data")

        # Should call LLM
        assert len(fake_client.calls) == 1
        # LLM values should fill gaps
        assert result.product_name == "LLM Filled"
        assert result.brand == "LLM Brand"
        # Original doc_code preserved
        assert result.doc_code == "X2"

    def test_heuristic_values_not_overwritten(self):
        """LLM should only fill gaps, not overwrite existing values."""
        fake_client = FakeLLMClient(
            responses={
                "test": LLMProductPatch(
                    product_name="LLM Name",
                    brand="LLM Brand",
                    colour="LLM Colour",
                )
            }
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=1,  # Lower threshold to trigger LLM
        )

        # Product with some existing values
        product = Product(
            doc_code="X3",
            product_name="Original Name",  # Should NOT be overwritten
            brand=None,  # Should be filled
            colour=None,  # Should be filled
        )

        result = extractor.extract(product, "test data")

        assert result.product_name == "Original Name"  # Preserved
        assert result.brand == "LLM Brand"  # Filled
        assert result.colour == "LLM Colour"  # Filled

    def test_threshold_boundary(self):
        """Test exact threshold boundary (3 missing = calls LLM, 2 missing = skips)."""
        fake_client = FakeLLMClient(
            responses={"data": LLMProductPatch(finish="LLM Finish")}
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
        )

        # Exactly 3 missing: product_name, brand, colour (finish and material present)
        product_at_threshold = Product(
            doc_code="X4",
            finish="Matte",
            material="Wood",
        )

        result = extractor.extract(product_at_threshold, "data text")

        # Should call LLM (3 >= 3)
        assert len(fake_client.calls) == 1

    def test_context_passed_to_llm(self):
        """Verify context is converted and passed to LLM client."""
        fake_client = FakeLLMClient()
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=1,
        )

        sparse_product = Product(doc_code="X5")
        ctx = ExtractionContext(sheet_name="Sales", row_index=15, section="LIGHTING")

        extractor.extract(sparse_product, "raw text", ctx)

        assert len(fake_client.calls) == 1
        raw_text, context_dict = fake_client.calls[0]
        assert raw_text == "raw text"
        assert context_dict == {"sheet": "Sales", "row": 15, "section": "LIGHTING"}


class TestProductExtractorRefineMode:
    """Tests for ProductExtractor in refine mode."""

    def test_always_calls_llm(self):
        """Refine mode should always call LLM regardless of missing fields."""
        fake_client = FakeLLMClient(
            responses={"text": LLMProductPatch(colour="LLM Colour")}
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="refine",
            min_missing_fields=3,
        )

        # Rich product with only 1 missing field
        rich_product = Product(
            doc_code="X1",
            product_name="Name",
            brand="Brand",
            colour="Original Colour",
            finish="Finish",
        )

        result = extractor.extract(rich_product, "text data")

        # Should still call LLM
        assert len(fake_client.calls) == 1
        # Original colour preserved (LLM only fills gaps)
        assert result.colour == "Original Colour"


class TestProductExtractorBatchExtraction:
    """Tests for batch extraction functionality."""

    def test_batch_extraction_basic(self):
        """Test basic batch extraction with mixed sparse/rich products."""
        fake_client = FakeLLMClient(
            responses={
                "sparse1": LLMProductPatch(product_name="Filled 1"),
                "sparse2": LLMProductPatch(product_name="Filled 2"),
            }
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
            batch_size=10,
        )

        items = [
            (Product(doc_code="A1"), "sparse1 text", None),  # Sparse - needs LLM
            (Product(doc_code="A2", product_name="Existing", brand="B", colour="C"), "rich text", None),  # Rich - skips LLM
            (Product(doc_code="A3"), "sparse2 text", None),  # Sparse - needs LLM
        ]

        results = extractor.extract_batch(items)

        # Two LLM calls (for sparse products)
        assert len(fake_client.calls) == 2
        # Results in order
        assert len(results) == 3
        assert results[0].product_name == "Filled 1"
        assert results[1].product_name == "Existing"  # Unchanged
        assert results[2].product_name == "Filled 2"

    def test_batch_respects_batch_size(self):
        """Test that batching respects batch_size parameter."""
        call_count = 0
        original_extract_batch = FakeLLMClient.extract_batch

        def counting_extract_batch(self, items):
            nonlocal call_count
            call_count += 1
            return original_extract_batch(self, items)

        fake_client = FakeLLMClient(
            responses={"sparse": LLMProductPatch(product_name="Filled")}
        )
        # Monkey-patch to count batch calls
        fake_client.extract_batch = lambda items: counting_extract_batch(fake_client, items)

        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
            batch_size=2,
        )

        # 5 sparse products with batch_size=2 should make 3 batch calls
        items = [
            (Product(doc_code=f"X{i}"), "sparse text", None)
            for i in range(5)
        ]

        extractor.extract_batch(items)

        assert call_count == 3  # ceil(5/2) = 3 batch calls

    def test_batch_empty_list(self):
        """Test batch extraction with empty input."""
        extractor = ProductExtractor(llm_client=NoopLLMClient())
        results = extractor.extract_batch([])
        assert results == []

    def test_batch_no_llm_needed(self):
        """Test batch when no products need LLM enhancement."""
        fake_client = FakeLLMClient()
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=3,
        )

        # All rich products
        items = [
            (Product(doc_code="A", product_name="N", brand="B", colour="C"), "text", None),
            (Product(doc_code="B", product_name="N", brand="B", colour="C"), "text", None),
        ]

        results = extractor.extract_batch(items)

        # No LLM calls
        assert len(fake_client.calls) == 0
        assert len(results) == 2

    def test_batch_with_context(self):
        """Test batch extraction passes context correctly."""
        fake_client = FakeLLMClient()
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=1,
        )

        ctx1 = ExtractionContext(sheet_name="Sheet1", row_index=5)
        ctx2 = ExtractionContext(sheet_name="Sheet2", row_index=10, section="FLOORING")

        items = [
            (Product(doc_code="A"), "text1", ctx1),
            (Product(doc_code="B"), "text2", ctx2),
        ]

        extractor.extract_batch(items)

        # Verify contexts passed correctly
        assert len(fake_client.calls) == 2
        assert fake_client.calls[0][1] == {"sheet": "Sheet1", "row": 5, "section": None}
        assert fake_client.calls[1][1] == {"sheet": "Sheet2", "row": 10, "section": "FLOORING"}


class TestProductExtractorMerging:
    """Tests for the merge behavior."""

    def test_merge_fills_all_none_fields(self):
        """Verify all None fields get filled from patch."""
        fake_client = FakeLLMClient(
            responses={
                "test": LLMProductPatch(
                    product_name="Name",
                    brand="Brand",
                    colour="Colour",
                    finish="Finish",
                    material="Material",
                    width=100,
                    length=200,
                    height=50,
                    qty=10,
                    rrp=99.99,
                )
            }
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=1,
        )

        empty_product = Product(doc_code="X1")
        result = extractor.extract(empty_product, "test data")

        assert result.product_name == "Name"
        assert result.brand == "Brand"
        assert result.colour == "Colour"
        assert result.finish == "Finish"
        assert result.material == "Material"
        assert result.width == 100
        assert result.length == 200
        assert result.height == 50
        assert result.qty == 10
        assert result.rrp == 99.99
        # Original field preserved
        assert result.doc_code == "X1"

    def test_merge_preserves_non_none_fields(self):
        """Verify existing non-None fields are never overwritten."""
        fake_client = FakeLLMClient(
            responses={
                "test": LLMProductPatch(
                    product_name="LLM Name",
                    width=9999,
                    rrp=1.00,
                )
            }
        )
        extractor = ProductExtractor(
            llm_client=fake_client,
            mode="fallback",
            min_missing_fields=1,
        )

        product = Product(
            doc_code="X1",
            product_name="Original Name",
            width=500,
        )
        result = extractor.extract(product, "test data")

        # Original values preserved
        assert result.product_name == "Original Name"
        assert result.width == 500
        # Only None fields filled
        assert result.rrp == 1.00


class TestCountMissingFields:
    """Tests for the _count_missing_fields helper."""

    def test_all_missing(self):
        """Product with no key fields set should count 5."""
        extractor = ProductExtractor()
        product = Product(doc_code="X1")
        assert extractor._count_missing_fields(product) == 5

    def test_none_missing(self):
        """Product with all key fields set should count 0."""
        extractor = ProductExtractor()
        product = Product(
            doc_code="X1",
            product_name="Name",
            brand="Brand",
            colour="Colour",
            finish="Finish",
            material="Material",
        )
        assert extractor._count_missing_fields(product) == 0

    def test_partial_missing(self):
        """Product with some key fields set."""
        extractor = ProductExtractor()
        product = Product(
            doc_code="X1",
            product_name="Name",
            brand="Brand",
        )
        # Missing: colour, finish, material
        assert extractor._count_missing_fields(product) == 3
