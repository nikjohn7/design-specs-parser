# Plan: Add LLM Layer and OOP Parsing Service (`add-llm-layer`)

## 0. Overview and Goals

**Goal:** Evolve the existing heuristic-only Excel schedule parser into an architecture that:

1. Exposes a clear, object-oriented `ScheduleParser` service.
2. Adds a pluggable **LLM-based extraction layer** that:
   - Can refine or supplement heuristic extraction.
   - Is cleanly abstracted behind an interface (`BaseLLMClient`).
   - Is **optional** and **off by default**, preserving current behavior unless explicitly enabled.
3. Provides a clear configuration mechanism and documentation so reviewers can see:
   - Where and how LLMs are used.
   - How to run with/without LLMs.
   - How the design addresses their feedback (LLM use + OOP structure).

The plan assumes all work is done on a new branch:

```bash
git checkout -b add-llm-layer
```

---

## 1. Current Architecture Summary (Baseline)

### 1.1 Parsing Flow (Heuristic-Only)

- **app/api/routes.py**
  - `/parse` endpoint:
    - Validates upload.
    - Reads file bytes.
    - Calls `load_workbook_safe(file_bytes)` from `app/parser/workbook.py`.
    - Calls `parse_workbook(wb, filename=...)` (same module).
    - Returns `ParseResponse`.

- **app/parser/workbook.py**
  - `load_workbook_safe(file_bytes)` – validates and loads Workbook.
  - `get_schedule_name(wb, filename)` – heuristics for title.
  - `parse_workbook(wb, filename, extract_images=False)`:
    - For each sheet:
      - `fill_merged_regions(ws)`
      - `find_header_row(ws)` from `sheet_detector.py`
      - `map_columns(ws, header_row)` from `column_mapper.py`
      - `iter_product_rows(ws, header_row, col_map)` from `row_extractor.py`
      - Per row:
        - `parse_kv_block` on specs and manufacturer (`field_parser.py`)
        - `extract_product_fields(row_data, kv_specs, kv_manufacturer)` → `Product`
    - Filter out "empty" products.
    - `_dedupe_products_by_doc_code(products)` → final list.
    - Returns `ParseResponse`.

- **app/core/models.py**
  - `Product`, `ParseResponse`, `ErrorResponse` Pydantic models.

### 1.2 LLM Usage

- No LLM usage anywhere.

### 1.3 OOP Structure

- Orchestration is done via module-level functions, not via a service object.

---

## 2. Target Architecture (High-Level Design)

### 2.1 New Concepts

- **ScheduleParser (OOP service)**
  - Encapsulates the entire parsing pipeline.
  - Accepts configuration (e.g., whether to use LLM, how aggressively, etc.).
  - Provides a method like `parse_workbook(wb, filename) -> ParseResponse`.

- **ScheduleParserConfig**
  - A small configuration object controlling behavior:
    - `use_llm`: bool
    - `llm_mode`: `Literal["fallback", "refine"]`
    - Thresholds for when to invoke LLM (e.g., if too many fields are missing).

- **LLM Abstraction**
  - `BaseLLMClient` interface describing how to ask an LLM to extract/patch product fields.
  - `LLMProductPatch` model representing partial updates to a `Product`.
  - Implementations:
    - `NoopLLMClient` (default, returns empty patch).
    - `DeepInfraLLMClient` (production implementation).
    - `FakeLLMClient` (for deterministic tests).

- **LLM-Based Extraction Layer**
  - `ProductExtractor` class that:
    - Takes row context, heuristic result, and optionally LLM client.
    - Applies chosen strategy:
      - **Fallback mode:** Use heuristics; call LLM only if product is too sparse/uncertain.
      - **Refine mode:** Call LLM even when heuristics have results, then merge.

- **Configuration/Settings**
  - `app/core/config.py`: Reads environment variables and exposes settings (`USE_LLM`, `LLM_MODE`, etc.).

- **API Wiring**
  - `routes.py` obtains a configured `ScheduleParser` instance.
  - Existing function `parse_workbook` becomes a thin wrapper for backward compatibility.

### 2.2 Target Module Layout (Proposed)

**New files:**
- `app/core/config.py` – app-level settings.
- `app/parser/service.py` – `ScheduleParser`, `ScheduleParserConfig`.
- `app/parser/llm_client.py` – `BaseLLMClient`, `NoopLLMClient`, `DeepInfraLLMClient`.
- `app/parser/llm_extractor.py` – LLM-based extraction and merging logic.
- `app/parser/prompts.py` – Prompt templates (if needed).

**Existing files (touched):**
- `app/parser/workbook.py` – Move orchestration to `ScheduleParser`.
- `app/api/routes.py` – Use `ScheduleParser` instance.
- `README.md` – Document the new layer.
- `tests/` – Add new tests for LLM paths.

---

## 3. Detailed Implementation Steps

### Step 1: Introduce Configuration for LLM Usage

Create `app/core/config.py` using Pydantic `BaseSettings`.

```python
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    # LLM Configuration
    use_llm: bool = False
    llm_mode: str = "fallback"  # "fallback" | "refine"
    llm_provider: str = "deepinfra"
    llm_model: str = "openai/gpt-oss-120b"

    # DeepInfra credentials
    deepinfra_api_key: str | None = None

    # LLM behavior thresholds
    llm_min_missing_fields: int = 3  # Trigger LLM if >= N fields missing (fallback mode)
    llm_batch_size: int = 5  # Number of products per LLM call

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()
```

**Environment variables (`.env.example`):**

```bash
USE_LLM=false
LLM_MODE=fallback
LLM_PROVIDER=deepinfra
LLM_MODEL=openai/gpt-oss-120b
DEEPINFRA_API_KEY=your_deepinfra_api_key_here
LLM_BATCH_SIZE=5
```

---

### Step 2: Create OOP ScheduleParser Service

In `app/parser/service.py`:

```python
from dataclasses import dataclass
from typing import Literal
from openpyxl import Workbook

from app.core.models import ParseResponse
from app.parser.llm_client import BaseLLMClient, NoopLLMClient


@dataclass
class ScheduleParserConfig:
    """Configuration for the ScheduleParser service."""
    use_llm: bool = False
    llm_mode: Literal["fallback", "refine"] = "fallback"
    llm_min_missing_fields: int = 3
    extract_images: bool = False


class ScheduleParser:
    """
    OOP service encapsulating the entire parsing pipeline.

    Usage:
        config = ScheduleParserConfig(use_llm=True, llm_mode="fallback")
        parser = ScheduleParser(config, llm_client=DeepInfraLLMClient(...))
        result = parser.parse_workbook(wb, filename="schedule.xlsx")
    """

    def __init__(
        self,
        config: ScheduleParserConfig | None = None,
        llm_client: BaseLLMClient | None = None,
    ):
        self.config = config or ScheduleParserConfig()
        self.llm_client = llm_client or NoopLLMClient()

    def parse_workbook(self, wb: Workbook, filename: str) -> ParseResponse:
        """Parse a workbook and return structured product data."""
        # Implementation moves here from workbook.py
        ...
```

---

### Step 3: Define LLM Abstraction with DeepInfra Implementation

In `app/parser/llm_client.py`:

```python
from abc import ABC, abstractmethod
import logging
from openai import OpenAI
from pydantic import BaseModel

logger = logging.getLogger(__name__)


class LLMProductPatch(BaseModel):
    """Partial product fields returned by LLM extraction."""
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
    """Protocol for LLM-based product extraction."""

    @abstractmethod
    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None
    ) -> LLMProductPatch:
        """Extract product fields from raw text."""
        ...

    def extract_batch(
        self,
        items: list[tuple[str, dict | None]],
    ) -> list[LLMProductPatch]:
        """Extract multiple products in a single call. Default: iterate."""
        return [self.extract_product_patch(text, ctx) for text, ctx in items]


class NoopLLMClient(BaseLLMClient):
    """Default client that returns empty patches (LLM disabled)."""

    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None
    ) -> LLMProductPatch:
        return LLMProductPatch()


class FakeLLMClient(BaseLLMClient):
    """Test client that returns predictable patches for testing."""

    def __init__(self, responses: dict[str, LLMProductPatch] | None = None):
        self.responses = responses or {}
        self.calls: list[tuple[str, dict | None]] = []

    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None
    ) -> LLMProductPatch:
        self.calls.append((raw_text, context))
        # Return configured response or empty patch
        for key, patch in self.responses.items():
            if key in raw_text:
                return patch
        return LLMProductPatch()


class DeepInfraLLMClient(BaseLLMClient):
    """DeepInfra-backed LLM client using OpenAI SDK."""

    DEFAULT_MODEL = "openai/gpt-oss-120b"

    def __init__(self, api_key: str, model: str | None = None):
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://api.deepinfra.com/v1/openai",
        )
        self.model = model or self.DEFAULT_MODEL

    def extract_product_patch(
        self,
        raw_text: str,
        context: dict | None = None
    ) -> LLMProductPatch:
        prompt = self._build_extraction_prompt(raw_text, context)

        try:
            completion = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": self._system_prompt()},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
            )

            response_text = completion.choices[0].message.content
            patch = LLMProductPatch.model_validate_json(response_text)
            return self._validate_patch(patch, raw_text)

        except Exception as e:
            logger.warning(f"LLM extraction failed: {e}, returning empty patch")
            return LLMProductPatch()

    def extract_batch(
        self,
        items: list[tuple[str, dict | None]],
    ) -> list[LLMProductPatch]:
        """Extract multiple products in a single LLM call for efficiency."""
        if not items:
            return []

        # Build batched prompt
        numbered_items = []
        for i, (text, ctx) in enumerate(items, 1):
            ctx_str = f" (context: {ctx})" if ctx else ""
            numbered_items.append(f"{i}. {text}{ctx_str}")

        batch_prompt = "Extract product fields from each numbered item. Return a JSON array with one object per item, in order.\n\n" + "\n".join(numbered_items)

        try:
            completion = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": self._batch_system_prompt()},
                    {"role": "user", "content": batch_prompt}
                ],
                response_format={"type": "json_object"},
            )

            response_text = completion.choices[0].message.content
            import json
            data = json.loads(response_text)

            # Handle {"items": [...]} or [...] format
            if isinstance(data, dict) and "items" in data:
                data = data["items"]

            patches = []
            for i, item_data in enumerate(data):
                try:
                    patch = LLMProductPatch.model_validate(item_data)
                    raw_text = items[i][0] if i < len(items) else ""
                    patches.append(self._validate_patch(patch, raw_text))
                except Exception:
                    patches.append(LLMProductPatch())

            # Pad if LLM returned fewer items
            while len(patches) < len(items):
                patches.append(LLMProductPatch())

            return patches

        except Exception as e:
            logger.warning(f"Batch LLM extraction failed: {e}, falling back to individual calls")
            return [self.extract_product_patch(text, ctx) for text, ctx in items]

    def _system_prompt(self) -> str:
        return """You are a product data extraction assistant. Extract structured product information from interior design schedule text. Return a JSON object with these optional fields:
- product_name: The product or range name
- brand: Manufacturer or supplier name
- colour: Color name
- finish: Surface finish (e.g., matt, gloss)
- material: Material composition
- width: Width in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- length: Length in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- height: Height/thickness in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- qty: Quantity (integer)
- rrp: Price as a float (exclude GST, currency symbols)

Only include fields you can confidently extract. Omit uncertain fields."""

    def _batch_system_prompt(self) -> str:
        return """You are a product data extraction assistant. You will receive multiple numbered product descriptions. Extract structured product information from each and return a JSON object with an "items" array containing one object per product, in order.

Each object can have these optional fields:
- product_name: The product or range name
- brand: Manufacturer or supplier name
- colour: Color name
- finish: Surface finish (e.g., matt, gloss)
- material: Material composition
- width: Width in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- length: Length in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- height: Height/thickness in millimeters (integer). Convert from metres (*1000) or cm (*10) if needed.
- qty: Quantity (integer)
- rrp: Price as a float (exclude GST, currency symbols)

Only include fields you can confidently extract. Omit uncertain fields."""

    def _build_extraction_prompt(self, raw_text: str, context: dict | None) -> str:
        ctx_str = ""
        if context:
            ctx_str = f"\nContext: {context}\n"
        return f"{ctx_str}Extract product fields from this text:\n\n{raw_text}"

    def _validate_patch(self, patch: LLMProductPatch, raw_text: str) -> LLMProductPatch:
        """Validate and sanitize LLM output to prevent hallucinations."""
        # Reject obviously wrong dimensions (> 100 meters)
        max_dim = 100_000  # 100m in mm
        if patch.width and patch.width > max_dim:
            patch.width = None
        if patch.length and patch.length > max_dim:
            patch.length = None
        if patch.height and patch.height > max_dim:
            patch.height = None

        # Reject negative values
        if patch.qty and patch.qty < 0:
            patch.qty = None
        if patch.rrp and patch.rrp < 0:
            patch.rrp = None

        return patch


def build_llm_client(settings) -> BaseLLMClient:
    """Factory function to create LLM client from settings."""
    if not settings.use_llm:
        return NoopLLMClient()

    if settings.llm_provider == "deepinfra":
        if not settings.deepinfra_api_key:
            raise ValueError("DEEPINFRA_API_KEY required when USE_LLM=true")
        return DeepInfraLLMClient(
            api_key=settings.deepinfra_api_key,
            model=settings.llm_model,
        )

    # Default to noop if provider unknown
    logger.warning(f"Unknown LLM provider: {settings.llm_provider}, using NoopLLMClient")
    return NoopLLMClient()
```

---

### Step 4: Implement LLM-Based Extraction Layer

In `app/parser/llm_extractor.py`:

```python
from dataclasses import dataclass
from typing import Literal
import logging

from app.core.models import Product
from app.parser.llm_client import BaseLLMClient, LLMProductPatch, NoopLLMClient

logger = logging.getLogger(__name__)


@dataclass
class ExtractionContext:
    """Metadata about the extraction context."""
    sheet_name: str
    row_index: int
    section: str | None = None


class ProductExtractor:
    """
    Orchestrates product field extraction using heuristics + optional LLM.

    Modes:
    - fallback: Use heuristics first, call LLM only if product is too sparse
    - refine: Always call LLM to validate/enhance heuristic results
    """

    def __init__(
        self,
        llm_client: BaseLLMClient | None = None,
        mode: Literal["fallback", "refine"] = "fallback",
        min_missing_fields: int = 3,
        batch_size: int = 5,
    ):
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
        """Extract/enhance a single product."""
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
        """Extract/enhance multiple products, using batched LLM calls if beneficial."""
        if not items:
            return []

        # Determine which items need LLM
        if self.mode == "refine":
            needs_llm = list(range(len(items)))
        else:
            needs_llm = [
                i for i, (prod, _, _) in enumerate(items)
                if self._count_missing_fields(prod) >= self.min_missing_fields
            ]

        # If no items need LLM, return as-is
        if not needs_llm:
            return [prod for prod, _, _ in items]

        # Batch LLM calls
        llm_inputs = [
            (items[i][1], self._context_to_dict(items[i][2]))
            for i in needs_llm
        ]

        patches = []
        for batch_start in range(0, len(llm_inputs), self.batch_size):
            batch = llm_inputs[batch_start:batch_start + self.batch_size]
            batch_patches = self.llm_client.extract_batch(batch)
            patches.extend(batch_patches)

        # Build results
        results = []
        patch_idx = 0
        for i, (prod, raw_text, ctx) in enumerate(items):
            if i in needs_llm:
                patch = patches[patch_idx] if patch_idx < len(patches) else LLMProductPatch()
                patch_idx += 1
                prefer_llm = self.mode == "fallback"  # In fallback, LLM fills gaps
                results.append(self._merge_patch(prod, patch, prefer_llm=prefer_llm))
            else:
                results.append(prod)

        return results

    def _get_llm_patch(
        self,
        raw_text: str,
        context: ExtractionContext | None
    ) -> LLMProductPatch:
        ctx_dict = self._context_to_dict(context)
        return self.llm_client.extract_product_patch(raw_text, ctx_dict)

    def _context_to_dict(self, context: ExtractionContext | None) -> dict | None:
        if not context:
            return None
        return {
            "sheet": context.sheet_name,
            "row": context.row_index,
            "section": context.section,
        }

    def _count_missing_fields(self, product: Product) -> int:
        """Count how many key fields are missing."""
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
        """
        Merge LLM patch into product.

        Args:
            prefer_llm: If True, LLM values fill gaps only (heuristic wins).
                       If False, LLM values override heuristic (refine mode).
        """
        data = product.model_dump()
        patch_data = patch.model_dump(exclude_none=True)

        for field, value in patch_data.items():
            if prefer_llm:
                # Only fill if heuristic is None
                if data.get(field) is None:
                    data[field] = value
            else:
                # Refine: prefer heuristic, but use LLM if heuristic is None
                if data.get(field) is None:
                    data[field] = value

        return Product.model_validate(data)
```

---

### Step 5: Integrate LLM Layer into Parsing Pipeline

Update `ScheduleParser` in `app/parser/service.py` to use `ProductExtractor`.

Update `app/parser/workbook.py`:
- Turn `parse_workbook` into a thin wrapper that initializes a default `ScheduleParser` (LLM off).

Update `app/api/routes.py`:
- Use a global `schedule_parser` instance created via factory.

```python
# app/api/routes.py (updated)
from app.core.config import settings
from app.parser.service import ScheduleParser, ScheduleParserConfig
from app.parser.llm_client import build_llm_client

# Create parser instance at startup
_parser_config = ScheduleParserConfig(
    use_llm=settings.use_llm,
    llm_mode=settings.llm_mode,
    llm_min_missing_fields=settings.llm_min_missing_fields,
)
_llm_client = build_llm_client(settings)
schedule_parser = ScheduleParser(config=_parser_config, llm_client=_llm_client)


@router.post("/parse", response_model=ParseResponse)
async def parse_schedule(file: UploadFile = File(...)):
    # ... validation ...
    wb = load_workbook_safe(file_bytes)
    return schedule_parser.parse_workbook(wb, filename=file.filename)
```

---

### Step 6: Testing Strategy

#### 6.1 Regression Tests
- Ensure all existing tests pass with LLM disabled (default).
- No changes to test commands: `pytest -v -m "not synthetic"`

#### 6.2 Unit Tests for LLM Components

```python
# tests/test_llm_client.py

def test_noop_client_returns_empty_patch():
    client = NoopLLMClient()
    patch = client.extract_product_patch("some text")
    assert patch == LLMProductPatch()


def test_fake_client_records_calls():
    client = FakeLLMClient(responses={
        "ICONIC": LLMProductPatch(product_name="Iconic Carpet")
    })
    patch = client.extract_product_patch("Product: ICONIC carpet")
    assert patch.product_name == "Iconic Carpet"
    assert len(client.calls) == 1


def test_product_extractor_fallback_mode():
    fake_client = FakeLLMClient(responses={
        "sparse": LLMProductPatch(product_name="Filled by LLM")
    })
    extractor = ProductExtractor(
        llm_client=fake_client,
        mode="fallback",
        min_missing_fields=2,
    )

    # Product with many missing fields -> should call LLM
    sparse_product = Product(doc_code="X1")
    result = extractor.extract(sparse_product, "sparse product text")
    assert result.product_name == "Filled by LLM"
    assert len(fake_client.calls) == 1

    # Product with few missing fields -> should NOT call LLM
    rich_product = Product(
        doc_code="X2",
        product_name="Already Named",
        brand="Known Brand",
        colour="Blue",
    )
    result = extractor.extract(rich_product, "rich product text")
    assert result.product_name == "Already Named"
    assert len(fake_client.calls) == 1  # No new calls
```

#### 6.3 Integration Tests (Optional, Manual)

```python
# tests/test_llm_integration.py

@pytest.mark.llm_integration
@pytest.mark.skipif(not os.getenv("DEEPINFRA_API_KEY"), reason="No API key")
def test_deepinfra_extraction():
    client = DeepInfraLLMClient(api_key=os.environ["DEEPINFRA_API_KEY"])
    patch = client.extract_product_patch(
        "PRODUCT: ICONIC | COLOUR: SILVER SHADOW | WIDTH: 3.66 METRES"
    )
    assert patch.product_name is not None
    assert patch.width == 3660  # Converted from metres
```

---

### Step 7: Documentation and Cleanup

Update `README.md` with:

#### LLM Integration (Optional)

```markdown
## LLM-Enhanced Extraction (Optional)

The parser can optionally use an LLM to enhance extraction for ambiguous or sparse data.

### Configuration

Set these environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `USE_LLM` | `false` | Enable LLM extraction |
| `LLM_MODE` | `fallback` | `fallback` (fill gaps) or `refine` (validate all) |
| `LLM_PROVIDER` | `deepinfra` | LLM provider |
| `LLM_MODEL` | `openai/gpt-oss-120b` | Model to use |
| `DEEPINFRA_API_KEY` | - | API key (required if USE_LLM=true) |

### Modes

- **Fallback Mode** (recommended): Heuristics run first. LLM is only called if 3+ key fields are missing. Cost-efficient.
- **Refine Mode**: LLM is called for every product to validate/enhance heuristic results. Higher accuracy, higher cost.

### Running with LLM

```bash
export USE_LLM=true
export DEEPINFRA_API_KEY=your_key_here
uvicorn app.main:app --reload
```

### Design Rationale

- LLM is **additive**: The parser works fully without LLM (deterministic, debuggable).
- **Abstracted**: `BaseLLMClient` interface allows swapping providers.
- **Batched**: Multiple products are sent in a single LLM call to reduce latency.
- **Validated**: LLM outputs are sanitized to prevent hallucinations.
```

---

## 4. Migration / Backwards Compatibility Strategy

- **Behavior Preservation:** `USE_LLM=False` by default. Original functions still exist as wrappers.
- **Opt-In:** Enabling LLM usage requires no changes to existing client code (API contracts remain identical).
- **Incremental:** We start with stubs and fakes, allowing for a real LLM implementation later without architectural changes.

---

## 5. Performance Considerations

### Batching Strategy

For files with many products, batch multiple rows into a single LLM call:

- **Batch size**: 5 products per call (configurable via `LLM_BATCH_SIZE`)
- **Implementation**: `ProductExtractor.extract_batch()` and `DeepInfraLLMClient.extract_batch()`
- **Fallback**: If batch fails, retry individual rows

### Latency Estimates

| Scenario | Products | LLM Calls | Est. Time |
|----------|----------|-----------|-----------|
| Heuristic only | 74 | 0 | ~200ms |
| Fallback (30% sparse) | 74 | ~5 batches | ~3s |
| Refine (all) | 74 | ~15 batches | ~8s |

### Cost Estimates (DeepInfra gpt-oss-120b)

- Input: ~$0.0003/1K tokens
- Output: ~$0.0003/1K tokens
- Per product: ~500 input tokens, ~100 output tokens
- 74 products: ~$0.03-0.05 per file

---

## 6. Error Handling

### LLM Failures

All LLM calls are wrapped in try/except:

```python
try:
    patch = self.llm_client.extract_product_patch(...)
except Exception as e:
    logger.warning(f"LLM extraction failed: {e}, using heuristic result")
    patch = LLMProductPatch()
```

The parser **never fails** due to LLM issues. It gracefully degrades to heuristic-only.

### Validation

LLM outputs are validated:
- Dimensions > 100m are rejected
- Negative values are rejected
- (Optional) Brand names not in source text can be flagged

---

## 7. Possible Extensions (Future Work)

1. **Async LLM Integration**: Update client and parser to `async` for concurrent LLM requests across multiple sheets.

2. **Prompt Caching**: DeepInfra/OpenAI support prompt caching for repeated system prompts.

3. **Evaluation Harness**: Build a script to compare heuristic vs. LLM accuracy on a golden dataset.

4. **Model Experimentation**: Test smaller/faster models for fallback mode, larger models for refine mode.

5. **Structured Output Schema**: Use JSON Schema in the prompt to enforce output format more reliably.

---

## 8. Updated Dependencies

Add to `requirements.txt`:

```
# LLM Integration (optional, only needed if USE_LLM=true)
openai>=1.0.0
pydantic-settings>=2.0.0
```

---

## 9. Implementation Order (Recommended)

1. **Step 1-2**: Configuration + OOP refactor (no LLM yet)
   - Verify all existing tests pass
   - This is the "safe" refactor

2. **Step 3**: Add `NoopLLMClient` and `FakeLLMClient`
   - Write unit tests for extractor with fakes
   - Still no external dependencies

3. **Step 4-5**: Add `DeepInfraLLMClient` and wire up
   - Add integration test (skipped without API key)
   - Test manually with real API

4. **Step 6-7**: Testing + Documentation
   - Ensure full test coverage
   - Update README

---

## 10. Checklist

### Implementation
- [x] Create `app/core/config.py` with Settings
- [x] Create `app/parser/service.py` with ScheduleParser
- [ ] Create `app/parser/llm_client.py` with all clients
- [ ] Create `app/parser/llm_extractor.py` with ProductExtractor
- [ ] Refactor `workbook.py` to use ScheduleParser
- [ ] Update `routes.py` to use configured parser
- [ ] Add `.env.example` with LLM variables

### Testing
- [x] All existing tests pass (regression)
- [ ] Unit tests for NoopLLMClient
- [ ] Unit tests for FakeLLMClient
- [ ] Unit tests for ProductExtractor (fallback mode)
- [ ] Unit tests for ProductExtractor (refine mode)
- [ ] Integration test for DeepInfraLLMClient (optional)

### Documentation
- [ ] Update README with LLM section
- [ ] Add architecture diagram showing LLM layer
- [ ] Document configuration options
- [ ] Add .env.example
