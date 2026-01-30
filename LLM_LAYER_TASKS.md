# LLM Layer Implementation Tasks

> **Reference:** See `LLM_LAYER_PLAN.md` for detailed design rationale and code samples.
> **Branch:** `add-llm-layer`

## Status Legend
- `[ ]` Pending
- `[x]` Complete

---

## Phase 1: Configuration & OOP Foundation

### Task 1.1: Create Settings Configuration
- [x] **Status: Complete**

**What:** Create `app/core/config.py` with Pydantic Settings for app configuration.

**Why:** Centralize configuration (especially LLM toggles) in one place. Currently config is scattered/hardcoded.

**Read:** `LLM_LAYER_PLAN.md` (Step 1 for Settings class structure)

**Create:** `app/core/config.py`

**Details:**
- Use `pydantic_settings.BaseSettings`
- Fields: `use_llm` (bool, default False), `llm_mode` (str, default "fallback"), `llm_provider`, `llm_model`, `deepinfra_api_key`, `llm_min_missing_fields`, `llm_batch_size`
- Load from `.env` file
- Export singleton `settings` instance

**Done when:** `from app.core.config import settings` works and reads from env vars.

---

### Task 1.2: Create ScheduleParserConfig Dataclass
- [x] **Status: Complete**

**What:** Create config dataclass for parser behavior in `app/parser/service.py`.

**Why:** Separate runtime parser config from app-level settings. Allows different parser configs per request if needed.

**Read:** `LLM_LAYER_PLAN.md` (Step 2 for ScheduleParserConfig)

**Create:** `app/parser/service.py` (new file)

**Details:**
- `@dataclass` with fields: `use_llm`, `llm_mode`, `llm_min_missing_fields`, `extract_images`
- All fields have sensible defaults (LLM off by default)

**Done when:** `ScheduleParserConfig()` instantiates with defaults.

---

### Task 1.3: Create ScheduleParser Class Skeleton
- [x] **Status: Complete**

**What:** Add `ScheduleParser` class to `app/parser/service.py` with method stubs.

**Why:** OOP service encapsulating parsing pipeline. Reviewers requested OOP structure.

**Read:**
- `app/parser/workbook.py` (understand current `parse_workbook` function)
- `LLM_LAYER_PLAN.md` (Step 2 for class structure)

**Modify:** `app/parser/service.py`

**Details:**
- Constructor takes `config: ScheduleParserConfig` and `llm_client: BaseLLMClient | None`
- Method `parse_workbook(wb: Workbook, filename: str) -> ParseResponse` — stub that raises NotImplementedError for now
- Import necessary types from `app.core.models`

**Done when:** Class instantiates without errors. Method exists but not implemented.

---

### Task 1.4: Move Parsing Logic to ScheduleParser
- [x] **Status: Complete**

**What:** Move logic from `parse_workbook()` function in `workbook.py` into `ScheduleParser.parse_workbook()` method.

**Why:** Encapsulate orchestration in the service class instead of module-level function.

**Read:**
- `app/parser/workbook.py` — the `parse_workbook` function (main logic to move)
- `app/parser/sheet_detector.py`, `column_mapper.py`, `row_extractor.py`, `field_parser.py` — understand dependencies

**Modify:**
- `app/parser/service.py` — implement `parse_workbook` method
- `app/parser/workbook.py` — keep `load_workbook_safe`, `get_schedule_name`, `fill_merged_regions` as utilities

**Details:**
- Copy the iteration logic (sheets → headers → rows → products)
- Keep using existing helper functions from other parser modules
- Do NOT add LLM calls yet — just move the heuristic logic
- Keep `_dedupe_products_by_doc_code` and `_is_empty_product` (can be private methods or stay in workbook.py)

**Done when:** `ScheduleParser(...).parse_workbook(wb, filename)` returns same results as old function.

---

### Task 1.5: Create Backward-Compatible Wrapper
- [x] **Status: Complete**

**What:** Make old `parse_workbook()` function a thin wrapper around `ScheduleParser`.

**Why:** Preserve backward compatibility. Existing code/tests can still call `parse_workbook()`.

**Modify:** `app/parser/workbook.py`

**Details:**
- `parse_workbook(wb, filename, extract_images=False)` creates a default `ScheduleParser` and calls its method
- Import `ScheduleParser`, `ScheduleParserConfig` from service module

**Done when:** Existing tests pass without modification.

---

### Task 1.6: Wire Routes to Use ScheduleParser
- [x] **Status: Complete**

**What:** Update `/parse` endpoint to use a configured `ScheduleParser` instance.

**Why:** API should use the new OOP service with config from settings.

**Read:**
- `app/api/routes.py`
- `app/core/config.py` (settings)

**Modify:** `app/api/routes.py`

**Details:**
- Import `settings` from config
- Import `ScheduleParser`, `ScheduleParserConfig` from service
- Create parser instance at module level (or use dependency injection)
- For now, `use_llm=settings.use_llm` but LLM client will be NoopLLMClient (added in Phase 2)

**Done when:** `/parse` endpoint works, using ScheduleParser internally. All API tests pass.

---

### Task 1.7: Verify Regression Tests Pass
- [x] **Status: Complete**

**What:** Run full test suite to ensure Phase 1 refactor didn't break anything.

**Why:** Safety checkpoint before adding LLM layer.

**Run:** `pytest -v -m "not synthetic"` and `pytest -v`

**Done when:** All existing tests pass. No behavior changes.

---

## Phase 2: LLM Client Abstraction

### Task 2.1: Create LLMProductPatch Model
- [x] **Status: Complete**

**What:** Create Pydantic model for partial product updates from LLM.

**Why:** LLM returns partial data that gets merged into heuristic results. Need a typed structure.

**Read:** `LLM_LAYER_PLAN.md` (Step 3 for model definition)

**Create:** `app/parser/llm_client.py` (new file)

**Details:**
- `LLMProductPatch(BaseModel)` with all Optional fields: `product_name`, `brand`, `colour`, `finish`, `material`, `width`, `length`, `height`, `qty`, `rrp`
- All fields default to `None`

**Done when:** `LLMProductPatch()` creates empty patch, `LLMProductPatch(brand="Test")` works.

---

### Task 2.2: Create BaseLLMClient ABC
- [x] **Status: Complete**

**What:** Define abstract base class for LLM clients.

**Why:** Allows swapping implementations (Noop, Fake, DeepInfra) without changing calling code.

**Modify:** `app/parser/llm_client.py`

**Details:**
- `BaseLLMClient(ABC)` with:
  - `@abstractmethod extract_product_patch(raw_text: str, context: dict | None) -> LLMProductPatch`
  - `extract_batch(items: list[tuple[str, dict | None]]) -> list[LLMProductPatch]` — default impl iterates

**Done when:** Class defined, cannot be instantiated directly.

---

### Task 2.3: Implement NoopLLMClient
- [ ] **Status: Pending**

**What:** Default client that returns empty patches (LLM disabled).

**Why:** Used when `USE_LLM=false`. Parser works without any LLM calls.

**Modify:** `app/parser/llm_client.py`

**Details:**
- `NoopLLMClient(BaseLLMClient)`
- `extract_product_patch()` returns `LLMProductPatch()` (all None)

**Done when:** `NoopLLMClient().extract_product_patch("any text")` returns empty patch.

---

### Task 2.4: Implement FakeLLMClient
- [ ] **Status: Pending**

**What:** Test client that returns configurable responses and records calls.

**Why:** Enables deterministic unit tests without real LLM calls.

**Modify:** `app/parser/llm_client.py`

**Details:**
- `FakeLLMClient(BaseLLMClient)`
- Constructor takes `responses: dict[str, LLMProductPatch]` — maps substring to response
- Tracks `calls: list[tuple[str, dict | None]]` for assertions
- `extract_product_patch()` checks if any key is substring of input, returns that patch

**Done when:** Can configure fake responses and verify calls were made.

---

### Task 2.5: Create build_llm_client Factory
- [ ] **Status: Pending**

**What:** Factory function to create appropriate LLM client from settings.

**Why:** Centralize client creation logic. Routes/service just call factory.

**Modify:** `app/parser/llm_client.py`

**Details:**
- `build_llm_client(settings) -> BaseLLMClient`
- If `not settings.use_llm`: return `NoopLLMClient()`
- If `settings.llm_provider == "deepinfra"`: placeholder (raise or return Noop until Task 4.1)
- Log warning for unknown provider

**Done when:** Factory returns NoopLLMClient when LLM disabled.

---

### Task 2.6: Write Unit Tests for LLM Clients
- [ ] **Status: Pending**

**What:** Test NoopLLMClient and FakeLLMClient behavior.

**Why:** Ensure abstractions work correctly before building on them.

**Create:** `tests/test_llm_client.py`

**Details:**
- Test `NoopLLMClient` returns empty patch
- Test `FakeLLMClient` returns configured responses
- Test `FakeLLMClient` records calls
- Test `FakeLLMClient` returns empty for unmatched input

**Done when:** `pytest tests/test_llm_client.py -v` passes.

---

## Phase 3: Product Extractor Layer

### Task 3.1: Create ExtractionContext Dataclass
- [ ] **Status: Pending**

**What:** Metadata structure passed to LLM for context.

**Why:** LLM can use sheet name, row index, section info to improve extraction.

**Create:** `app/parser/llm_extractor.py` (new file)

**Details:**
- `@dataclass ExtractionContext` with: `sheet_name: str`, `row_index: int`, `section: str | None = None`

**Done when:** Dataclass instantiates correctly.

---

### Task 3.2: Implement ProductExtractor (Fallback Mode)
- [ ] **Status: Pending**

**What:** Class that orchestrates heuristic + LLM extraction with fallback strategy.

**Why:** Core logic for when to call LLM and how to merge results.

**Read:** `LLM_LAYER_PLAN.md` (Step 4 for ProductExtractor)

**Modify:** `app/parser/llm_extractor.py`

**Details:**
- `ProductExtractor` with constructor: `llm_client`, `mode`, `min_missing_fields`, `batch_size`
- `extract(heuristic_product: Product, raw_text: str, context) -> Product`
- `_count_missing_fields(product)` — count None in key fields
- `_merge_patch(product, patch, prefer_llm)` — merge LLM patch into product
- Fallback mode: only call LLM if missing_count >= threshold

**Done when:** Extractor calls LLM only for sparse products, merges correctly.

---

### Task 3.3: Add Batch Extraction to ProductExtractor
- [ ] **Status: Pending**

**What:** Method to process multiple products efficiently.

**Why:** Batch LLM calls reduce latency for large files.

**Modify:** `app/parser/llm_extractor.py`

**Details:**
- `extract_batch(items: list[tuple[Product, str, ExtractionContext | None]]) -> list[Product]`
- Identify which items need LLM (based on mode and missing fields)
- Call `llm_client.extract_batch()` for efficiency
- Merge patches back into products

**Done when:** Batch extraction works, batches LLM calls appropriately.

---

### Task 3.4: Write Unit Tests for ProductExtractor
- [ ] **Status: Pending**

**What:** Test extraction logic using FakeLLMClient.

**Why:** Verify fallback/refine modes work correctly without real LLM.

**Create:** `tests/test_llm_extractor.py`

**Details:**
- Test fallback mode skips LLM for rich products
- Test fallback mode calls LLM for sparse products
- Test merge prefers heuristic values (fills gaps only)
- Test batch extraction

**Done when:** `pytest tests/test_llm_extractor.py -v` passes.

---

### Task 3.5: Integrate Extractor into ScheduleParser
- [ ] **Status: Pending**

**What:** Wire ProductExtractor into the parsing pipeline.

**Why:** Complete the integration — parser can now use LLM when enabled.

**Modify:** `app/parser/service.py`

**Details:**
- Create `ProductExtractor` in `ScheduleParser.__init__` using config and llm_client
- In `parse_workbook`, after heuristic extraction, call extractor
- Build raw_text for LLM from row data (concatenate relevant cell values)
- Use batch extraction for efficiency

**Done when:** Parser uses extractor. With NoopLLMClient, behavior unchanged. With FakeLLMClient, can verify LLM is called.

---

### Task 3.6: Verify Regression Tests Still Pass
- [ ] **Status: Pending**

**What:** Run tests to ensure LLM integration (with Noop) doesn't break anything.

**Run:** `pytest -v`

**Done when:** All tests pass. Default behavior (LLM off) unchanged.

---

## Phase 4: DeepInfra Implementation

### Task 4.1: Implement DeepInfraLLMClient (Single Extraction)
- [ ] **Status: Pending**

**What:** Real LLM client using DeepInfra's OpenAI-compatible API.

**Why:** Production implementation for LLM extraction.

**Read:** `LLM_LAYER_PLAN.md` (Step 3 for DeepInfraLLMClient code)

**Modify:** `app/parser/llm_client.py`

**Add dependency:** `openai>=1.0.0` to `requirements.txt`

**Details:**
- `DeepInfraLLMClient(BaseLLMClient)` with `api_key`, `model` params
- Use `OpenAI` client with `base_url="https://api.deepinfra.com/v1/openai"`
- Implement `extract_product_patch()` with JSON response format
- Add `_system_prompt()` and `_build_extraction_prompt()` helpers

**Done when:** Client can make real API calls (test manually with API key).

---

### Task 4.2: Add Output Validation to DeepInfraLLMClient
- [ ] **Status: Pending**

**What:** Sanitize LLM outputs to prevent hallucinations.

**Why:** LLM might return invalid dimensions, negative values, etc.

**Modify:** `app/parser/llm_client.py`

**Details:**
- `_validate_patch(patch, raw_text) -> LLMProductPatch`
- Reject dimensions > 100m (100,000 mm)
- Reject negative qty/rrp
- Called after parsing LLM response

**Done when:** Invalid values are set to None instead of passed through.

---

### Task 4.3: Add Batch Extraction to DeepInfraLLMClient
- [ ] **Status: Pending**

**What:** Override `extract_batch()` for efficient multi-product calls.

**Why:** Single LLM call for multiple products reduces latency significantly.

**Modify:** `app/parser/llm_client.py`

**Details:**
- Build numbered prompt with all items
- `_batch_system_prompt()` instructs LLM to return JSON array
- Parse response as `{"items": [...]}` or `[...]`
- Validate each patch
- Fallback to individual calls if batch fails

**Done when:** Batch extraction works, handles edge cases gracefully.

---

### Task 4.4: Update build_llm_client Factory
- [ ] **Status: Pending**

**What:** Wire up DeepInfraLLMClient in factory.

**Modify:** `app/parser/llm_client.py`

**Details:**
- If `settings.llm_provider == "deepinfra"` and API key present: return `DeepInfraLLMClient`
- Raise `ValueError` if API key missing when LLM enabled

**Done when:** `build_llm_client(settings)` returns correct client based on config.

---

### Task 4.5: Write Integration Test (Skippable)
- [ ] **Status: Pending**

**What:** Test real DeepInfra API calls (skipped without API key).

**Create:** `tests/test_llm_integration.py`

**Details:**
- Mark with `@pytest.mark.llm_integration`
- Skip if `DEEPINFRA_API_KEY` not set
- Test single extraction returns sensible values
- Test unit conversion (metres → mm)

**Done when:** Test passes when API key provided, skips gracefully otherwise.

---

## Phase 5: Documentation & Finalization

### Task 5.1: Create .env.example
- [ ] **Status: Pending**

**What:** Example environment file documenting all config options.

**Create:** `.env.example`

**Details:**
```
USE_LLM=false
LLM_MODE=fallback
LLM_PROVIDER=deepinfra
LLM_MODEL=openai/gpt-oss-120b
DEEPINFRA_API_KEY=your_key_here
LLM_MIN_MISSING_FIELDS=3
LLM_BATCH_SIZE=5
```

**Done when:** File exists with all documented options.

---

### Task 5.2: Update README with LLM Section
- [ ] **Status: Pending**

**What:** Document LLM integration for users and reviewers.

**Modify:** `README.md`

**Details:**
- Add "LLM-Enhanced Extraction (Optional)" section
- Document env vars in table format
- Explain fallback vs refine modes
- Show how to run with LLM enabled
- Note that LLM is additive (parser works without it)

**Done when:** README clearly explains LLM feature.

---

### Task 5.3: Update Architecture Diagram
- [ ] **Status: Pending**

**What:** Add LLM layer to existing Mermaid diagram.

**Modify:** `README.md` (or separate diagram file)

**Details:**
- Show ScheduleParser as central service
- Show optional LLM client path
- Show fallback decision point

**Done when:** Diagram reflects new architecture.

---

### Task 5.4: Update CLAUDE.md
- [ ] **Status: Pending**

**What:** Update project instructions with new architecture info.

**Modify:** `CLAUDE.md`

**Details:**
- Update file structure section with new files
- Add LLM config to quick reference
- Update "Current State" section

**Done when:** CLAUDE.md reflects new architecture.

---

### Task 5.5: Final Regression Testing
- [ ] **Status: Pending**

**What:** Full test suite with both LLM off and LLM on (with fake).

**Run:**
- `pytest -v` (LLM off)
- `pytest -v -m llm_integration` (if API key available)

**Done when:** All tests pass. Ready for merge.

---

## Progress Summary

| Phase | Tasks | Complete |
|-------|-------|----------|
| 1. Foundation | 7 | 7 |
| 2. LLM Abstraction | 6 | 2 |
| 3. Extractor | 6 | 0 |
| 4. DeepInfra | 5 | 0 |
| 5. Documentation | 5 | 0 |
| **Total** | **29** | **9** |

---

## Quick Reference: File Map

| New File | Created In |
|----------|------------|
| `app/core/config.py` | Task 1.1 |
| `app/parser/service.py` | Task 1.2 |
| `app/parser/llm_client.py` | Task 2.1 |
| `app/parser/llm_extractor.py` | Task 3.1 |
| `tests/test_llm_client.py` | Task 2.6 |
| `tests/test_llm_extractor.py` | Task 3.4 |
| `tests/test_llm_integration.py` | Task 4.5 |
| `.env.example` | Task 5.1 |

| Modified File | Tasks |
|---------------|-------|
| `app/parser/workbook.py` | 1.4, 1.5 |
| `app/api/routes.py` | 1.6 |
| `requirements.txt` | 4.1 |
| `README.md` | 5.2, 5.3 |
| `CLAUDE.md` | 5.4 |
