# CLAUDE.md

## Project Overview

FastAPI REST API that parses interior designer Excel schedules (.xlsx) into structured JSON. Take-home challenge for Programa MLE position.

**Core endpoint**: `POST /parse` — accepts .xlsx upload, returns JSON with `schedule_name` and `products[]`

## Quick Reference

```bash
# Development
uvicorn app.main:app --reload

# Testing
pytest -v -m "not synthetic"      # Fast tests
pytest -v -m synthetic            # Robustness tests
pytest -v                         # All tests

# Generate test data
python tools/generate_programa_test_schedules.py --mode both --samples_dir ./data --output_dir ./synthetic_out --num_generated 20 --seed 12345

# Test endpoint
curl -X POST -F "file=@data/schedule_sample1.xlsx" http://localhost:8000/parse | jq
curl -X POST -F "file=@data/schedule_sample3.xlsx" http://localhost:8000/parse | jq

# Docker
docker build -t programa-parser . && docker run -p 8000:8000 programa-parser
```

## Key Architecture Decisions

1. **openpyxl over pandas** — Need fine-grained control over merged cells, images, formulas
2. **All fields Optional** — Graceful degradation; missing data → `null`, never crash
3. **De-dup by doc_code only** — Simple strategy, documented in README
4. **No external LLM calls** — Pure rules + heuristics for parsing (deterministic, debuggable)

## File Structure

```
app/
├── main.py              # FastAPI app entry
├── api/routes.py        # /parse, /health endpoints
├── core/models.py       # Pydantic: Product, ParseResponse, ErrorResponse
└── parser/
    ├── workbook.py      # Load workbook, get schedule_name
    ├── sheet_detector.py # Find headers, detect schedule sheets
    ├── column_mapper.py  # Map headers → canonical columns
    ├── row_extractor.py  # Iterate products (row-per-product + grouped detail rows)
    ├── field_parser.py   # Parse KEY: VALUE from spec text
    └── normalizers.py    # Dimensions (→mm), qty, prices
tools/
└── generate_programa_test_schedules.py  # ✅ Synthetic test generator
data/
├── schedule_sample1.xlsx
├── schedule_sample2.xlsx
└── schedule_sample3.xlsx
```

## Known Gotchas

- **Formula references**: Sample2 has `='[1]Cover Sheet'!A6` (external workbook ref) — openpyxl can't resolve. Fallback to reading Cover Sheet!A6 directly.
- **doc_code variability**: `doc_code` can be short alphanumeric (e.g., `L1`, `F64`) or complex (e.g., `FCA-01 A`, `PTF-*K`). Treat it as an opaque string; don’t use a strict regex as a gate for product rows.
- **KV separators**: Specs use both `:` and `-` as delimiters. Some have no space: `FINISH- MATT`
- **Merged cells**: openpyxl returns `None` for non-top-left cells in merged ranges. Must fill before reading.
- **Sheet name trailing space**: Sample2 has `"Sales Schedule "` (with space). Use `.strip()` for comparisons.
- **Grouped product rows**: Sample3 uses a product “item row” followed by detail rows (`Maker:`, `Name:`, `Finish:`, `Size:`, `Notes:`) that must be attached to the preceding item row.

## Code Style

- Python 3.11+ (use `str | None` not `Optional[str]`)
- Type hints on all functions
- Docstrings for public functions
- Keep parsing functions pure where possible (easier to test)
- Wrap risky operations in try/except, return partial results over crashing

## Skills & Hooks

Check `.claude/` for available skills and hooks before starting tasks:

```
.claude/
├── skills/          # Reusable capabilities (if defined)
├── hooks/           # Pre/post command hooks (if defined)
└── settings.json    # Project-specific settings
```

**Use skills** for repetitive patterns (e.g., creating new parser modules, writing tests).
**Use hooks** for automated checks (e.g., lint on save, test on commit).

## Current State

- [x] Test data generator implemented (`tools/`)
- [x] Phase 1: API skeleton + models
- [x] Phase 2: Sheet detection + merged cells
- [x] Phase 3: Row extraction + field parsing
- [x] Phase 4: Orchestration + de-dup
- [x] Phase 5: Testing
- [ ] Phase 6: Polish (Docker, README, diagram)

## Sample Data Quick Facts

| File | Sheets | Header Row | Products | Has Prices |
|------|--------|------------|----------|------------|
| sample1.xlsx | APARTMENTS | 4 | ~74 | Yes |
| sample2.xlsx | Cover Sheet, Schedule, Sales Schedule | 9 | ~53 | No |
| sample3.xlsx | Schedule | 10 | Varies (row-grouped items) | Yes |

## Don't Forget

- Run `pytest` before committing
- Update this file's "Current State" as phases complete
- Check `/docs` endpoint works after API changes
- Include a Mermaid architecture diagram in `README.md` for submission
