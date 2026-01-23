# Interior Schedule Parser API

A Python REST API that parses interior designer Excel schedules (.xlsx) into structured JSON. Built with FastAPI, this service handles the variability of real-world schedule formats including inconsistent column names, merged cells, multiple sheets, and diverse layouts, accurately extracting and standardizing product data for seamless import into product management platforms.

## Table of Contents

- [Quick Start](#quick-start)
- [API Documentation](#api-documentation)
- [Architecture](#architecture)
- [Design Decisions](#design-decisions)
- [Known Limitations](#known-limitations)
- [Testing](#testing)
- [Project Structure](#project-structure)

## Quick Start

### Prerequisites

- Python 3.11+
- pip

### Local Setup

```bash
# Clone and navigate to project
git clone <repository-url>
cd design-specs-parser

# Create virtual environment and install dependencies
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Run the server
uvicorn app.main:app --reload
```

The API will be available at `http://localhost:8000`.

### Docker

```bash
# Build the image
docker build -t programa-parser .

# Run the container
docker run -p 8000:8000 programa-parser
```

## API Documentation

Interactive API documentation is available at:

- **Swagger UI**: http://localhost:8000/docs
- **ReDoc**: http://localhost:8000/redoc

### Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/parse` | Parse an Excel schedule into JSON |
| GET | `/health` | Health check |

### Example Request/Response

**Request:**

```bash
curl -X POST -F "file=@data/schedule_sample1.xlsx" http://localhost:8000/parse
```

**Response:**

```json
{
  "schedule_name": "12006: GEM, WATERLINE PLACE, WILLIAMSTOWN",
  "products": [
    {
      "doc_code": "FCA-01 A",
      "product_name": "ICONIC",
      "brand": "VICTORIA CARPETS",
      "colour": "SILVER SHADOW",
      "finish": null,
      "material": "80% WOOL 20% SYNTHETIC",
      "width": 3660,
      "length": null,
      "height": null,
      "qty": null,
      "rrp": 45.5,
      "feature_image": null,
      "product_description": "FLOORING | CARPET\n\nAPARTMENTS\nGOLD SCHEME",
      "product_details": "CODE: 50/2833 | STYLE: TWIST | CARPET_THICKNESS: 11 MM | ..."
    }
  ]
}
```

### Output Fields

| Field | Type | Description |
|-------|------|-------------|
| `doc_code` | string | Drawing or reference code |
| `product_name` | string | Product display name |
| `brand` | string | Manufacturer/brand name |
| `colour` | string | Colour name |
| `finish` | string | Surface finish |
| `material` | string | Main material |
| `width` | integer | Width in mm |
| `length` | integer | Length in mm |
| `height` | integer | Height/depth in mm |
| `qty` | integer | Quantity |
| `rrp` | decimal | Recommended retail price |
| `feature_image` | string | Image filename (optional) |
| `product_description` | string | Short description |
| `product_details` | string | Additional specifications |

## Architecture

### Parsing Pipeline

```mermaid
flowchart LR
    subgraph Input
        A[Excel Upload]
    end

    subgraph Validation
        B[File Validator]
    end

    subgraph Parser
        C[Workbook Loader]
        D[Sheet Detector]
        E[Merged Cell Handler]
        F[Column Mapper]
        G[Row Extractor]
        H[Field Parser]
        I[Normalizers]
    end

    subgraph Output
        J[JSON Response]
    end

    A --> B
    B --> C
    C --> D
    D --> E
    E --> F
    F --> G
    G --> H
    H --> I
    I --> J
```

### Component Responsibilities

| Component | Purpose |
|-----------|---------|
| **Workbook Loader** | Load .xlsx files safely, extract schedule name |
| **Sheet Detector** | Identify schedule sheets by scoring header rows |
| **Merged Cell Handler** | Fill merged cell regions for consistent reading |
| **Column Mapper** | Map varied header names to canonical columns |
| **Row Extractor** | Iterate product rows, handle grouped layouts |
| **Field Parser** | Extract key-value pairs from specification text |
| **Normalizers** | Convert dimensions to mm, parse prices |

### Proposed Production Deployment

```mermaid
flowchart TB
    subgraph Clients
        Web[Web Application]
        Mobile[Mobile App]
    end

    subgraph "Load Balancer"
        LB[Application Load Balancer]
    end

    subgraph "Container Service"
        API1[Parser API - Instance 1]
        API2[Parser API - Instance 2]
        API3[Parser API - Instance N]
    end

    subgraph Storage
        S3[Object Storage]
    end

    subgraph "Async Processing"
        Queue[Message Queue]
        Worker[Background Workers]
    end

    subgraph Monitoring
        Logs[Log Aggregation]
        Metrics[Metrics & Alerts]
    end

    Web --> LB
    Mobile --> LB
    LB --> API1
    LB --> API2
    LB --> API3
    API1 --> S3
    API2 --> S3
    API3 --> S3
    API1 --> Queue
    API2 --> Queue
    API3 --> Queue
    Queue --> Worker
    Worker --> S3
    API1 --> Logs
    API2 --> Logs
    API3 --> Logs
    API1 --> Metrics
    API2 --> Metrics
    API3 --> Metrics
```

**Key production considerations:**

- **Horizontal scaling**: Stateless API instances behind a load balancer
- **Async processing**: Queue large files for background processing
- **Object storage**: Store uploaded files and extracted images
- **Health checks**: Docker HEALTHCHECK enables orchestrator monitoring
- **Non-root user**: Container runs as unprivileged user for security

## Design Decisions

### 1. openpyxl over pandas

Chosen for fine-grained control over merged cells, formula handling, and cell-level operations. Pandas abstracts away details needed to handle real-world schedule complexity.

### 2. Rules + Heuristics Parsing (No ML/LLM)

The parser uses deterministic rules and pattern matching rather than ML models:
- Explainable and debuggable behavior
- No external API dependencies
- Consistent, reproducible results
- Lower operational complexity

### 3. All Fields Optional

Every output field defaults to `null`. This enables graceful degradation—missing data in the input produces partial output rather than validation errors.

### 4. Fuzzy Header Matching

Headers are matched against synonym lists using case-insensitive comparison and fuzzy matching. This handles variations like `QTY`, `Quantity`, `QUANTITY`, `Qty.` mapping to the same canonical column.

### 5. De-duplication by doc_code

When the same `doc_code` appears across multiple sheets (common in schedules with both "Schedule" and "Sales Schedule" sheets), the first occurrence is kept. Products with no `doc_code` are all retained.

### 6. Dimension Normalization to mm

All dimensions are converted to millimeters:
- `3.66 METRES` → `3660`
- `60 CM` → `600`
- `600 W X 600 H MM` → `width: 600, height: 600`

### 7. Multi-Layout Support

The parser handles three layout patterns:
- **Single-row-per-product**: Each row is a complete product (samples 1, 2)
- **Grouped rows**: Product header followed by detail rows with `Key: Value` format (sample 3)
- **Section headers**: Merged cells indicating category context (sample 1)

## Known Limitations

1. **No image extraction**: Embedded images are not currently extracted. The `feature_image` field is reserved for future implementation.

2. **External workbook references**: Formula references like `='[1]Cover Sheet'!A6` cannot be resolved by openpyxl. The parser falls back to reading the target cell directly or using the filename.

3. **Ambiguous dimensions**: When a specification contains only `SIZE: 5500 X 2800 MM` without W/L/H indicators, the parser assigns width and length but cannot determine orientation.

4. **Price context ignored**: Prices extracted as numeric values only. Context like "PER SQM" or "PER LINEAR METRE" is not preserved.

5. **No password-protected files**: Encrypted or password-protected workbooks will fail to parse.

6. **Sheet name matching**: Sheets are detected by header content, not name. A sheet named "Notes" with a valid product table header will be parsed.

## Testing

```bash
# Activate virtual environment
source .venv/bin/activate

# Run fast tests (excludes synthetic data tests)
pytest -v -m "not synthetic"

# Run synthetic robustness tests
pytest -v -m synthetic

# Run all tests
pytest -v

# Run with coverage
pytest --cov=app --cov-report=term-missing
```

### Test Categories

- **Unit tests**: Field parsing, dimension normalization, price extraction
- **Integration tests**: Full API endpoint testing with sample files
- **Synthetic tests**: Robustness testing with generated and mutated schedules

### Generating Synthetic Test Data

```bash
python tools/generate_programa_test_schedules.py \
  --mode both \
  --samples_dir ./data \
  --output_dir ./synthetic_out \
  --num_generated 20 \
  --seed 12345
```

## Project Structure

```
.
├── app/
│   ├── main.py              # FastAPI application entry point
│   ├── api/
│   │   └── routes.py        # API endpoint definitions
│   ├── core/
│   │   └── models.py        # Pydantic models (Product, ParseResponse)
│   └── parser/
│       ├── workbook.py      # Workbook loading, schedule name extraction
│       ├── sheet_detector.py # Header detection, schedule sheet identification
│       ├── merged_cells.py  # Merged cell region handling
│       ├── column_mapper.py # Header-to-column mapping
│       ├── row_extractor.py # Product row iteration
│       ├── field_parser.py  # Key-value text parsing
│       └── normalizers.py   # Dimension and price normalization
├── data/                    # Sample schedule files
├── tests/                   # Test suite
├── tools/                   # Synthetic data generator
├── Dockerfile
├── requirements.txt
└── README.md
```

## License

This project was created as part of a technical assessment.
