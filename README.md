# MDL DOCX Builder

A FastAPI-based service for generating Management Decision Letters (MDL) as DOCX documents for U.S. Department of the Treasury audit findings. The service integrates with the Federal Audit Clearinghouse (FAC) API to automatically fetch audit data and generate professionally formatted documents.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
- [API Reference](#api-reference)
- [Data Models](#data-models)
- [Services](#services)
- [Testing](#testing)
- [Development](#development)

## Overview

The MDL DOCX Builder automates the creation of Management Decision Letters for Treasury programs including:

- **SLFRF** (21.027) - Coronavirus State and Local Fiscal Recovery Funds
- **ERA** (21.023) - Emergency Rental Assistance Program
- **HAF** (21.026) - Homeowner Assistance Fund
- **CPF** (21.029) - Capital Projects Fund
- **SSBCI** (21.031) - State Small Business Credit Initiative
- **LATCF** (21.032) - Local Assistance and Tribal Consistency Fund

The service fetches audit findings from the FAC API, processes the data, applies compliance type mappings, and generates formatted DOCX documents using customizable templates.

## Features

- **Automatic FAC Integration**: Fetches audit data directly from the Federal Audit Clearinghouse API
- **Template-Based Generation**: Uses DOCX templates with placeholder replacement
- **Smart Finding Categorization**: Maps compliance types to standardized categories using AI assistance
- **Grammar Correction**: Automatically fixes singular/plural agreement in generated text
- **Multiple Output Options**: Supports Azure Blob Storage or local file storage
- **Flexible API Endpoints**: Multiple endpoints for different use cases (manual, FAC-based, fully automatic)
- **Treasury Program Filtering**: Filter findings by specific Treasury assistance listings

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Client/API    │────▶│   FastAPI App   │────▶│   FAC API       │
│   Consumer      │     │   (main.py)     │     │   (External)    │
└─────────────────┘     └────────┬────────┘     └─────────────────┘
                                 │
                    ┌────────────┼────────────┐
                    │            │            │
              ┌─────▼─────┐ ┌────▼────┐ ┌─────▼─────┐
              │  Services │ │  Models │ │   Utils   │
              │           │ │         │ │           │
              │ • MDL     │ │ Pydantic│ │ • Text    │
              │   Builder │ │ Schemas │ │ • DOCX    │
              │ • Template│ │         │ │           │
              │ • Storage │ │         │ │           │
              │ • FAC API │ │         │ │           │
              └─────┬─────┘ └─────────┘ └───────────┘
                    │
              ┌─────▼─────┐
              │  Storage  │
              │           │
              │ • Azure   │
              │   Blob    │
              │ • Local   │
              │   Files   │
              └───────────┘
```

## Project Structure

```
mdl-backend-demo/
├── main.py                     # FastAPI application entry point
├── config.py                   # Environment variables and constants
├── mdl_helpers.py              # Finding summaries and compliance type mappings
├── pytest.ini                  # Pytest configuration
├── requirements-test.txt       # Test dependencies
│
├── models/
│   ├── __init__.py
│   └── schemas.py              # Pydantic request/response models
│
├── routes/
│   ├── __init__.py
│   └── docx_routes.py          # API route definitions
│
├── services/
│   ├── __init__.py
│   ├── document_editor.py      # Post-processing and grammar fixes
│   ├── fac_api.py              # FAC API integration
│   ├── html_converter.py       # HTML to DOCX conversion
│   ├── mdl_builder.py          # MDL model building logic
│   ├── storage.py              # Azure/local storage operations
│   └── template_processor.py   # Template placeholder replacement
│
├── utils/
│   ├── __init__.py
│   ├── docx_utils.py           # DOCX manipulation utilities
│   └── text_utils.py           # Text processing utilities
│
├── templates/                  # DOCX template files
│   └── MDL_Template_Data_Mapping_Comments.docx
│
└── tests/
    ├── __init__.py
    ├── conftest.py             # Pytest fixtures
    ├── test_config.py
    ├── test_document_editor.py
    ├── test_docx_utils.py
    ├── test_fac_api.py
    ├── test_html_converter.py
    ├── test_mdl_builder.py
    ├── test_routes.py
    ├── test_schemas.py
    ├── test_storage.py
    └── test_text_utils.py
```

## Installation

### Prerequisites

- Python 3.9+
- pip

### Key Dependencies

| Package | Purpose |
|---------|---------|
| `fastapi` | Web framework for API endpoints |
| `uvicorn` | ASGI server |
| `python-docx` | DOCX document creation and manipulation |
| `pydantic` | Data validation and settings management |
| `requests` | HTTP client for FAC API |
| `azure-storage-blob` | Azure Blob Storage integration |
| `beautifulsoup4` | HTML parsing |
| `html2docx` | HTML to DOCX conversion |
| `openpyxl` | Excel file processing |
| `Jinja2` | Template rendering |

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd mdl-backend-demo
   ```

2. **Create virtual environment**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install test dependencies** (optional)
   ```bash
   pip install -r requirements-test.txt
   ```

5. **Run the application**
   ```bash
   uvicorn main:app --reload --host 0.0.0.0 --port 8000
   ```

## Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `FAC_API_KEY` | API key for Federal Audit Clearinghouse | Required |
| `FAC_API_BASE` | FAC API base URL | `https://api.fac.gov` |
| `AZURE_STORAGE_CONNECTION_STRING` | Azure Blob Storage connection string | Optional |
| `AZURE_BLOB_CONTAINER` | Azure container name | `mdl-output` |
| `LOCAL_SAVE_DIR` | Local directory for file storage | `./_out` |
| `PUBLIC_BASE_URL` | Base URL for local file access | `http://localhost:8000` |
| `MDL_TEMPLATE_PATH` | Default template file path | Optional |
| `OPENAI_API_KEY` | OpenAI API key for AI-assisted categorization | Optional |
| `AZURITE_SAS_VERSION` | SAS version for Azurite emulator | `2021-08-06` |

### Example .env file

```env
FAC_API_KEY=your-fac-api-key
AZURE_STORAGE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=...
AZURE_BLOB_CONTAINER=mdl-output
LOCAL_SAVE_DIR=./_out
PUBLIC_BASE_URL=http://localhost:8000
OPENAI_API_KEY=your-openai-key
```

## API Reference

### Health Check

```http
GET /healthz
```

Returns service health status.

**Response:**
```json
{
  "ok": true,
  "time": "2024-06-15T10:30:00.000000"
}
```

### Echo (Debug)

```http
POST /echo
```

Echoes the received payload (for debugging).

### Debug Endpoints

```http
GET /debug/env      # Check FAC API key status
GET /debug/storage  # Check storage configuration
GET /debug/sas      # Test Azure SAS URL generation
```

### Build DOCX from HTML

```http
POST /build-docx
```

Builds a DOCX document from provided HTML content.

**Request Body:**
```json
{
  "auditee_name": "City of Ann Arbor",
  "ein": "386004534",
  "audit_year": 2023,
  "body_html": "<p>Document content...</p>",
  "dest_path": "output/",
  "filename": "custom-name.docx"
}
```

### Build DOCX from FAC Data

```http
POST /build-docx-from-fac
```

Builds MDL document from provided FAC data arrays.

**Request Body:**
```json
{
  "auditee_name": "City of Ann Arbor",
  "ein": "386004534",
  "audit_year": 2023,
  "fac_general": [...],
  "fac_findings": [...],
  "fac_findings_text": [...],
  "fac_caps": [...],
  "federal_awards": [...]
}
```

### Build DOCX by Report ID

```http
POST /build-docx-by-report
```

Fetches FAC data by report ID and generates MDL document.

**Request Body:**
```json
{
  "auditee_name": "City of Ann Arbor",
  "ein": "386004534",
  "audit_year": 2023,
  "report_id": "2023-06-GSAFAC-0000123456",
  "only_flagged": false,
  "max_refs": 15,
  "include_awards": true
}
```

### Build MDL DOCX (Templated)

```http
POST /build-mdl-docx-by-report-templated
```

Builds MDL using a DOCX template with placeholder replacement.

**Request Body:**
```json
{
  "auditee_name": "City of Ann Arbor",
  "ein": "386004534",
  "audit_year": 2023,
  "report_id": "2023-06-GSAFAC-0000123456",
  "template_path": "templates/MDL_Template.docx",
  "auditor_name": "Rehmann Robson LLC",
  "fy_end_text": "June 30, 2023",
  "recipient_name": "City of Ann Arbor",
  "street_address": "123 Main Street",
  "city": "Ann Arbor",
  "state": "MI",
  "zip_code": "48104",
  "poc_name": "John Smith",
  "poc_title": "Finance Director",
  "treasury_listings": ["21.027", "21.023"]
}
```

### Build MDL DOCX (Automatic)

```http
POST /build-mdl-docx-auto
```

Fully automatic MDL generation - fetches all data from FAC by EIN and year.

**Request Body:**
```json
{
  "auditee_name": "City of Ann Arbor",
  "ein": "386004534",
  "audit_year": 2023,
  "treasury_listings": ["21.027"],
  "max_refs": 15,
  "only_flagged": false,
  "include_awards": true
}
```

**Response:**
```json
{
  "ok": true,
  "url": "https://storage.blob.core.windows.net/mdl-output/MDL-City_of_Ann_Arbor-386004534-2023.docx?sv=...",
  "blob_path": "mdl-output/mdl/2023/MDL-City_of_Ann_Arbor-386004534-2023.docx"
}
```

### Local File Access

```http
GET /local/{path}
```

Serves locally stored DOCX files.

## Data Models

### BuildDocx

Basic document build request with HTML content.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `auditee_name` | string | Yes | Name of the auditee organization |
| `ein` | string | Yes | Employer Identification Number |
| `audit_year` | integer | Yes | Fiscal year of the audit |
| `body_html` | string | No | HTML content for document body |
| `body_html_b64` | string | No | Base64-encoded HTML content |
| `dest_path` | string | No | Destination folder path |
| `filename` | string | No | Custom output filename |

### BuildFromFAC

Build request with pre-fetched FAC data arrays.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `auditee_name` | string | Yes | Name of the auditee organization |
| `ein` | string | Yes | Employer Identification Number |
| `audit_year` | integer | Yes | Fiscal year of the audit |
| `fac_general` | array | No | General audit information |
| `fac_findings` | array | No | Audit findings data |
| `fac_findings_text` | array | No | Finding descriptions |
| `fac_caps` | array | No | Corrective action plans |
| `federal_awards` | array | No | Federal awards data |

### BuildAuto

Fully automatic build with all options.

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `auditee_name` | string | Yes | - | Name of the auditee |
| `ein` | string | Yes | - | Employer Identification Number |
| `audit_year` | integer | Yes | - | Fiscal year of the audit |
| `max_refs` | integer | No | 15 | Maximum findings to include |
| `only_flagged` | boolean | No | false | Only include flagged findings |
| `include_awards` | boolean | No | true | Include federal awards data |
| `treasury_listings` | array | No | null | Filter by ALN numbers |
| `template_path` | string | No | null | Custom template path |
| `treasury_contact_email` | string | No | null | Contact email for letter |

## Services

### MDL Builder (`services/mdl_builder.py`)

Core service for building MDL data models from FAC data.

**Key Functions:**
- `build_mdl_model_from_fac()` - Builds complete MDL model from FAC data
- `render_mdl_html()` - Renders MDL model as HTML
- `format_letter_date()` - Formats dates for letter headers
- `summarize_finding_text()` - Creates finding summaries
- `best_summary_label()` - Matches findings to standard categories

### Template Processor (`services/template_processor.py`)

Handles DOCX template processing and placeholder replacement.

**Key Functions:**
- `build_docx_from_template()` - Main template processing function
- `replace_placeholders_docwide()` - Replaces placeholders throughout document
- `insert_program_tables_at_anchor()` - Inserts finding tables
- `build_program_table()` - Creates formatted finding tables

### Document Editor (`services/document_editor.py`)

Post-processing for generated documents.

**Key Functions:**
- `postprocess_docx()` - Final document cleanup and formatting
- `apply_mdl_grammar()` - Fixes grammar throughout document
- `fix_mdl_grammar_text()` - Fixes singular/plural agreement
- `replace_email_with_mailto_link()` - Adds clickable email links

### FAC API (`services/fac_api.py`)

Integration with Federal Audit Clearinghouse API.

**Key Functions:**
- `fac_get()` - Makes authenticated GET requests to FAC API
- `from_fac_general()` - Extracts header data from general records
- `aln_overrides_from_summary()` - Parses ALN data from FAC Excel summaries

### Storage (`services/storage.py`)

File storage operations for Azure Blob Storage and local filesystem.

**Key Functions:**
- `upload_and_sas()` - Uploads to Azure Blob and returns SAS URL
- `save_local_and_url()` - Saves locally and returns access URL
- `blob_service_client()` - Creates Azure Blob service client

### HTML Converter (`services/html_converter.py`)

Converts HTML content to DOCX format.

**Key Functions:**
- `html_to_docx_bytes()` - Converts HTML string to DOCX bytes
- `basic_html_to_docx()` - Basic HTML parsing with formatting support
- `apply_inline_formatting()` - Handles bold, italic, underline tags

## Testing

### Running Tests

```bash
# Run all tests
pytest

# Run with verbose output
pytest -v

# Run with coverage report
pytest --cov=. --cov-report=term-missing

# Run specific test file
pytest tests/test_mdl_builder.py

# Run tests matching pattern
pytest -k "test_build"

# Run only unit tests
pytest -m unit
```

### Test Coverage

The test suite includes **200 tests** covering:

| Module | Coverage |
|--------|----------|
| config.py | 100% |
| models/schemas.py | 100% |
| utils/text_utils.py | 100% |
| main.py | 92% |
| services/html_converter.py | 91% |
| services/mdl_builder.py | 76% |
| routes/docx_routes.py | 67% |
| services/document_editor.py | 60% |

### Test Fixtures

Common fixtures are defined in `tests/conftest.py`:

- `test_client` - FastAPI test client
- `sample_fac_general` - Sample FAC general data
- `sample_fac_findings` - Sample findings data
- `sample_fac_findings_text` - Sample finding text data
- `sample_fac_caps` - Sample corrective action plans
- `sample_federal_awards` - Sample federal awards
- `sample_mdl_model` - Complete MDL model for testing
- `mock_azure_storage` - Mocked Azure storage client
- `mock_fac_api` - Mocked FAC API responses
- `temp_docx_file` - Temporary DOCX template

## Development

### Code Style

- Follow PEP 8 guidelines
- Use type hints for function parameters and return values
- Document functions with docstrings

### Adding New Endpoints

1. Define request model in `models/schemas.py`
2. Add route handler in `routes/docx_routes.py`
3. Implement business logic in appropriate service
4. Add tests in `tests/test_routes.py`

### Adding New Services

1. Create service file in `services/`
2. Export from `services/__init__.py`
3. Add comprehensive tests in `tests/`

### Template Placeholders

Templates support the following placeholders:

| Placeholder | Description |
|-------------|-------------|
| `[Recipient Name]` | Auditee/recipient name |
| `[EIN]` | Employer Identification Number |
| `[Street Address]` | Street address line |
| `[City]` | City name |
| `[State]` | State abbreviation |
| `[Zip Code]` | ZIP code |
| `[Auditor Name]` | Auditing firm name |
| `[Fiscal Year End Date]` | Period end date |
| `[POC Name]` | Point of contact name |
| `[POC Title]` | Point of contact title |
| `[[PROGRAM_TABLES]]` | Anchor for program tables |

### Compliance Type Codes

| Code | Description |
|------|-------------|
| A | Activities allowed or unallowed |
| B | Allowable costs/cost principles |
| C | Cash management |
| E | Eligibility |
| F | Equipment and real property management |
| G | Matching, level of effort, earmarking |
| H | Period of performance |
| I | Procurement and suspension and debarment |
| J | Program income |
| L | Reporting |
| M | Subrecipient monitoring |
| N | Special tests and provisions |
| P | Other |

## License

[Add license information]

## Contributing

[Add contribution guidelines]

## Support

For questions or issues, please [create an issue](link-to-issues) or contact the development team.
