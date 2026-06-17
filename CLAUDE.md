# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Document Generator is a Python Panel web application for generating personalized documents by merging Excel data with Word templates. It uses the Panelini framework (a wrapper around Panel) for the UI structure and python-docx-template for document templating.

Key technologies:
- **Panel/Panelini**: Interactive web UI framework
- **python-docx-template**: Word document templating with Jinja2-like syntax
- **pandas/openpyxl**: Excel data parsing
- **LibreOffice**: Headless PDF conversion
- **UV**: Fast Python package manager

## Development Commands

### Setup and Installation

```bash
# Install dependencies
uv sync

# Run the application (development)
python src/document_generator/main.py

# Run with Panel's hot reload (alternative)
panel serve src/document_generator/main.py --port 5010 --dev
```

### Docker Commands

```bash
# Build and run with Docker Compose
docker-compose up --build

# Run in background
docker-compose up -d

# View logs
docker logs document-generator

# Stop container
docker-compose down
```

### Environment Configuration

Create a `.env` file with:
```bash
INTERNAL_PORT=5010    # Port inside container
EXTERNAL_PORT=5010    # Port exposed to host
LIBREOFFICE_PATH=libreoffice  # Path to LibreOffice executable
CONVERSION_TIMEOUT=60  # PDF conversion timeout in seconds
```

## Architecture

### Single-File Application Structure

The application is implemented as a single-file monolith in `src/document_generator/main.py`:

1. **DocumentGeneratorApp class**: Main application logic
   - `_create_widgets()`: Initializes all Panel widgets (file uploaders, tables, buttons, previews)
   - `_on_excel_upload()`: Handles Excel file upload, validates/sanitizes column names
   - `_on_template_upload()`: Handles Word template upload
   - `_render_template()`: Static method that renders Word templates with row data using python-docx-template
   - `_convert_to_pdf()`: Converts Word documents to PDF using LibreOffice headless mode
   - `_preview_documents()`: Generates PDF previews for selected rows
   - `_download_documents()`: Core method that generates Word and/or PDF documents as ZIP files
   - `_download_word_documents()`: Callback for Word-only download (uses `_download_documents(as_pdf=False)`)
   - `_download_pdf_documents()`: Callback for PDF-only download (uses `_download_documents(as_pdf=True)`)
   - `_download_all_formats()`: Callback for combined download (creates nested ZIP with both formats)
   - `_create_empty_zip()`: Static helper that creates valid empty ZIP files
   - `_create_zip_file_info()`: Helper for creating ZipFileInfo TypedDict instances
   - `get_sidebar()` / `get_main()`: Returns UI components for layout

2. **Type Definitions**:
   - `FileFormat` enum: Defines format types (WORD, PDF, ALL)
   - `ZipFileInfo` TypedDict: Structure for passing ZIP buffers with format metadata

3. **Panelini integration**: At the bottom of main.py
   - Creates Panelini app instance with title and sidebar
   - Attaches DocumentGeneratorApp's components to sidebar and main area
   - Makes app servable with `app.servable()`
   - Direct serve with `pn.serve()` when run as `__main__`

### Data Flow

1. User uploads Excel file → Column names are sanitized (spaces/special chars → underscores)
2. User uploads Word template(s) → Templates stored in memory with naming patterns
3. User selects rows in data table → Download buttons become enabled
4. User clicks one of the download buttons:
   - **"Download Word Files"**: Generates Word documents only (`documents_word.zip`)
   - **"Download PDF Files"**: Generates Word documents, converts to PDF, returns PDFs only (`documents_pdf.zip`)
   - **"Download Word & PDF Files"**: Generates both, returns nested ZIP containing `documents_word.zip` and `documents_pdf.zip`
   - **"Preview PDFs"**: Generates PDFs and displays them inline for review
5. For each (template, row) combination:
   - Template is rendered with row data using python-docx-template
   - Word document is generated to temp file
   - (Optional) LibreOffice converts to PDF if requested
   - Files are collected into ZIP or shown as previews
   - Temporary files are cleaned up in `finally` blocks

### Excel Column Name Sanitization

The app automatically sanitizes Excel column names to be valid Python identifiers (src/document_generator/main.py:186-226):
- Strips whitespace from column names
- Replaces spaces and hyphens with underscores
- Adds underscore prefix if column name starts with a digit
- Replaces invalid characters (except alphanumeric and underscore) with underscores
- Handles duplicate names by appending numeric suffixes (_1, _2, etc.)
- Reports all renames in status message with warning
- Preserves Unicode characters (umlauts are allowed)

Users must use the sanitized names in their Word templates (e.g., `{{ First_Name }}` for "First Name" column). The status message shows the exact mapping to use.

### Word Template Rendering

Uses python-docx-template (DocxTemplate) which supports Jinja2-like syntax:
- Variable insertion: `{{ column_name }}`
- Conditionals: `{% if condition %}...{% endif %}`
- Loops and filters supported

Templates are rendered with a dictionary of row data where all values are converted to strings.

### PDF Conversion

LibreOffice headless mode is used for Word → PDF conversion (src/document_generator/main.py:369-420):
- Command: `libreoffice --headless --convert-to pdf <input> --outdir <dir>`
- Subprocess with configurable timeout (default 60 seconds, configurable via `CONVERSION_TIMEOUT` env var)
- Temporary files are created and cleaned up after conversion in `finally` blocks
- Requires LibreOffice to be installed on the system
- **Error Handling**: Individual conversion failures don't stop the batch process
  - Failed conversions are logged with document name and error
  - User gets a warning notification showing failure count
  - Word documents are still available for download even if PDF conversion fails
- PDF path is determined by replacing `.docx` extension with `.pdf`
- Waits for PDF file to exist (with timeout) after conversion command completes

### UI Layout Pattern

The Panelini framework wraps Panel and provides:
- `app.sidebar_set()`: Sets sidebar components (wrapped in Cards)
- `app.main_set()`: Sets main area components
- Automatic styling and layout management
- Cards with collapsible sections for organization

## Important Implementation Details

### Threading and Deployment

From recent commit (c8f0c6b): The app uses `python src/document_generator/main.py` to serve directly instead of `panel serve` to prevent unexpected behavior and ensure threading support (see https://panel.holoviz.org/how_to/concurrency/threading.html).

### FileDownload Widget Pattern

The app uses three Panel `FileDownload` widgets (lines 100-132) with callbacks that return `io.BytesIO` objects:

1. **`word_file_download`**: Downloads `documents_word.zip` containing only Word (.docx) files
   - Callback: `_download_word_documents()` → calls `_download_documents(as_pdf=False)`

2. **`pdf_file_download`**: Downloads `documents_pdf.zip` containing only PDF files
   - Callback: `_download_pdf_documents()` → calls `_download_documents(as_pdf=True)`
   - Handles LibreOffice conversion failures gracefully (see Error Handling below)

3. **`all_formats_download`**: Downloads `documents.zip` containing nested ZIPs
   - Callback: `_download_all_formats()` → creates nested ZIP structure
   - Contains `documents_word.zip` and `documents_pdf.zip` inside

**Core Implementation**: The `_download_documents(as_pdf: bool)` method (lines 559-710) is the workhorse that:
- Validates templates and selection early (returns valid empty ZIPs if validation fails)
- Generates Word documents for all (template × row) combinations
- Optionally converts to PDF if `as_pdf=True`
- Returns a list of `ZipFileInfo` dicts containing format type and BytesIO buffer
- Tracks PDF conversion failures individually without stopping the entire process

**Empty ZIP Handling**: The helper `_create_empty_zip()` (lines 319-326) ensures that even when no documents are generated (validation failures, errors, no selection), the download returns a valid empty ZIP file instead of a corrupted/invalid BytesIO object.

### Error Handling and Resilience

The application implements robust error handling to ensure partial failures don't crash the entire process:

**PDF Conversion Failures** (lines 646-663):
- Each PDF conversion is wrapped in a try-except block
- Failures are logged and tracked in a `conversion_failures` list
- Processing continues for remaining documents
- User receives a warning notification showing how many conversions failed
- Word documents are still generated successfully even if PDF conversion fails
- Detailed error information is logged for debugging

**Early Validation** (lines 571-580):
- Checks for uploaded templates before attempting document generation
- Checks for row selection before processing
- Returns valid empty ZIPs with appropriate error notifications if validation fails

**Defensive Buffer Management** (line 768):
- All buffer reads are preceded by `seek(0)` to ensure position is at start
- Prevents corrupted nested ZIPs in combined download mode

**Logging** (lines 16, 568, 597, 663, 682, 692-694, 763, 774):
- Logger instance: `logging.getLogger("DocumentGeneratorApp")`
- Logs download start with parameters (template count, as_pdf flag)
- Logs processing details (rows × templates)
- Logs ZIP sizes after creation
- Logs PDF conversion failures with document names and errors
- Makes debugging issues much easier

### Temporary File Management

All document generation uses temporary files/directories that are cleaned up in `finally` blocks to prevent disk space leaks.

### Button State Management

`_update_button_states()` is called whenever data/templates/selection changes to enable/disable action buttons based on application state.

## Docker Deployment

The Docker setup (from Dockerfile and docker-compose.yml):
- Base: `python:3.12-slim-trixie` with UV package manager
- Installs LibreOffice and Java support for PDF conversion
- Volume mounts: Source code at `/app/src` for live updates, separate volume for `.venv`
- Uses `uv run python` to execute main.py directly
- Environment variables passed via `.env` file

## Recent Fixes and Known Issues

### Recently Fixed (2026-02-06)

**Download Button Issues Resolved**:
- Fixed PDF-only and combined download buttons that were not working correctly
- Empty BytesIO objects replaced with valid empty ZIP files (prevents corrupted downloads)
- Added defensive buffer management with `seek(0)` before reads
- PDF conversion failures now tracked individually without crashing entire process
- Better error handling: Word downloads succeed even if PDF conversion fails
- Added comprehensive logging for debugging download issues
- Early validation for templates and row selection

### Known Issues

- Docker logs may not show properly (todo comment in docker-compose.yml:22)

### Recent Improvements

From git history:
- UI background/layout fixed (fix-ui-background-layout branch)
- Sidebar elements wrapped in Card components to preserve background spacing
- Download system refactored with three separate FileDownload widgets for better UX
