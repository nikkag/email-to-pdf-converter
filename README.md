# Email to PDF Converter

A modern Python script that converts `.eml` and `.msg` files to PDFs with date-based naming, following the same logic as the `amdate.py` script. Features a cross-platform file selector dialog and proper dependency management.

## Features

- **Cross-platform directory selection**: Works on both macOS and Windows
- **Multi-format support**: Handles both EML and MSG (Outlook) email files
- **HTML rendering**: Uses Playwright to render HTML emails as they appear in email clients
- **Date-based naming**: Uses email headers to extract dates and create organized filenames
- **Conflict resolution**: Automatically handles filename conflicts by adding counters
- **Modern Python practices**: Type hints, comprehensive error handling, and modular design
- **Dependency management**: Uses `uv` for fast, reliable dependency management

## Installation

### Prerequisites

- Python 3.10 or higher
- `uv` package manager (recommended) or `pip`

### Using uv (Recommended)

```bash
# Install uv if you haven't already
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone or download this project
cd eml-to-pdf-converter

# Install dependencies
uv sync

# Install Playwright browsers (required for HTML rendering)
uv run python -m playwright install chromium

# Run the script
uv run python eml_to_pdf_converter.py
```

### Using pip

```bash
# Install dependencies
pip install -r requirements.txt

# Install Playwright browsers (required for HTML rendering)
python -m playwright install chromium

# Run the script
python eml_to_pdf_converter.py
```

## Usage

1. Run the script: `python eml_to_pdf_converter.py`
2. A file dialog will open - select the directory containing your `.eml` and `.msg` files
3. The script will process all email files and convert them to PDFs
4. A summary will be displayed showing successful conversions and any failures

### HTML Rendering

The script uses Playwright to render HTML emails exactly as they appear in email clients:
- **HTML emails**: Rendered with proper styling, images, and formatting
- **Text emails**: Converted to clean, readable text
- **Fallback**: If browser rendering fails, falls back to text-only PDFs
- **Performance**: Concurrent processing with up to 50 simultaneous conversions

## Naming Convention

The script follows the same naming logic as `amdate.py`:

- **Date format**: `YYYY-MM-DD` (e.g., `2024-01-15`)
- **Filename pattern**: `{date}_Email_{first_three_words}.pdf`
- **Example**: `2024-01-15_Email_John_Doe_Inquiry.pdf`

### Filename Logic

1. Extract the first 3 words from the original email filename (`.eml` or `.msg`)
2. Sanitize the words (remove special characters)
3. Combine with the email date in `YYYY-MM-DD` format
4. Add `_Email_` prefix to indicate the source
5. Handle conflicts by adding `_1`, `_2`, etc.

## Project Structure

```
eml-to-pdf-converter/
├── pyproject.toml          # Project configuration and dependencies
├── eml_to_pdf_converter.py # Main script
├── README.md              # This file
└── tests/                 # Test files (if any)
```

## Dependencies

- **fpdf**: PDF generation library
- **extract-msg**: MSG file parsing library
- **beautifulsoup4**: HTML parsing and cleaning
- **playwright**: Browser automation for HTML rendering (requires Chromium browser)
- **tkinter**: Built-in GUI library for file dialogs
- **email**: Built-in email parsing library

**Note**: Playwright requires the Chromium browser to be installed separately. This is done automatically during installation, but you can manually install it with `python -m playwright install chromium`.

## Development

### Code Quality

The project uses modern Python development tools:

- **Ruff**: Fast Python linter and formatter
- **MyPy**: Static type checking
- **pytest**: Testing framework

### Running Tests

```bash
# Install development dependencies
uv sync --dev

# Install Playwright browsers (required for tests)
uv run python -m playwright install chromium

# Run tests
uv run pytest

# Format code
uv run ruff format .

# Lint code
uv run ruff check .
```

### Type Checking

```bash
uv run mypy eml_to_pdf_converter.py
```

### Building Executables

```bash
# Install build dependencies
uv sync --dev
uv sync --group build

# Build Windows executable
uv run pyinstaller --onefile --windowed --name "EmailToPDFConverter" eml_to_pdf_converter.py

# Build Linux/macOS executable
uv run pyinstaller --onefile --name "EmailToPDFConverter" eml_to_pdf_converter.py
```

**Note**: The built executable will include all dependencies including Playwright browsers, so users don't need to install them separately.

## Error Handling

The script includes comprehensive error handling:

- **Invalid directories**: Checks if selected directory exists
- **Missing dates**: Handles emails without valid date headers
- **File conflicts**: Automatically resolves naming conflicts
- **Encoding issues**: Gracefully handles various text encodings
- **PDF creation errors**: Catches and reports PDF generation failures

## Examples

### Input Files
```
emails/
├── John Doe Inquiry.eml
├── Meeting Notes.eml
└── Client Communication.eml
```

### Output Files
```
emails/
├── 2024-01-15_Email_John_Doe_Inquiry.pdf
├── 2024-01-16_Email_Meeting_Notes.pdf
└── 2024-01-17_Email_Client_Communication.pdf
```

## Troubleshooting

### Common Issues

1. **No .eml files found**: Ensure your selected directory contains `.eml` files
2. **Date parsing errors**: Some emails may have malformed date headers
3. **Permission errors**: Ensure you have write permissions in the target directory
4. **GUI not appearing**: On some systems, you may need to run with a display server
5. **HTML rendering issues**: Ensure Playwright browsers are installed with `python -m playwright install chromium`
6. **Browser initialization fails**: The script will fall back to text-only PDFs if browser rendering fails

### Debug Mode

For detailed error information, you can modify the script to include more verbose logging.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes following the coding guidelines
4. Add tests for new functionality
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Authors

**Nikka Ghalili** - n@nikkag.com
**Dima Ghalili** - dghalili@fdmattorneys.com

## Acknowledgments

- Based on the naming logic from `amdate.py`
- Uses modern Python best practices and type hints
- Cross-platform compatibility with tkinter