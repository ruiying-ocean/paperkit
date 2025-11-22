# PyPaper

A complete Python package for creating, converting, and formatting academic manuscripts in Word format.

## Features

- ✅ Initialize new papers with proper structure and formatting
- ✅ Convert LaTeX to Word with automatic formatting
- ✅ Apply professional formatting to existing Word documents
- ✅ Support for both .tex and .docx inputs
- ✅ Automatic citation handling (APA style)
- ✅ UK English language settings

## Installation

No installation needed! Just ensure you have the dependencies:

```bash
pip install python-docx
brew install pandoc  # For LaTeX conversion
```

## Package Structure

```
pypaper/
├── __init__.py       # Package initialization
├── __main__.py       # Module entry point
├── cli.py            # Command-line interface
├── config.py         # Configuration settings
├── formatter.py      # Document formatting
├── converter.py      # LaTeX/Word conversion
├── initializer.py    # New paper creation
└── README.md         # This file
```

## Usage

### Command Line

```bash
# Initialize a new paper
python -m pypaper init "Your Paper Title" paper.docx

# Convert LaTeX to Word
python -m pypaper convert manuscript.tex output.docx

# Format existing Word document
python -m pypaper format draft.docx formatted.docx

# Show help
python -m pypaper help
```

### As Python Module

```python
from pypaper import init_paper, convert_to_docx, apply_formatting

# Create new paper
init_paper("Research Title", "paper.docx")

# Convert LaTeX
convert_to_docx("manuscript.tex", "output.docx")

# Format existing document
apply_formatting("draft.docx", "formatted.docx")
```

## Default Formatting

- **Font**: Arial, 12pt
- **Line spacing**: 1.5
- **Margins**: 1 inch (all sides)
- **Title**: 16pt, Bold, Black
- **Heading 1**: 14pt, Bold, Black
- **Heading 2**: 12pt, Bold, Black
- **Language**: UK English (en-GB)
- **Citations**: APA style (author-year)

## Customization

You can customize settings by modifying `config.py` or passing a config dict:

```python
custom_config = {
    'font': 'Times New Roman',
    'font_size': 11,
    'line_spacing': 2.0,
}

apply_formatting("input.docx", "output.docx", config=custom_config)
```

## Examples

### Creating a New Paper

```bash
python -m pypaper init "Climate Change Impacts on Marine Ecosystems" climate_paper.docx
```

Creates a document with:
- Title and author placeholders
- Standard sections (Abstract, Introduction, Methods, etc.)
- Proper formatting applied

### Converting LaTeX Manuscript

```bash
python -m pypaper convert main.tex manuscript.docx
```

Two-step process:
1. Converts LaTeX to raw Word (via pandoc)
2. Applies professional formatting

### Formatting Existing Document

```bash
python -m pypaper format draft.docx final.docx
```

Applies all formatting rules to any Word document.

## Requirements

- Python 3.6+
- python-docx
- pandoc (for LaTeX conversion)

## License

Free to use and modify.
