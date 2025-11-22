# PaperKit

A complete Python package for creating, converting, and formatting academic manuscripts in Word format.

## Features

- Initialize new papers with proper structure and formatting
- Convert LaTeX to Word with automatic formatting
- Apply professional formatting to existing Word documents
- Journal-specific templates (AGU, Nature, Science, PNAS)
- Support for both .tex and .docx inputs
- Automatic citation handling (APA style)
- Flexible paper sizes (A4, Letter, Legal, etc.)
- UK/US English language settings

## Installation

### From source (using uv)

```bash
cd ~/paperkit
uv pip install -e .
```

### From source (using pip)

```bash
cd ~/paperkit
pip install -e .
```

### Dependencies

```bash
# Install Python dependencies (uv)
uv pip install python-docx

# Or using pip
pip install python-docx

# Install pandoc (required for LaTeX conversion)
brew install pandoc  # macOS
# sudo apt install pandoc  # Linux
```

## Quick Start

### Command Line

Using `uv run` (recommended):

```bash
# List available journal templates
uv run paperkit templates

# Create new paper with default template
uv run paperkit init "Your Paper Title" paper.docx

# Create paper for specific journal
uv run paperkit init "Research Title" paper.docx --template agu
uv run paperkit init "Research Title" paper.docx --template nature

# Convert LaTeX to Word
uv run paperkit convert manuscript.tex output.docx

# Format existing Word document
uv run paperkit format draft.docx formatted.docx
```

Using `python -m` (alternative):

```bash
# List available journal templates
python -m paperkit templates

# Create new paper with default template
python -m paperkit init "Your Paper Title" paper.docx

# Create paper for specific journal
python -m paperkit init "Research Title" paper.docx --template agu
python -m paperkit init "Research Title" paper.docx --template nature

# Convert LaTeX to Word
python -m paperkit convert manuscript.tex output.docx

# Format existing Word document
python -m paperkit format draft.docx formatted.docx
```

### As Python Module

```python
from paperkit import init_paper, convert_to_docx, apply_formatting

# Create new paper
init_paper("Research Title", "paper.docx", template='nature')

# Convert LaTeX
convert_to_docx("manuscript.tex", "output.docx")

# Format existing document
apply_formatting("draft.docx", "formatted.docx")
```

## Journal Templates

### Available Templates

- **AGU** - American Geophysical Union (Times New Roman, Letter, author-year)
- **Nature** - Nature journal (Arial, A4, numbered citations)
- **Science** - Science journal (Times New Roman, Letter, numbered)
- **PNAS** - PNAS journal (Times New Roman, Letter, numbered)
- **Default** - General academic (Arial, A4, author-year)

### Template Features

Each template includes:
- Journal-specific fonts and sizes
- Correct line spacing
- Proper section structure
- Citation style (author-year vs numbered)
- Paper size (A4 vs Letter)
- Language settings (US/UK English)

## Paper Sizes

Supported paper sizes:
- `a4` - 210mm × 297mm (default)
- `letter` - US Letter (8.5" × 11")
- `legal` - US Legal (8.5" × 14")
- `a5` - 148mm × 210mm
- `b5` - 176mm × 250mm

## Default Formatting

- **Font**: Arial, 12pt
- **Line spacing**: 1.5 (varies by template)
- **Margins**: 1 inch (all sides)
- **Title**: 16pt, Bold, Black
- **Heading 1**: 14pt, Bold, Black
- **Heading 2**: 12pt, Bold, Black
- **Language**: UK English (en-GB)
- **Citations**: APA style (author-year)
- **Paper size**: A4

## Package Structure

```
paperkit/
├── __init__.py       # Package exports
├── __main__.py       # Module entry point
├── cli.py            # Command-line interface
├── config.py         # Configuration settings
├── formatter.py      # Document formatting
├── converter.py      # LaTeX/Word conversion
├── initializer.py    # New paper creation
├── templates.py      # Journal templates
└── README.md         # Package documentation
```

## Development

### Testing

```bash
# Test package installation
python -m paperkit help

# Create test documents
python -m paperkit init "Test Paper" test.docx
python -m paperkit templates
```

## Requirements

- Python 3.9+
- python-docx >= 0.8.0
- pandoc (for LaTeX conversion)
