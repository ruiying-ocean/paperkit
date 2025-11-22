"""
Conversion module for LaTeX and Word documents
"""

import subprocess
from pathlib import Path
from .formatter import apply_formatting
from .config import DEFAULT_CONFIG


def convert_to_docx(input_file, output_file=None, config=None):
    """
    Convert LaTeX or Word to formatted Word document.

    Supports:
        - .tex files (via pandoc)
        - .docx files (apply formatting only)

    Args:
        input_file: Path to input file (.tex or .docx)
        output_file: Path to output .docx file
        config: Optional configuration dict

    Returns:
        True if successful, False otherwise
    """

    cfg = DEFAULT_CONFIG.copy()
    if config:
        cfg.update(config)

    input_path = Path(input_file)

    if not input_path.exists():
        print(f"✗ Error: File not found: {input_file}")
        return False

    # Determine input type
    suffix = input_path.suffix.lower()

    if suffix == '.tex':
        return _convert_tex_to_docx(input_file, output_file, cfg)
    elif suffix == '.docx':
        return _convert_docx_to_docx(input_file, output_file, cfg)
    else:
        print(f"✗ Error: Unsupported file type: {suffix}")
        print(f"   Supported: .tex, .docx")
        return False


def _convert_tex_to_docx(tex_file, output_file, config):
    """Convert LaTeX to formatted Word document (two-step process)."""

    if output_file is None:
        output_file = Path(tex_file).with_suffix('.docx')

    print()
    print("=" * 60)
    print("LaTeX to Word Conversion")
    print("=" * 60)
    print(f"Input:  {tex_file}")
    print(f"Output: {output_file}")
    print()

    # Check if pandoc is installed
    try:
        subprocess.run(['pandoc', '--version'],
                      capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("✗ Error: Pandoc not found")
        print("   Install with: brew install pandoc")
        return False

    # Step 1: Convert with pandoc
    print("Step 1/2: Converting LaTeX to Word with pandoc...")

    temp_file = Path(output_file).parent / f".temp_{Path(output_file).name}"

    pandoc_cmd = [
        'pandoc',
        str(tex_file),
        '-s',
        '-o', str(temp_file)
    ]

    # Add bibliography if specified in config
    if 'bibliography' in config and config['bibliography']:
        pandoc_cmd.extend(['--bibliography', config['bibliography']])
        pandoc_cmd.extend(['--citeproc'])
        pandoc_cmd.extend(['--csl', config['csl_style']])

    try:
        result = subprocess.run(
            pandoc_cmd,
            capture_output=True,
            text=True,
            timeout=300
        )

        if not temp_file.exists():
            print(f"✗ Pandoc conversion failed")
            if result.stderr:
                print(f"   {result.stderr}")
            return False

        print(f"✓ Raw conversion complete")

    except subprocess.TimeoutExpired:
        print("✗ Conversion timed out (5 minutes)")
        return False
    except Exception as e:
        print(f"✗ Conversion failed: {e}")
        return False

    # Step 2: Apply formatting
    print()
    print("Step 2/2: Applying formatting...")

    success = apply_formatting(str(temp_file), str(output_file), config)

    # Clean up temp file
    if temp_file.exists():
        temp_file.unlink()

    return success


def _convert_docx_to_docx(input_file, output_file, config):
    """Apply formatting to existing Word document."""

    if output_file is None:
        output_file = input_file

    print()
    print("=" * 60)
    print("Word Document Formatting")
    print("=" * 60)
    print(f"Input:  {input_file}")
    print(f"Output: {output_file}")
    print()

    return apply_formatting(input_file, output_file, config)
