#!/usr/bin/env python3
"""
Command-line interface for Academic Paper Toolkit
"""

import sys
from pathlib import Path
from .formatter import apply_formatting
from .converter import convert_to_docx
from .initializer import init_paper
from .templates import print_templates


def print_help():
    """Print help message."""
    print("""
DocTool - Academic manuscript toolkit
====================================================================

Commands:
    init       Initialize a new paper with proper formatting
    convert    Convert LaTeX/Word to formatted Word document
    format     Apply formatting to existing Word document
    templates  List available journal templates
    help       Show this help message

Usage:
    python -m doctool init "Paper Title" [output.docx] [--template JOURNAL]
    python -m doctool convert input.tex [output.docx]
    python -m doctool format input.docx [output.docx]
    python -m doctool templates

Examples:
    # Create new paper with default template
    python -m doctool init "Climate Change Impacts" paper.docx

    # Create new paper for specific journal
    python -m doctool init "My Research" paper.docx --template agu
    python -m doctool init "My Research" paper.docx --template nature
    python -m doctool init "My Research" paper.docx --template science

    # Convert LaTeX to Word
    python -m doctool convert manuscript.tex manuscript.docx

    # Format existing Word document
    python -m doctool format draft.docx formatted.docx

    # List available templates
    python -m doctool templates

Configuration:
    - Font: Arial, 12pt
    - Line spacing: 1.5
    - Margins: 1 inch
    - Headings: Bold, black, sized appropriately
    - Language: UK English (en-GB)
    - Citations: APA style (author-year)
    """)


def main():
    """Main CLI entry point."""

    if len(sys.argv) < 2:
        print_help()
        sys.exit(1)

    command = sys.argv[1].lower()

    try:
        if command == 'help' or command == '--help' or command == '-h':
            print_help()
            sys.exit(0)

        elif command == 'templates':
            print_templates()
            sys.exit(0)

        elif command == 'init':
            if len(sys.argv) < 3:
                print("✗ Error: Title required")
                print('Usage: python -m doctool init "Paper Title" [output.docx] [--template JOURNAL]')
                sys.exit(1)

            title = sys.argv[2]

            # Parse arguments
            output = "paper.docx"
            template = None

            i = 3
            while i < len(sys.argv):
                if sys.argv[i] == '--template' or sys.argv[i] == '-t':
                    if i + 1 < len(sys.argv):
                        template = sys.argv[i + 1]
                        i += 2
                    else:
                        print("✗ Error: --template requires a value")
                        sys.exit(1)
                else:
                    output = sys.argv[i]
                    i += 1

            success = init_paper(title, output, template=template)
            sys.exit(0 if success else 1)

        elif command == 'convert':
            if len(sys.argv) < 3:
                print("✗ Error: Input file required")
                print("Usage: python -m doctool convert input.tex [output.docx]")
                sys.exit(1)

            input_file = sys.argv[2]
            output_file = sys.argv[3] if len(sys.argv) > 3 else None

            # Check for bibliography config
            config = {}
            bib_file = Path(input_file).parent / "library.bib"
            if bib_file.exists():
                config['bibliography'] = str(bib_file)

            success = convert_to_docx(input_file, output_file, config)
            sys.exit(0 if success else 1)

        elif command == 'format':
            if len(sys.argv) < 3:
                print("✗ Error: Input file required")
                print("Usage: python -m doctool format input.docx [output.docx]")
                sys.exit(1)

            input_file = sys.argv[2]
            output_file = sys.argv[3] if len(sys.argv) > 3 else None

            success = apply_formatting(input_file, output_file)
            sys.exit(0 if success else 1)

        else:
            print(f"✗ Error: Unknown command: {command}")
            print("Run 'python -m doctool help' for usage")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n\n⚠️  Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
