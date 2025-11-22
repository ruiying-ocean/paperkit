"""
Configuration settings for academic papers
"""

# Paper size definitions (width, height in inches)
PAPER_SIZES = {
    'a4': (8.27, 11.69),      # 210mm × 297mm
    'letter': (8.5, 11.0),    # US Letter
    'legal': (8.5, 14.0),     # US Legal
    'a5': (5.83, 8.27),       # 148mm × 210mm
    'b5': (6.93, 9.84),       # 176mm × 250mm
}

DEFAULT_CONFIG = {
    'font': 'Arial',
    'font_size': 12,
    'title_size': 16,
    'heading1_size': 14,
    'heading2_size': 12,
    'heading3_size': 12,
    'line_spacing': 1.5,
    'margins': 1.0,
    'language': 'en-GB',
    'csl_style': 'https://www.zotero.org/styles/apa',
    'paper_size': 'a4',  # Default to A4
}
