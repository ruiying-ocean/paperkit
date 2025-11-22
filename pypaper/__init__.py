"""
PyPaper - Academic manuscript toolkit

A complete solution for creating, converting, and formatting academic manuscripts.
"""

__version__ = "1.0.0"
__author__ = "PyPaper Contributors"

from .formatter import apply_formatting
from .converter import convert_to_docx
from .initializer import init_paper
from .templates import get_template, list_templates

__all__ = [
    'apply_formatting',
    'convert_to_docx',
    'init_paper',
    'get_template',
    'list_templates',
]
