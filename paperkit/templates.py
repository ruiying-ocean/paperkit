"""
Journal-specific templates and configurations

Similar to Quarto journal templates, provides formatting
configurations for major academic journals.
"""

# Journal template configurations
JOURNAL_TEMPLATES = {
    'agu': {
        'name': 'American Geophysical Union (AGU)',
        'font': 'Times New Roman',
        'font_size': 12,
        'title_size': 16,
        'heading1_size': 14,
        'heading2_size': 12,
        'heading3_size': 12,
        'line_spacing': 2.0,
        'margins': 1.0,
        'language': 'en-US',
        'paper_size': 'letter',  # US journals use Letter
        'csl_style': 'https://www.zotero.org/styles/american-geophysical-union',
        'citation_type': 'author-year',
        'sections': [
            'Abstract',
            'Introduction',
            'Methods',
            'Results',
            'Discussion',
            'Conclusions',
            'Data Availability Statement',
            'Acknowledgments',
            'References',
        ]
    },

    'nature': {
        'name': 'Nature',
        'font': 'Arial',
        'font_size': 12,
        'title_size': 14,
        'heading1_size': 12,
        'heading2_size': 12,
        'heading3_size': 11,
        'line_spacing': 2.0,
        'margins': 1.0,
        'language': 'en-GB',
        'paper_size': 'a4',  # UK journals use A4
        'csl_style': 'https://www.zotero.org/styles/nature',
        'citation_type': 'numbered',
        'sections': [
            'Abstract',
            'Introduction',
            'Results',
            'Discussion',
            'Methods',
            'Data availability',
            'Code availability',
            'References',
            'Acknowledgements',
            'Author contributions',
            'Competing interests',
        ]
    },

    'science': {
        'name': 'Science',
        'font': 'Times New Roman',
        'font_size': 12,
        'title_size': 14,
        'heading1_size': 12,
        'heading2_size': 12,
        'heading3_size': 11,
        'line_spacing': 2.0,
        'margins': 1.0,
        'language': 'en-US',
        'paper_size': 'letter',  # US journals use Letter
        'csl_style': 'https://www.zotero.org/styles/science',
        'citation_type': 'numbered',
        'sections': [
            'Abstract',
            'Introduction',
            'Results',
            'Discussion',
            'Materials and Methods',
            'References and Notes',
            'Acknowledgments',
            'Supplementary Materials',
        ]
    },

    'pnas': {
        'name': 'Proceedings of the National Academy of Sciences (PNAS)',
        'font': 'Times New Roman',
        'font_size': 11,
        'title_size': 13,
        'heading1_size': 11,
        'heading2_size': 11,
        'heading3_size': 11,
        'line_spacing': 2.0,
        'margins': 1.0,
        'language': 'en-US',
        'paper_size': 'letter',  # US journals use Letter
        'csl_style': 'https://www.zotero.org/styles/pnas',
        'citation_type': 'numbered',
        'sections': [
            'Abstract',
            'Significance Statement',
            'Introduction',
            'Results',
            'Discussion',
            'Materials and Methods',
            'Acknowledgments',
            'References',
        ]
    },

    'default': {
        'name': 'Default (General Academic)',
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
        'citation_type': 'author-year',
        'sections': [
            'Abstract',
            'Introduction',
            'Methods',
            'Results',
            'Discussion',
            'Conclusions',
            'Acknowledgements',
            'Data Availability',
            'Author Contributions',
            'Competing Interests',
            'References',
        ]
    },
}


def get_template(template_name):
    """
    Get journal template configuration.

    Args:
        template_name: Name of the template (e.g., 'agu', 'nature', 'science')

    Returns:
        dict: Template configuration

    Raises:
        ValueError: If template name is not found
    """
    template_name = template_name.lower()

    if template_name not in JOURNAL_TEMPLATES:
        available = ', '.join(JOURNAL_TEMPLATES.keys())
        raise ValueError(
            f"Template '{template_name}' not found. "
            f"Available templates: {available}"
        )

    return JOURNAL_TEMPLATES[template_name].copy()


def list_templates():
    """
    List all available journal templates.

    Returns:
        list: List of (template_key, template_name) tuples
    """
    return [
        (key, config['name'])
        for key, config in JOURNAL_TEMPLATES.items()
    ]


def print_templates():
    """Print available templates in a formatted way."""
    print("\nAvailable Journal Templates:")
    print("=" * 60)

    for key, config in JOURNAL_TEMPLATES.items():
        print(f"\n{key.upper():10s} - {config['name']}")
        print(f"           Font: {config['font']}, {config['font_size']}pt")
        print(f"           Spacing: {config['line_spacing']}")
        print(f"           Citations: {config['citation_type']}")
