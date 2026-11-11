"""
Configuration settings for the Automated Manuscript Template Compliance Checker
"""

# Application Settings
APP_TITLE = "Automated Manuscript Template Compliance Checker"
APP_VERSION = "1.0.0"
APP_AUTHOR = "FYP Project"

# Default formatting rules (JIWE template style)
DEFAULT_RULES = {
    "margins": {
        "left": 1.0,
        "right": 1.0,
        "top": 1.0,
        "bottom": 1.0
    },
    "journal_header": {
      "font_name": "Book Antiqua",
      "font_size": 24,
      "bold": True,
      "alignment": "CENTER"
    },

    "title": {
      "font_name": "Times New Roman",
      "font_size": 14,
      "bold": True,
      "alignment": "CENTER"
    },
    "author": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "alignment": "CENTER"
    },
    "affiliation": {
      "font_name": "Times New Roman",
      "font_size": 10,
      "bold": False,
      "alignment": "CENTER"
    },
    "corresponding_author": {
      "font_name": "Times New Roman",
      "font_size": 10,
      "bold": False,
      "italic": False,
      "alignment": "CENTER",
      "email_required": True,
      "orcid_required": True
    },
    "abstract": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "max_words": 300
    },
    "abstract_heading": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "all_caps": True,
      "alignment": "LEFT"
    },
    "keywords": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "min_count": 5,
      "bold": False,
      "italic": False,
      "alignment": "JUSTIFY"
    },
    "keywords_heading": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "all_caps": False,
      "alignment": "LEFT"
    },
    "heading": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "italic": False,
      "all_caps":False,
      "alignment":"JUSTIFY",
      "space_before": 0,
      "space_after": 8,
      "line_spacing": 1.0
    },
    "subheading": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "italic": True,
      "alignment": "JUSTIFY",
      "space_before": 0,
      "space_after": 8,
      "line_spacing": 1.0
    },
    "body": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "line_spacing":1.0,
      "bold":False,
      "space_after":8,
      "space_before":0,
      "alignment": "JUSTIFY"
    },
    "caption": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True,
      "italic": False,
      "space_before": 0,
      "space_after": 8,
      "line_spacing": 1.0,
      "alignment": "CENTER"
    },
    "table_content": {
      "font_name": "Times New Roman",
      "font_size": 11
    },
    "reference": {
      "font_name": "Times New Roman",
      "font_size": 8,
      "bold": False,
      "alignment": "JUSTIFY",
      "line_spacing": 1.0,
      "space_before": 0,
      "space_after": 8,
      "hanging_indent": 0.5,
      "style": "APA7",
      "publication_italic_required": True,
      "doi_required": True
    },
    "reference_heading": {
      "font_name": "Times New Roman",
      "font_size": 11,
      "bold": True
    },
    "layout": {
      "columns": 1,
      "page_size": "A4",
      "orientation": "PORTRAIT"
    }
}

# Paragraph classification patterns
JOURNAL_HEADER_PATTERNS = [
    r'journal\s+of',
    r'vol\.\s*\d+',
    r'volume\s*\d+',
    r'issue\s*\d+',
    r'issn[:\s]*[\d\-]+',
    r'doi[:\s]*10\.',
    r'https?://',
    r'www\.',
    r'\u00a9\s*\d{4}',
    r'copyright',
    r'open\s*access',
    r'received[:\s]*\d',
    r'accepted[:\s]*\d',
    r'published[:\s]*\d',
    r'e-issn',
    r'p-issn',
]

AUTHOR_INFO_PATTERNS = [
    r'@[\w\.-]+\.\w+',  # Email pattern
    r'university',
    r'universiti',  # Malay
    r'fakulti',  # Malay
    r'faculty',
    r'department',
    r'jabatan',  # Malay
    r'college',
    r'institute',
    r'institut',
    r'school\s+of',
    r'center\s+for',
    r'centre\s+for',
    r'laboratory',
    r'lab\s+of',
    r'orcid',
    r'\d{4}-\d{4}-\d{4}-\d{4}',  # ORCID pattern
    r'corresponding\s+author',
    r'\*.*author',
    r'author.*\*',
    r'^[A-Za-z]+\s+[A-Za-z]+\s*,',  # Name, affiliation pattern
    r'^[A-Za-z]+\s+[A-Za-z]+\s*\d',  # Name with number superscript
]

SECTION_HEADING_PATTERNS = [
    r'^abstract$',
    r'^keywords?$',
    r'^key\s*words?$',
    r'^introduction$',
    r'^background$',
    r'^literature\s+review$',
    r'^related\s+work',
    r'^methodology$',
    r'^method(s)?$',
    r'^materials?\s+(and|&)\s+methods?$',
    r'^experimental',
    r'^results?$',
    r'^findings?$',
    r'^discussion$',
    r'^analysis$',
    r'^results?\s+(and|&)\s+discussion$',
    r'^conclusion(s)?$',
    r'^summary$',
    r'^recommendations?$',
    r'^future\s+work',
    r'^references?$',
    r'^bibliography$',
    r'^works?\s+cited',
    r'^acknowledgements?$',
    r'^funding',
    r'^conflict\s+of\s+interest',
    r'^data\s+availability',
    r'^data\s+access\s+statement',
    r'^authors?[\u2019\']?\s+contribution',
    r'^ethics?\s+statement',
    r'^authors?[\u2019\']?\s+biograph(y|ies)',
    r'^ai\s+declaration',
    r'^appendix',
    r'^appendices',
    # Numbered headings - MUST be followed by a known section word to avoid matching body text
    # Matches: "1. Introduction", "2 Methodology", "3. Results" etc.
    # Does NOT match: "1. This is body text", "2 items were found"
    r'^\d+\.?\s+(introduction|background|methodology|method|methods|results?|discussion|conclusion|conclusions|references?|abstract|summary|findings?|analysis|experiments?|implementation|evaluation|related\s+work|literature\s+review|future\s+work|acknowledgements?)',
    r'^[IVXLC]+\.?\s+(introduction|background|methodology|method|results?|discussion|conclusion|references?)',
    r'^[A-Z]\.\s+(introduction|background|methodology|method|results?|discussion|conclusion|references?)',
]

CAPTION_PATTERNS = [
    r'^figure\s*\d+',
    r'^fig\.\s*\d+',
    r'^table\s*\d+',
    r'^chart\s*\d+',
    r'^diagram\s*\d+',
    r'^image\s*\d+',
    r'^plate\s*\d+',
    r'^graph\s*\d+',
    r'^rajah\s*\d+',  # Malay - figure
    r'^jadual\s*\d+',  # Malay - table
]

REFERENCE_PATTERNS = [
    r'^\[\d+\]',  # IEEE style [1]
    r'^\d+\.\s+\w',  # Numbered list
    r'^\(\d{4}\)',  # Year in parentheses
    r',\s*\d{4}\.',  # Year with comma
    r'et\s+al\.',
    r'pp\.\s*\d+',
    r'vol\.\s*\d+',
    r'doi:',
    r'isbn',
    r'retrieved\s+from',
    r'accessed',
]

# Required document sections
REQUIRED_SECTIONS = [
    "abstract",
    "keywords",
    "introduction",
    "literature review",
    "methods",
    "results",
    "discussion",
    "conclusion",
    "references"
]

# LLM Configuration (optional explanation layer)
LLM_CONFIG = {
    "nvidia_api_key": "",  # Load from .env file for security
    "nvidia_model": "meta/llama-3.1-8b-instruct",
    "max_tokens": 500,
    "temperature": 0.1,
    "timeout_seconds": 20
}

# Alignment mapping
ALIGNMENT_MAP = {
    0: "LEFT",
    1: "CENTER", 
    2: "RIGHT",
    3: "JUSTIFY"
}

ALIGNMENT_REVERSE_MAP = {
    "LEFT": 0,
    "CENTER": 1,
    "RIGHT": 2,
    "JUSTIFY": 3
}

# Color codes for UI
COLORS = {
    "error": "#ffcccc",  # Light red
    "correct": "#ccffcc",  # Light green
    "warning": "#ffffcc",  # Light yellow
    "info": "#cce5ff",  # Light blue
    "error_text": "#cc0000",
    "correct_text": "#008000"
}

# File size limits
MAX_FILE_SIZE_MB = 50