"""
Paragraph classification module.

The checker uses deterministic rules as the primary classifier because
format validation and auto-fixing must be stable and explainable. LLM support
is intentionally kept out of the core classification path.
"""

import re
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, List

from .utils import get_paragraph_alignment, get_paragraph_font_info, get_paragraph_text


class ParagraphType(Enum):
    """Paragraph type labels used by the checker and auto-fixer."""

    EMPTY = "empty"
    JOURNAL_HEADER = "journal_header"
    AUTHOR_INFO = "author_info"
    ABSTRACT_LABEL = "abstract_label"
    KEYWORDS_LABEL = "keywords_label"
    PAPER_TITLE = "paper_title"
    SECTION_HEADING = "section_heading"
    BODY = "body"
    ABSTRACT_CONTENT = "abstract_content"
    KEYWORDS_CONTENT = "keywords_content"
    CAPTION = "caption"
    REFERENCE = "reference"
    TABLE = "table"
    FIGURE = "figure"
    ALGORITHM = "algorithm"
    CODE_BLOCK = "code_block"
    EQUATION = "equation"
    TEMPLATE_INSTRUCTION = "template_instruction"
    UNKNOWN = "unknown"


@dataclass
class ClassifiedParagraph:
    """A paragraph classification with formatting metadata."""

    index: int
    text: str
    paragraph_type: ParagraphType
    confidence: float
    font_info: Dict[str, Any]
    alignment: str
    should_fix: bool
    classification_reason: str


class ParagraphClassifier:
    """
    Rule-first paragraph classifier.

    LLM classification is not used in the main path. The goal is to prevent
    repeated runs from producing different correction targets.
    """

    SKIP_TYPES = {
        ParagraphType.EMPTY,
        ParagraphType.AUTHOR_INFO,
        ParagraphType.ABSTRACT_LABEL,
        ParagraphType.KEYWORDS_LABEL,
        ParagraphType.ALGORITHM,
        ParagraphType.CODE_BLOCK,
        ParagraphType.EQUATION,
        ParagraphType.TEMPLATE_INSTRUCTION,
    }

    FIX_TYPES = {
        ParagraphType.PAPER_TITLE,
        ParagraphType.JOURNAL_HEADER,
        ParagraphType.SECTION_HEADING,
        ParagraphType.BODY,
        ParagraphType.ABSTRACT_CONTENT,
        ParagraphType.KEYWORDS_CONTENT,
        ParagraphType.CAPTION,
        ParagraphType.REFERENCE,
    }

    DEFAULT_CLASSIFICATION_PATTERNS = {
        "journal_header": [
            r"journal\s+of",
            r"web\s+engineering$",
            r"vol\.\s*\d+",
            r"volume\s*\d+",
            r"eissn|pissn|issn",
            r"doi[:\s]",
            r"https?://",
            r"www\.",
        ],
        "caption": [
            r"^(figure|fig\.?|table|chart|diagram|image)\s*\d+",
        ],
        "reference": [
            r"^\[\d+\]",
        ],
        "author_info": [
            r"orcid",
            r"corresponding\s+author",
            r"university|universiti|faculty|fakulti|department|school|college|institute|centre|center",
        ],
    }

    DEFAULT_SECTION_TERMS = [
        "introduction",
        "background",
        "literature review",
        "related work",
        "methodology",
        "method",
        "methods",
        "materials and methods",
        "research methodology",
        "results",
        "discussion",
        "results and discussion",
        "results and discussions",
        "implementation",
        "evaluation",
        "experiment",
        "experiments",
        "experimental setup",
        "test configuration",
        "conclusion",
        "conclusions",
        "conclusion and future work",
        "future work",
        "acknowledgement",
        "acknowledgements",
        "acknowledgment",
        "acknowledgments",
        "funding statement",
        "author contributions",
        "conflict of interests",
        "conflict of interest",
        "ethics statements",
        "ethics statement",
        "references",
        "bibliography",
        "biographies of authors",
    ]

    def __init__(self, llm_integration=None, profile: Dict[str, Any] = None):
        self.llm = llm_integration
        self.profile = profile or {}
        self.classification_patterns = self._load_classification_patterns(self.profile)
        self.section_terms = self._load_section_terms(self.profile)
        self.classifications: List[ClassifiedParagraph] = []

    def _load_classification_patterns(self, profile: Dict[str, Any]) -> Dict[str, List[str]]:
        """Merge profile classification patterns with safe defaults."""
        patterns = {
            name: list(values)
            for name, values in self.DEFAULT_CLASSIFICATION_PATTERNS.items()
        }
        profile_patterns = profile.get("classification_patterns", {}) if profile else {}
        if isinstance(profile_patterns, dict):
            for name, values in profile_patterns.items():
                if not isinstance(values, list):
                    continue
                existing = patterns.setdefault(name, [])
                for value in values:
                    if isinstance(value, str) and value not in existing:
                        existing.append(value)
        return patterns

    def _load_section_terms(self, profile: Dict[str, Any]) -> List[str]:
        """Merge profile heading terms with safe defaults."""
        terms = list(self.DEFAULT_SECTION_TERMS)
        profile_terms = profile.get("heading_patterns", []) if profile else []
        if isinstance(profile_terms, list):
            for term in profile_terms:
                if isinstance(term, str):
                    normalized = re.sub(r"\s+", " ", term.lower().strip())
                    if normalized and normalized not in terms:
                        terms.append(normalized)
        return terms

    def _matches_pattern_group(self, group_name: str, text: str) -> bool:
        """Return True when text matches any profile/default pattern group."""
        for pattern in self.classification_patterns.get(group_name, []):
            try:
                if re.search(pattern, text):
                    return True
            except re.error:
                continue
        return False

    def classify_document(self, document) -> List[ClassifiedParagraph]:
        """Classify all paragraphs in a document."""
        para_data = []
        for index, paragraph in enumerate(document.paragraphs):
            para_data.append(
                {
                    "index": index,
                    "text": get_paragraph_text(paragraph),
                    "font_info": get_paragraph_font_info(paragraph),
                    "alignment": get_paragraph_alignment(paragraph),
                }
            )

        self.classifications = self._rule_classify_all(para_data)
        return self.classifications

    def _rule_classify_all(self, para_data: List[Dict]) -> List[ClassifiedParagraph]:
        """Classify every paragraph with deterministic rules."""
        classifications = []
        context = {
            "found_title": False,
            "in_abstract": False,
            "in_keywords": False,
            "in_references": False,
            "past_author_block": False,
        }

        for i, data in enumerate(para_data):
            # One paragraph of lookahead (skipping a single blank line,
            # to allow for the blank-line-after-title spacing rule) so
            # _is_probable_title can recognise the title by its
            # *position* relative to what follows - a following
            # author-name-like line - rather than only by the title
            # paragraph's own current formatting, which is exactly what
            # may still be wrong on a manuscript that hasn't been fixed
            # yet.
            next_text = ""
            for j in range(i + 1, min(i + 3, len(para_data))):
                candidate = para_data[j]["text"].strip()
                if candidate:
                    next_text = candidate
                    break
            context["next_text"] = next_text
            classification = self._rule_classify_single(data, context)
            classifications.append(classification)
            self._update_context(context, classification)

        return classifications

    def _fallback_classify_all(self, para_data: List[Dict]) -> List[ClassifiedParagraph]:
        """Backward-compatible alias for older code."""
        return self._rule_classify_all(para_data)

    def _fallback_classify_single(self, data: Dict, context: Dict = None) -> ClassifiedParagraph:
        """Backward-compatible alias for older code."""
        return self._rule_classify_single(data, context)

    def _update_context(self, context: Dict, classification: ClassifiedParagraph) -> None:
        """Update section context after each paragraph."""
        paragraph_type = classification.paragraph_type
        text_lower = classification.text.lower()

        if paragraph_type == ParagraphType.PAPER_TITLE:
            context["found_title"] = True

        if paragraph_type == ParagraphType.ABSTRACT_LABEL:
            context["in_abstract"] = True
            context["in_keywords"] = False
            context["past_author_block"] = True

        if paragraph_type == ParagraphType.ABSTRACT_CONTENT:
            context["in_abstract"] = True
            context["past_author_block"] = True

        if paragraph_type == ParagraphType.KEYWORDS_LABEL:
            context["in_abstract"] = False
            context["in_keywords"] = True
            context["past_author_block"] = True

        if paragraph_type == ParagraphType.KEYWORDS_CONTENT:
            context["in_abstract"] = False
            context["in_keywords"] = False
            context["past_author_block"] = True

        if paragraph_type == ParagraphType.SECTION_HEADING:
            context["in_abstract"] = False
            context["in_keywords"] = False
            context["past_author_block"] = True
            if "reference" in text_lower or "bibliography" in text_lower:
                context["in_references"] = True
            elif context.get("in_references"):
                context["in_references"] = False

    def _rule_classify_single(self, data: Dict, context: Dict = None) -> ClassifiedParagraph:
        """Classify one paragraph with deterministic rules."""
        if context is None:
            context = {
                "found_title": False,
                "in_abstract": False,
                "in_keywords": False,
                "in_references": False,
                "past_author_block": False,
            }

        text = data["text"].strip()
        index = data["index"]
        font_info = data["font_info"]
        alignment = data["alignment"]
        text_lower = text.lower()

        if not text:
            return self._create_classification(data, ParagraphType.EMPTY, 1.0, "Empty paragraph")

        if self._is_template_instruction(text_lower):
            return self._create_classification(data, ParagraphType.TEMPLATE_INSTRUCTION, 0.95, "Template instruction")

        if self._is_journal_header(text_lower, index):
            return self._create_classification(data, ParagraphType.JOURNAL_HEADER, 0.95, "Journal header")

        if self._is_caption(text_lower):
            return self._create_classification(data, ParagraphType.CAPTION, 0.90, "Caption")

        if self._is_probable_title(text, font_info, alignment, index, context):
            return self._create_classification(data, ParagraphType.PAPER_TITLE, 0.92, "Paper title")

        # A long paper title can be stored as two or more separate
        # paragraphs by Word/PDF-to-DOCX conversion. Once the first
        # title paragraph has been detected, a following centered
        # paragraph before the author block is treated as a title
        # continuation so it receives the same formatting.
        if self._is_title_continuation(text, font_info, alignment, index, context):
            return self._create_classification(
                data,
                ParagraphType.PAPER_TITLE,
                0.88,
                "Paper title continuation"
            )

        if text_lower in {"abstract", "abstract:"}:
            return self._create_classification(data, ParagraphType.ABSTRACT_LABEL, 0.96, "Abstract label")

        if self._starts_with_label(text_lower, "abstract"):
            return self._create_classification(data, ParagraphType.ABSTRACT_CONTENT, 0.92, "Abstract content")

        if text_lower in {"keywords", "keywords:", "key words", "key words:"}:
            return self._create_classification(data, ParagraphType.KEYWORDS_LABEL, 0.96, "Keywords label")

        if self._starts_with_keywords(text_lower):
            return self._create_classification(data, ParagraphType.KEYWORDS_CONTENT, 0.92, "Keywords content")

        if context.get("in_references") and self._is_reference_section_terminator(text_lower):
            return self._create_classification(data, ParagraphType.SECTION_HEADING, 0.90, "Post-reference section heading")

        if self._is_reference_entry(text, context.get("in_references", False)):
            return self._create_classification(data, ParagraphType.REFERENCE, 0.86, "Reference entry")

        if self._is_section_heading(text):
            return self._create_classification(data, ParagraphType.SECTION_HEADING, 0.92, "Section heading")

        if self._is_author_info(text, text_lower, index, alignment, context):
            return self._create_classification(data, ParagraphType.AUTHOR_INFO, 0.90, "Author information")

        if self._is_algorithm(text, text_lower):
            return self._create_classification(data, ParagraphType.ALGORITHM, 0.88, "Algorithm or pseudocode")

        if self._is_equation(text):
            return self._create_classification(data, ParagraphType.EQUATION, 0.84, "Equation or formula")

        if self._is_code_block(text, font_info):
            return self._create_classification(data, ParagraphType.CODE_BLOCK, 0.82, "Code-like text")

        if context.get("in_abstract"):
            return self._create_classification(data, ParagraphType.ABSTRACT_CONTENT, 0.76, "Abstract continuation")

        if context.get("in_keywords"):
            return self._create_classification(data, ParagraphType.KEYWORDS_CONTENT, 0.74, "Keywords continuation")

        if len(text) > 30:
            return self._create_classification(data, ParagraphType.BODY, 0.66, "Body text")

        return self._create_classification(data, ParagraphType.UNKNOWN, 0.50, "Unknown short text")

    def _is_journal_header(self, text_lower: str, index: int) -> bool:
        """Return True for journal names and publication metadata."""
        if index > 8:
            return False
        if self._starts_with_label(text_lower, "abstract") or self._starts_with_keywords(text_lower):
            return False

        return self._matches_pattern_group("journal_header", text_lower)

    def _is_author_info(self, text: str, text_lower: str, index: int, alignment: str, context: Dict) -> bool:
        """Return True for author names, affiliations, emails, and ORCID lines."""
        if index > 25:
            return False

        if context.get("past_author_block"):
            return False

        if re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text):
            return True

        if self._matches_pattern_group("author_info", text_lower):
            return True

        if context.get("found_title") and not self._starts_with_label(text_lower, "abstract"):
            if index <= 12 and not text.isupper() and len(text) <= 180:
                if re.search(r"\d", text) and len(text.split()) >= 1:
                    return True
                if "," in text and len(text.split()) <= 20:
                    return True
                if text_lower in {"&", "and"}:
                    return True
                if alignment == "CENTER" and len(text.split()) <= 8:
                    return True

        return False

    def _is_title_continuation(
        self,
        text: str,
        font_info: Dict,
        alignment: str,
        index: int,
        context: Dict
    ) -> bool:
        """Return True for a paragraph that continues a multi-line paper title."""
        if not context.get("found_title"):
            return False

        # Title continuation must remain in the front-matter area.
        if index > 12:
            return False

        text = (text or "").strip()
        if not text or len(text) < 2 or len(text) > 240:
            return False

        text_lower = text.lower()

        # A continuation cannot itself be another structural element.
        if self._is_journal_header(text_lower, index):
            return False
        if self._is_section_heading(text, allow_custom_all_caps=False):
            return False
        if self._starts_with_label(text_lower, "abstract"):
            return False
        if self._starts_with_keywords(text_lower):
            return False
        if self._is_caption(text_lower):
            return False

        # Strong author-information signals stop title continuation.
        # Do not call _is_author_info() here because centered title lines
        # can contain commas as normal title punctuation.
        if re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text):
            return False
        if re.search(
            r"\b(university|universiti|faculty|department|school|college|institute|orcid|corresponding author)\b",
            text_lower
        ):
            return False
        if re.search(r"\d", text):
            return False

        # A title continuation is normally centered. This prevents a
        # normal body paragraph from being mistaken for a title.
        if alignment != "CENTER":
            return False

        # Very short centered text is more likely to be an author line.
        if len(text.split()) < 2:
            return False

        # If a large font is already present, it is strong evidence of
        # a title continuation. Otherwise centered text in the title
        # area is still accepted.
        font_size = font_info.get("font_size") or 0
        if font_size >= 12:
            return True

        return index <= 5

    def _is_probable_title(self, text: str, font_info: Dict, alignment: str, index: int, context: Dict) -> bool:
        """Return True for the actual paper title, not the journal title."""
        if context.get("found_title") or index > 12:
            return False

        text_lower = text.lower()
        if self._is_journal_header(text_lower, index):
            return False
        if self._is_section_heading(text, allow_custom_all_caps=False):
            return False
        if self._starts_with_label(text_lower, "abstract") or self._starts_with_keywords(text_lower):
            return False
        if len(text) < 15 or len(text) > 240:
            return False

        font_size = font_info.get("font_size") or 0
        is_bold = bool(font_info.get("bold"))
        is_centered = alignment == "CENTER"

        if font_size >= 16 and is_centered:
            return True
        if font_size >= 14 and is_bold and is_centered:
            return True
        if index <= 6 and is_centered and len(text.split()) >= 4 and not text.endswith("."):
            return True

        # Formatting-independent fallback: an early, title-length,
        # non-sentence-like paragraph immediately followed by something
        # that reads like an author-name line is almost certainly the
        # paper title even when it isn't (yet) centered/bold/large -
        # exactly the state a not-yet-fixed manuscript's title is likely
        # to be in, which the checks above alone would never catch,
        # leaving the title (and everything positioned relative to it)
        # misclassified for the rest of the document.
        if index <= 4 and len(text.split()) >= 4 and not text.endswith(".") and self._looks_like_author_line(context.get("next_text", "")):
            return True

        return False

    def _looks_like_author_line(self, text: str) -> bool:
        """Return True for a short line that reads like a list of
        author names (with or without superscript affiliation markers),
        used only as a positional signal for recognising the paper
        title above it - not a claim about the line's own formatting.
        """
        if not text or len(text) > 200 or text.endswith("."):
            return False
        has_multiple_words = len(text.split()) >= 2
        has_separator = ("," in text) or (" and " in text.lower()) or ("&" in text)
        has_name_like_caps = bool(re.search(r"[A-Z][a-z]+\s+[A-Z][a-z]+", text))
        return has_multiple_words and has_name_like_caps and (has_separator or len(text.split()) <= 6)

    def _is_section_heading(self, text: str, allow_custom_all_caps: bool = True) -> bool:
        """Return True for common academic section headings."""
        text_clean = re.sub(r"\s+", " ", text.strip())
        # Strip apostrophes/curly quotes before matching - "Authors'
        # Biographies" (curly \u2019) must match a pattern list entry
        # written as "authors biographies" (no apostrophe at all), and a
        # bare `in section_terms` set-membership check is exact-match
        # only, so the apostrophe difference alone would otherwise make
        # an already-listed heading fail to match.
        text_lower = re.sub(r"[\u2019\u2018']", "", text_clean.lower()).strip(".:")

        section_terms = set(self.section_terms)

        if text_lower in section_terms:
            return True

        heading_words = "|".join(re.escape(term) for term in sorted(section_terms, key=len, reverse=True))
        if re.match(rf"^(\d+(\.\d+)*\.?|[ivxlc]+\.?)\s*({heading_words})$", text_lower):
            return True

        if re.match(r"^\d+\.\d+(\.\d+)*\.?\s*\S", text_lower) and len(text_clean) <= 100:
            return True

        for term in section_terms:
            if text_lower.startswith(f"{term} ") and re.search(r"\d+\.\d+", text_lower):
                return True

        if text_clean.isupper() and 4 <= len(text_clean) <= 80:
            if any(word in text_lower for word in section_terms):
                return True

        if allow_custom_all_caps and self._is_custom_all_caps_section_heading(text_clean, text_lower):
            return True

        return False

    def _is_custom_all_caps_section_heading(self, text_clean: str, text_lower: str) -> bool:
        """Return True for short custom section headings that are not profile terms."""
        if not text_clean.isupper():
            return False
        if not 10 <= len(text_clean) <= 100:
            return False
        if text_clean.endswith((".", ",", ";", ":")):
            return False
        if re.search(r"[@:/\\]|https?|www\.|doi|orcid|eissn|issn", text_lower):
            return False
        if re.search(r"\b(figure|fig|table|vol|no|pp|pages?)\b", text_lower):
            return False

        words = re.findall(r"[A-Z][A-Z0-9'-]*", text_clean)
        if not 2 <= len(words) <= 12:
            return False
        if any(len(word) > 28 for word in words):
            return False

        return True

    def _is_template_instruction(self, text_lower: str) -> bool:
        """Return True for template-only formatting instructions."""
        return bool(re.match(r"^\(?\s*\d+(?:\.\d+)?\s*[-\s]*(?:font size|point|pt)", text_lower))

    def _is_caption(self, text_lower: str) -> bool:
        """Return True for figure and table captions.

        A genuine caption ("Table 1", "Table 1: Comparison of X and Y")
        is short - the pattern itself only requires the text to *start*
        with "Table N" / "Figure N", with nothing capping how much more
        text can follow, so an ordinary body sentence that happens to
        start the same way ("Table 1 displays a cross-journal breakdown
        of...") matches identically to a real caption. Capped at a
        length no real caption line reaches, so a full descriptive
        sentence doesn't classify as one.
        """
        if len(text_lower) > 150:
            return False
        return self._matches_pattern_group("caption", text_lower)

    def _is_reference_entry(self, text: str, in_references: bool) -> bool:
        """Return True for bibliography entries."""
        stripped = text.strip()
        if self._matches_pattern_group("reference", stripped.lower()):
            return True
        if in_references and stripped.lower() not in {"references", "bibliography", "works cited"}:
            return True
        return False

    def _is_reference_section_terminator(self, text_lower: str) -> bool:
        """Return True for headings that end the references section."""
        normalized = re.sub(r"\s+", " ", text_lower.strip())
        return (
            normalized.startswith("appendix")
            or normalized.startswith("biograph")
            or normalized in {"acknowledgement", "acknowledgements"}
        )

    def _is_algorithm(self, text: str, text_lower: str) -> bool:
        """Return True for algorithm and pseudocode lines."""
        if re.match(r"^algorithm\s*\d*", text_lower):
            return True
        if re.match(r"^(input|output|begin|end|procedure|function)\s*:", text_lower):
            return True
        if re.match(r"^\d+\s*:\s+", text_lower):
            return True
        pseudocode_markers = [
            "for each ",
            "while ",
            "if ",
            "return ",
            "initialize ",
        ]
        return any(text_lower.startswith(marker) for marker in pseudocode_markers) and len(text.split()) <= 25

    def _is_equation(self, text: str) -> bool:
        """Return True for standalone equations or formula-like lines."""
        stripped = text.strip()
        if len(stripped) > 120 or len(stripped.split()) > 20:
            return False
        operator_count = sum(stripped.count(token) for token in ["=", "+", "-", "*", "/", "^", "_"])
        return operator_count >= 2 and bool(re.search(r"[A-Za-z0-9]\s*=", stripped))

    def _is_code_block(self, text: str, font_info: Dict) -> bool:
        """Return True for monospace code-like paragraphs."""
        font_name = (font_info.get("font_name") or "").lower()
        if any(name in font_name for name in ["courier", "consolas", "monaco"]):
            return True
        if re.search(r"[{};]|==|!=|<=|>=", text) and len(text.split()) <= 30:
            return True
        return False

    def _starts_with_label(self, text_lower: str, label: str) -> bool:
        """Return True when text begins with a named label."""
        return bool(re.match(rf"^{label}\s*[-\u2013\u2014:]", text_lower))

    def _starts_with_keywords(self, text_lower: str) -> bool:
        """Return True for keyword lines."""
        return bool(re.match(r"^(keywords?|key\s+words?)\s*[-\u2013\u2014:]", text_lower))

    def _create_classification(
        self,
        data: Dict,
        para_type: ParagraphType,
        confidence: float,
        reason: str,
    ) -> ClassifiedParagraph:
        """Create a classification record."""
        return ClassifiedParagraph(
            index=data["index"],
            text=data["text"],
            paragraph_type=para_type,
            confidence=confidence,
            font_info=data["font_info"],
            alignment=data["alignment"],
            should_fix=para_type in self.FIX_TYPES,
            classification_reason=reason,
        )

    def _string_to_paragraph_type(self, type_str: str) -> ParagraphType:
        """Convert a string label to a ParagraphType enum."""
        if not type_str:
            return ParagraphType.UNKNOWN

        type_mapping = {
            "journal_header": ParagraphType.JOURNAL_HEADER,
            "paper_title": ParagraphType.PAPER_TITLE,
            "author_info": ParagraphType.AUTHOR_INFO,
            "abstract_label": ParagraphType.ABSTRACT_LABEL,
            "abstract_content": ParagraphType.ABSTRACT_CONTENT,
            "keywords_label": ParagraphType.KEYWORDS_LABEL,
            "keywords_content": ParagraphType.KEYWORDS_CONTENT,
            "section_heading": ParagraphType.SECTION_HEADING,
            "body": ParagraphType.BODY,
            "caption": ParagraphType.CAPTION,
            "reference": ParagraphType.REFERENCE,
            "algorithm": ParagraphType.ALGORITHM,
            "code_block": ParagraphType.CODE_BLOCK,
            "equation": ParagraphType.EQUATION,
        }
        return type_mapping.get(type_str.lower().strip(), ParagraphType.UNKNOWN)

    def get_paragraphs_to_fix(self) -> List[ClassifiedParagraph]:
        """Return paragraphs that may be auto-fixed."""
        return [cp for cp in self.classifications if cp.should_fix]

    def get_paragraphs_to_skip(self) -> List[ClassifiedParagraph]:
        """Return paragraphs that should be preserved."""
        return [cp for cp in self.classifications if not cp.should_fix]

    def get_classification_summary(self) -> Dict[str, int]:
        """Return counts by paragraph type."""
        summary = {}
        for cp in self.classifications:
            type_name = cp.paragraph_type.value
            summary[type_name] = summary.get(type_name, 0) + 1
        return summary
