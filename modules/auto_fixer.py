"""
Auto-Fixer Module
Automatically fixes formatting issues while preserving special formatting
"""

from docx import Document
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docx.shared import Pt, Inches, Emu
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from copy import deepcopy

from collections import Counter
import re
from io import BytesIO

from .utils import (
    load_document, get_paragraph_text, get_paragraph_alignment, truncate_text,
    is_font_equivalent, get_run_font_info, get_sdt_reference_paragraphs,
    classify_author_info_role, get_space_after_pt, get_space_before_pt,
    get_hanging_indent_inches, to_journal_caption_title_case,
    set_space_after_pt, set_space_before_pt, has_lines_based_spacing,
    contains_east_asian_text
)
from .paragraph_classifier import (
    ParagraphType, ClassifiedParagraph, ParagraphClassifier
)
from config import ALIGNMENT_REVERSE_MAP


# Characters of right-to-left / Indic / Southeast-Asian complex scripts (Hebrew,
# Arabic, Devanagari ... Sinhala, Thai, Lao) - runs containing these keep their
# own complex-script (w:cs) font when the run's Latin font is corrected.
_COMPLEX_SCRIPT_RE = re.compile(r"[\u0590-\u08FF\u0900-\u0DFF\u0E00-\u0EFF]")


@dataclass
class ChangeRecord:
    """Record of a change made during auto-fix"""
    paragraph_index: int
    location: str
    change_type: str
    before: str
    after: str
    text_preview: str
    property_name: str = "formatting"
    current_value: str = ""
    target_value: str = ""
    paragraph_type: str = ""
    evidence: str = ""


@dataclass
class PostFixValidationResult:
    """Result of checking the corrected document after auto-fix."""
    is_safe: bool
    before_issues: int
    after_issues: int
    before_score: float
    after_score: float
    issue_delta: int
    score_delta: float
    new_or_increased_categories: Dict[str, int]
    message: str


def _category_issue_counts(result: Any) -> Dict[str, int]:
    """Return issue counts per category from a checker result."""
    issues_by_category = getattr(result, "issues_by_category", {}) or {}
    return {
        category: len(issues)
        for category, issues in issues_by_category.items()
        if issues
    }


def validate_post_fix_result(before_result: Any, after_result: Any) -> PostFixValidationResult:
    """Compare pre-fix and post-fix checker results."""
    before_issues = int(getattr(before_result, "total_issues", 0) or 0)
    after_issues = int(getattr(after_result, "total_issues", 0) or 0)
    before_score = float(getattr(before_result, "compliance_score", 0.0) or 0.0)
    after_score = float(getattr(after_result, "compliance_score", 0.0) or 0.0)
    issue_delta = after_issues - before_issues
    score_delta = round(after_score - before_score, 2)

    before_counts = _category_issue_counts(before_result)
    after_counts = _category_issue_counts(after_result)
    increased_categories = {
        category: count - before_counts.get(category, 0)
        for category, count in after_counts.items()
        if count > before_counts.get(category, 0)
    }

    score_dropped = after_score < before_score
    is_safe = issue_delta <= 0 and not score_dropped
    if is_safe:
        message = "Post-fix validation did not increase detected issues."
    elif issue_delta > 0:
        message = "Post-fix validation found more issues after auto-fix. Review the corrected document before using it."
    else:
        message = "Post-fix validation found a lower compliance index after auto-fix. Review the corrected document before using it."

    return PostFixValidationResult(
        is_safe=is_safe,
        before_issues=before_issues,
        after_issues=after_issues,
        before_score=before_score,
        after_score=after_score,
        issue_delta=issue_delta,
        score_delta=score_delta,
        new_or_increased_categories=increased_categories,
        message=message,
    )


def validate_fixed_document(rules: Dict[str, Any], fixed_doc_bytes: bytes):
    """Run the checker again on a corrected document."""
    from .manuscript_checker import ManuscriptChecker

    checker = ManuscriptChecker(rules, None)
    checker.load_manuscript(BytesIO(fixed_doc_bytes))
    return checker.check_all()


def summarize_remaining_issues(result: Any) -> List[Dict[str, Any]]:
    """Create a compact table summary of issues remaining after auto-fix."""
    issues_by_category = getattr(result, "issues_by_category", {}) or {}
    rows = []
    for category, issues in issues_by_category.items():
        if not issues:
            continue
        first_issue = issues[0]
        rows.append({
            "Category": category.replace("_", " ").title(),
            "Count": len(issues),
            "First Location": getattr(first_issue, "location", ""),
            "First Issue": getattr(first_issue, "description", ""),
        })
    return sorted(rows, key=lambda row: (-row["Count"], row["Category"]))


# ---------------------------------------------------------------------------
# Reference publication-source italic detection
# ---------------------------------------------------------------------------
# APA7 (the JCLC reference style) italicizes the periodical/journal name and
# its volume number - e.g. "...study. *Frontiers in Psychology*, 13,
# 1053310." The checker already flags when this is missing (see
# manuscript_checker.py's publication_italic_required handling), but nothing
# ever fixed it - this pair of patterns identifies exactly that span so it
# can be italicized without touching the rest of the entry. Only the common
# "Journal Name, Volume(Issue)" shape is targeted; anything that doesn't
# match (books, edited chapters, conference proceedings without this exact
# shape) is left alone rather than guessed at.
_REFERENCE_YEAR_RE = re.compile(r"\(\d{4}[a-z]?(?:,\s*\d{4}[a-z]?)*\)\.\s*")
_REFERENCE_SOURCE_RE = re.compile(
    r"(?P<source>[A-Za-z][^.,][^.]{2,}?),\s*"
    r"(?P<volume>\d{1,4})"
    r"(?P<issue>\(\d+(?:[/\u2013\-]\d+)?[a-zA-Z]?\))?"
    r"(?=,\s*(?:[A-Za-z]+\s+)?[\w\u2013\-]*\d|\s*[.,]?\s*$)"
)


class AutoFixer:
    """
    Automatically fixes formatting issues in a manuscript
    while preserving special formatting (italic, underline, subscript, etc.)
    """
    
    def __init__(
        self,
        rules: Dict[str, Any],
        classifications: List[ClassifiedParagraph],
        issues_by_category: Optional[Dict[str, List[Any]]] = None,
    ):
        """
        Initialize auto-fixer with template rules and paragraph classifications
        
        Args:
            rules: Template formatting rules
            classifications: List of classified paragraphs from ParagraphClassifier
        """
        self.rules = rules
        self.classifications = classifications
        self.issues_by_category = issues_by_category or {}
        self.changes: List[ChangeRecord] = []
        self.document = None
        self.original_document = None
        self._classification_map: Dict[int, ClassifiedParagraph] = {}
        self._issue_map: Dict[int, set] = {}
        self._issue_property_map: Dict[int, set] = {}
        self._global_issue_categories: set = set()
        self._global_issue_property_map: Dict[str, set] = {}
        self._has_explicit_issues = bool(issues_by_category)
        self._consistent_heading_space_before: Optional[float] = None
        self._build_issue_map()
    
    def load_manuscript(self, file_path_or_bytes):
        """Load the manuscript to fix"""
        self.document = load_document(file_path_or_bytes)
        self.original_document = deepcopy(self.document)
        # Create index map for quick lookup
        self._classification_map = {cp.index: cp for cp in self.classifications}
        return self

    def _build_issue_map(self):
        """Build paragraph-to-category and paragraph-to-property maps from detected issues."""
        self._issue_map = {}
        self._issue_property_map = {}
        self._global_issue_categories = set()
        self._global_issue_property_map = {}
        for category, issues in self.issues_by_category.items():
            for issue in issues:
                para_index = getattr(issue, "paragraph_index", -1)
                if para_index is None or para_index < 0:
                    self._global_issue_categories.add(category)
                    properties = self._infer_global_issue_properties(category, issue)
                    if properties:
                        self._global_issue_property_map.setdefault(category, set()).update(properties)
                    continue
                self._issue_map.setdefault(para_index, set()).add(category)
                properties = self._infer_issue_properties(issue)
                if properties:
                    self._issue_property_map.setdefault(para_index, set()).update(properties)

    def _infer_issue_properties(self, issue: Any) -> set:
        """Infer one or more affected formatting properties from a checker issue."""
        property_name = self._infer_issue_property(issue)
        if property_name:
            return {property_name}

        description = (getattr(issue, "description", "") or "").lower()
        current_value = (getattr(issue, "current_value", "") or "").lower()
        if "formatting does not match template" not in description:
            return set()

        properties = set()
        if "font " in current_value:
            properties.add("font_name")
        if "size " in current_value:
            properties.add("font_size")
        if "bold" in current_value:
            properties.add("bold")
        if "italic" in current_value:
            properties.add("italic")
        if "alignment" in current_value:
            properties.add("alignment")
        return properties

    def _infer_issue_property(self, issue: Any) -> Optional[str]:
        """Infer the affected formatting property from a checker issue."""
        description = (getattr(issue, "description", "") or "").lower()
        location = (getattr(issue, "location", "") or "").lower()

        if "number" in description or "heading number" in location:
            if "bold" in description:
                return "number_bold"
            if "font size" in description or "size" in description:
                return "number_font_size"
            if "font" in description:
                return "number_font_name"

        if "line spacing" in description:
            return "line_spacing"
        if "spacing before" in description or "before spacing" in description:
            return "space_before"
        if "spacing after" in description or "after spacing" in description:
            return "space_after"
        if "hanging indent" in description:
            return "hanging_indent"
        if "manual tab" in description or "manual tabs" in description:
            return "manual_tabs"
        if "capitalization" in description or "all capital" in description:
            return "capitalization"
        if "alignment" in description:
            return "alignment"
        if "list item" in description and "indent" in description:
            return "list_item_indent"
        if "font size" in description or " size " in f" {description} ":
            return "font_size"
        if "font" in description:
            return "font_name"
        if "bold" in description:
            return "bold"
        if "italic" in description:
            return "italic"
        return None

    def _infer_global_issue_properties(self, category: str, issue: Any) -> set:
        """Infer properties represented by an aggregate category-level issue."""
        properties = self._infer_issue_properties(issue)
        if properties:
            return properties

        description = (getattr(issue, "description", "") or "").lower()
        if category == "body_text" and "body text formatting issues" in description:
            return {"font_name", "font_size", "bold"}
        if category == "line_spacing":
            return {"line_spacing"}
        if category == "other" and "manual tab" in (getattr(issue, "description", "") or "").lower():
            return {"manual_tabs"}
        return set()

    def _should_fix_category(self, index: int, categories: List[str]) -> bool:
        """Return True when a paragraph has a detected issue for the target categories."""
        if not self._has_explicit_issues:
            return True
        found_categories = self._issue_map.get(index, set())
        return any(
            category in found_categories or category in self._global_issue_categories
            for category in categories
        )

    def _allowed_properties_for(
        self,
        index: int,
        categories: Optional[List[str]] = None,
        fallback: Optional[List[str]] = None,
    ) -> Optional[set]:
        """Return properties that may be fixed for a paragraph."""
        if not self._has_explicit_issues:
            return None
        properties = set(self._issue_property_map.get(index, set()))
        for category in categories or []:
            properties.update(self._global_issue_property_map.get(category, set()))
        if properties:
            return properties
        return set()

    def _property_allowed(self, property_name: str, allowed_properties: Optional[set]) -> bool:
        """Return True when a property can be changed under the current issue filter."""
        return allowed_properties is None or property_name in allowed_properties

    def _has_category_issue(self, category: str) -> bool:
        """Return True when the checker reported at least one issue in a category."""
        if not self._has_explicit_issues:
            return True
        return bool(self.issues_by_category.get(category))

    def _category_allows_global_property(self, category: str, property_name: str) -> bool:
        """Return True when a category-level issue permits a document-level fix."""
        if not self._has_explicit_issues:
            return True
        return property_name in self._global_issue_property_map.get(category, set())
    
    def _compute_consistent_heading_space_before(self) -> Optional[float]:
        """Find the dominant space-before value used across section headings.

        The template itself often doesn't specify an explicit space-before
        for headings (it's inherited from further up the style chain, with
        nothing concrete to extract), but a real manuscript's own headings
        are still supposed to all look the same. When pdf2docx introduces
        a specific heading that deviates from what every other similar
        heading in the same document uses (e.g. "References" sitting at
        roughly half the gap "Acknowledgements"/"Funding"/etc. use), that
        is the actual, visible problem - normalizing to this document's
        own dominant value directly fixes it without inventing an
        arbitrary universal number the template never specified.

        Restricted to non-numbered headings (skips "1.0 Introduction",
        "4.1.2 Determiners", etc.) - those numbered section/subsection
        headings form their own, separate, considerably messier cluster
        of values in practice, and folding them into the same vote as the
        declaration-style headings (Acknowledgements, Funding, References,
        Author Biography, ...) - which really are consistent with each
        other - would just dilute the vote with an unrelated convention.
        """
        values = []
        for index, classification in self._classification_map.items():
            if classification.paragraph_type != ParagraphType.SECTION_HEADING:
                continue
            if re.match(r"^\s*\d", classification.text or ""):
                continue
            if index >= len(self.document.paragraphs):
                continue
            space_before = self.document.paragraphs[index].paragraph_format.space_before
            if space_before is not None:
                values.append(round(space_before.pt, 1))
        if not values:
            return None
        return Counter(values).most_common(1)[0][0]

    def fix_all(self) -> Tuple[Document, List[ChangeRecord]]:
        """
        Apply all formatting fixes to the document
        
        Returns:
            Tuple of (fixed document, list of changes made)
        """
        if not self.document:
            raise ValueError("No document loaded. Call load_manuscript() first.")
        
        self.changes = []

        if self._has_category_issue("layout"):
            self._fix_layout()
        
        # Fix margins first only when margin issues were detected.
        if self._has_category_issue("margins"):
            self._fix_margins()
        if self._category_allows_global_property("other", "manual_tabs"):
            self._fix_page_header_layout()

        self._consistent_heading_space_before = self._compute_consistent_heading_space_before()
        self._author_block_after = self._compute_author_block_after()

        # Fix paragraphs based on their classification
        for i, para in enumerate(self.document.paragraphs):
            classification = self._classification_map.get(i)
            
            if not classification:
                continue
            if (
                not classification.should_fix
                and classification.paragraph_type not in (
                    ParagraphType.AUTHOR_INFO,
                    ParagraphType.ABSTRACT_LABEL,
                )
            ):
                continue
            
            para_type = classification.paragraph_type
            
            if para_type == ParagraphType.JOURNAL_HEADER and self._is_journal_title_header(classification.text):
                if not self._should_fix_category(i, ["journal_header"]):
                    continue
                self._fix_journal_header(para, i)
            elif para_type == ParagraphType.JOURNAL_HEADER:
                # A JOURNAL_HEADER-classified paragraph that isn't
                # actually the journal title/issue line itself (e.g. the
                # "Corresponding author: ..." line) falls through here -
                # it still gets the same standard front-matter spacing
                # normalization, just not the journal-title-specific font
                # treatment from _fix_journal_header above.
                self._fix_front_matter_spacing(para, i)
            elif para_type == ParagraphType.PAPER_TITLE:
                # A multi-line paper title may be represented by multiple
                # paragraphs. The checker intentionally reports the title
                # as one logical item, so only the first paragraph may have
                # a title issue attached to its index. Continuation paragraphs
                # must still receive the same title formatting.
                #
                # Always called now (not gated behind _should_fix_category)
                # - the space_before/after/line_spacing fix inside
                # _fix_title is itself unconditional, while font/alignment
                # still respect their own allowed_properties gating
                # (force=True for continuation paragraphs, as before).
                is_title_continuation = (
                    "continuation"
                    in (classification.classification_reason or "").lower()
                )
                self._fix_title(para, i, force=is_title_continuation)
            elif para_type == ParagraphType.AUTHOR_INFO:
                # Always called (not gated behind _should_fix_category) -
                # the spacing normalization inside _fix_author_info is
                # itself unconditional (always correct regardless of what
                # the checker specifically flagged), while its other
                # properties (font, bold, ...) still respect their own
                # allowed_properties gating internally.
                self._fix_author_info(para, i)
            elif para_type == ParagraphType.SECTION_HEADING:
                # Always called - the spacing fix inside _fix_heading is
                # itself unconditional now, while font/bold/etc. still
                # respect their own allowed_properties gating.
                self._fix_heading(para, i)
            elif para_type == ParagraphType.ABSTRACT_LABEL:
                # The "Abstract" label paragraph is a heading-like element
                # (never previously dispatched to any fixer at all - it
                # simply fell through unhandled), so it's fixed the same
                # way a section heading is, using its own dedicated
                # spacing rule from _heading_rules_for_text.
                self._fix_heading(para, i)
            elif para_type == ParagraphType.BODY:
                # Always called - the core per-run font checks inside
                # _fix_body_text are themselves unconditional (see the
                # comment there for why: the checker's own "most common
                # value per paragraph" detection can miss a minority-
                # mismatched run), so gating the call itself on the
                # checker having flagged something for this paragraph
                # would prevent the function from ever running at all for
                # exactly the paragraphs that need it.
                self._fix_body_text(para, i)
            elif para_type == ParagraphType.ABSTRACT_CONTENT:
                # Always called, for the same reason as AUTHOR_INFO above
                # - the space_before/outline_level fix inside _fix_abstract
                # is unconditional, while font/line_spacing/space_after
                # still respect their own allowed_properties gating.
                self._fix_abstract(para, i)
            elif para_type == ParagraphType.KEYWORDS_CONTENT:
                # Always called - the spacing fix inside _fix_keywords is
                # itself unconditional, while its other properties still
                # respect their own allowed_properties gating.
                self._fix_keywords(para, i)
            elif para_type == ParagraphType.CAPTION:
                # Always called - the core per-run font checks and
                # space_before/after/line_spacing inside _fix_caption
                # are themselves unconditional (same reasoning as
                # _fix_body_text above).
                self._fix_caption(para, i)
            elif para_type == ParagraphType.REFERENCE:
                # Always called - the core per-run font checks inside
                # _fix_reference are themselves unconditional (same
                # reasoning as _fix_body_text above).
                self._fix_reference(para, i)

        self._fix_unclassified_text_font()

        if self._has_category_issue("references"):
            self._fix_sdt_references()

        self._fix_table_content()
        self._fix_caption_title_spacing()
        self._remove_caption_trailing_periods()
        self._fix_three_line_tables()
        self._remove_blank_lines_before_figures_and_tables()
        self._remove_blank_lines_within_author_block()
        self._ensure_blank_lines_between_front_matter_and_headings()

        return self.document, self.changes

    def _insert_blank_paragraph_after(self, paragraph):
        """Insert one empty paragraph immediately after the given
        paragraph, matching the template's own convention of a blank
        line before each major structural break, and return it."""
        new_p_element = OxmlElement("w:p")
        paragraph._p.addnext(new_p_element)
        return Paragraph(new_p_element, paragraph._parent)

    def _fix_unclassified_text_font(self):
        """Font name and size only, for short lines the classifier could
        not place anywhere (ParagraphType.UNKNOWN).

        The classifier only calls a paragraph "body" once it is long
        enough to read as prose; a short hand-typed item ("a. Semantic
        Errors", "Aim 5: Explore", "Note. Values are means.") comes back
        UNKNOWN, and the main dispatch loop skips UNKNOWN entirely - so
        such a line used to keep whatever font and size the author had
        left on it (e.g. Arial 10) in a document that is otherwise
        Times New Roman 11. Only font and size are normalised here: what
        the line *is* (a sub-item, a table note, a stray fragment) is
        unknown, so its bold/italic, indentation, alignment and spacing
        are left exactly as the author had them.
        """
        body_rules = self.rules.get("body", {})
        expected_font = body_rules.get("font_name", "Times New Roman")
        expected_size = body_rules.get("font_size", 11)
        for i, para in enumerate(self.document.paragraphs):
            classification = self._classification_map.get(i)
            if not classification or classification.paragraph_type != ParagraphType.UNKNOWN:
                continue
            if not para.text.strip():
                continue
            details = []
            for run in para.runs:
                if run.text.strip():
                    details.extend(self._fix_run_formatting(
                        run, expected_font, expected_size, None, allowed_properties=None,
                    ))
            if details:
                self._add_property_changes(
                    paragraph_index=i,
                    location="Unclassified short text",
                    change_type="body",
                    details=details,
                    text_preview=truncate_text(para.text, 50),
                    paragraph_type=ParagraphType.UNKNOWN.value,
                )

    # A caption label line: "Table 1", "Figure 2", "Table 3. Title" /
    # "Table 3: Title" - but NOT a sentence that merely starts with one
    # ("Table 3 shows the results..."), where the number is followed by
    # a space and a word rather than by nothing or by "."/":".
    _CAPTION_LABEL_RE = re.compile(r"^(table|figure|fig\.?)\s*\d+\s*(?:[.:]\s*(?:\S.*)?)?$", re.IGNORECASE)
    _LABEL_ONLY_RE = re.compile(r"^(table|figure|fig\.?)\s*\d+\s*[.:]?$", re.IGNORECASE)

    def _fix_caption_title_spacing(self):
        """Give the title line under a "Table N" / "Figure N" label line
        the caption spacing (space before 0, after 8pt, single line).

        In this template a caption is two paragraphs: the label ("Table
        1") and, directly under it, the title ("List of Relevant
        Attributes ..."). Only the label is classified as a caption; the
        title line is classified as ordinary body text, so it only picks
        up body spacing - and only if the checker happened to flag it,
        which is why some titles came out with spacing and others with
        none. It is recognised here by position instead (the non-empty
        short line directly under a label line), so the rule applies to
        every caption title regardless of what the checker flagged.
        Nothing else about the line (italic, alignment, ...) is changed.
        """
        caption_rules = self.rules.get("caption", {})
        expected_before = caption_rules.get("space_before", 0)
        expected_after = caption_rules.get("space_after", 8)

        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        type_by_index = {c.index: c.paragraph_type for c in classifier.classify_document(self.document)}

        for i in range(len(paragraphs) - 1):
            if type_by_index.get(i) != ParagraphType.CAPTION:
                continue
            if not self._LABEL_ONLY_RE.match(paragraphs[i].text.strip()):
                continue
            title_p = paragraphs[i + 1]
            title_text = title_p.text.strip()
            if not title_text or len(title_text) > 250:
                continue
            if type_by_index.get(i + 1) in (ParagraphType.CAPTION, ParagraphType.SECTION_HEADING):
                continue
            if self._is_plain_empty_paragraph(title_p):
                continue

            pf = title_p.paragraph_format
            problems = []
            current_before = get_space_before_pt(title_p)
            current_after = get_space_after_pt(title_p)
            if current_before is None or abs(float(current_before) - float(expected_before)) > 0.05 or has_lines_based_spacing(title_p):
                set_space_before_pt(title_p, float(expected_before))
                problems.append(("space_before", f"{current_before}pt" if current_before is not None else "(inherited)", f"{expected_before}pt"))
            if current_after is None or abs(float(current_after) - float(expected_after)) > 0.05 or has_lines_based_spacing(title_p):
                set_space_after_pt(title_p, float(expected_after))
                problems.append(("space_after", f"{current_after}pt" if current_after is not None else "(inherited)", f"{expected_after}pt"))
            if pf.line_spacing_rule != WD_LINE_SPACING.SINGLE or (pf.line_spacing not in (None, 1.0)):
                before_ls = f"{pf.line_spacing_rule} ({pf.line_spacing})"
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                problems.append(("line_spacing", before_ls, "single (1.0)"))

            for prop, cur, tgt in problems:
                self._add_change_record(
                    paragraph_index=i + 1,
                    location=f"Caption title: {truncate_text(title_text, 30)}",
                    change_type="caption",
                    property_name=prop,
                    current_value=cur,
                    target_value=tgt,
                    text_preview=truncate_text(title_text, 50),
                    paragraph_type=ParagraphType.CAPTION.value,
                    evidence="Caption title line uses space before 0, space after 8pt, single line spacing",
                )

    def _fix_three_line_tables(self):
        """JCLC tables: three horizontal lines only (top, under the header row, bottom)."""
        from .table_borders import normalise_table
        for n, tbl in enumerate(self.document.element.body.findall(qn("w:tbl")), 1):
            if normalise_table(tbl):
                self._add_change_record(
                    paragraph_index=-1, location=f"Table {n}", change_type="table",
                    property_name="table_borders", current_value="extra / missing lines",
                    target_value="three horizontal lines (top, under header, bottom)",
                    text_preview="", paragraph_type="table", evidence="JCLC three-line table")

    def _remove_caption_trailing_periods(self):
        """JCLC: a table / figure caption must not end with a full stop.

        The caption is either one paragraph ("Table 1. Title.") or a label line ("Table 1") followed by a
        title line; the final full stop of the last caption line is removed. An ellipsis ("...") is kept.
        """
        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        type_by_index = {c.index: c.paragraph_type for c in classifier.classify_document(self.document)}
        full_re = re.compile(r"^(?:table|figure|fig\.?)\s*\d+\s*[.:]\s*\S", re.IGNORECASE)

        for i, para in enumerate(paragraphs):
            if type_by_index.get(i) != ParagraphType.CAPTION:
                continue
            text = para.text.strip()
            target = None
            if self._LABEL_ONLY_RE.match(text):
                j = i + 1
                while j < len(paragraphs) and self._is_plain_empty_paragraph(paragraphs[j]):
                    j += 1
                if j < len(paragraphs):
                    cand = paragraphs[j]
                    ctext = cand.text.strip()
                    if ctext and len(ctext) <= 250 and not self._LABEL_ONLY_RE.match(ctext) \
                            and type_by_index.get(j) != ParagraphType.SECTION_HEADING:
                        target, tindex = cand, j
            elif full_re.match(text) and len(text) <= 250:
                target, tindex = para, i
            if target is None:
                continue
            before = target.text.strip()
            if not re.search(r"(?<!\.)\.\s*$", before):
                continue
            text_nodes = [t for t in target._p.iter(qn("w:t")) if t.getparent().tag == qn("w:r")]
            for t in reversed(text_nodes):
                value = t.text or ""
                if not value.strip():
                    continue
                stripped = value.rstrip()
                if stripped.endswith("."):
                    t.text = stripped[:-1].rstrip() if stripped[:-1].strip() else ""
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    # a lone "." run: also trim the space left at the end of the previous text
                    if not (t.text or "").strip():
                        idx = text_nodes.index(t)
                        for prev in reversed(text_nodes[:idx]):
                            if (prev.text or "").strip():
                                prev.text = prev.text.rstrip()
                                break
                break
            self._add_change_record(
                paragraph_index=tindex,
                location=f"Caption: {truncate_text(before, 30)}",
                change_type="caption",
                property_name="trailing_full_stop",
                current_value="ends with a full stop",
                target_value="no full stop at the end of the caption",
                text_preview=truncate_text(before, 50),
                paragraph_type=ParagraphType.CAPTION.value,
                evidence="JCLC table / figure captions do not end with a full stop",
            )

    def _remove_blank_lines_before_figures_and_tables(self):
        """No blank line before a figure or a table.

        Removes the plain-empty paragraphs directly above the start of a
        figure/table block: a "Table N"/"Figure N" caption label line, a
        table, or a picture-only paragraph. Authors are inconsistent
        about leaving an empty paragraph above these, so it is made
        uniform; paragraphs that merely *look* empty but carry a section
        break, picture, page break etc. are never removed (see
        _is_plain_empty_paragraph). Nothing below a figure/table is
        touched.
        """
        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        type_by_index = {c.index: c.paragraph_type for c in classifier.classify_document(self.document)}
        index_by_element = {p._p: n for n, p in enumerate(paragraphs)}

        block_starts = []  # (element, label)
        body = self.document.element.body
        for child in list(body):
            if child.tag == qn("w:tbl"):
                table_text = " ".join(t.text or "" for t in child.iter(qn("w:t"))).lower()
                # A journal masthead/branding table is not a data table.
                if "doi:" in table_text or "mmu" in table_text or "eissn" in table_text:
                    continue
                block_starts.append((child, "Table"))
            elif child.tag == qn("w:p") and child in index_by_element:
                n = index_by_element[child]
                para = paragraphs[n]
                text = para.text.strip()
                if type_by_index.get(n) == ParagraphType.CAPTION and self._CAPTION_LABEL_RE.match(text):
                    block_starts.append((child, truncate_text(text, 20)))
                elif not text and any(True for _ in child.iter(qn("w:drawing"))):
                    block_starts.append((child, "Figure"))

        for element, label in block_starts:
            removed = 0
            prev = element.getprevious()
            while prev is not None and prev.tag == qn("w:p") and self._is_plain_empty_paragraph(Paragraph(prev, self.document._body)):
                to_remove = prev
                prev = prev.getprevious()
                to_remove.getparent().remove(to_remove)
                removed += 1
            if not removed:
                continue
            anchor = element
            while anchor is not None and not (anchor.tag == qn("w:p") and anchor in index_by_element):
                anchor = anchor.getprevious()
            anchor_index = index_by_element.get(anchor, 0) if anchor is not None else 0
            self._add_change_record(
                paragraph_index=anchor_index,
                location=f"Before {label}",
                change_type="blank_line",
                property_name="blank_line",
                current_value=f"{removed} blank line(s) before figure/table",
                target_value="no blank line before figure/table",
                text_preview=label,
                paragraph_type=ParagraphType.CAPTION.value,
                evidence="Template has no blank line before a figure or table",
            )

    _CORRESPONDING_RE = re.compile(r"corresponding\s+author", re.IGNORECASE)

    def _find_author_block(self, paragraphs, type_for_index):
        """Indices of the author block that follows the paper title: the
        author-names line, the affiliation line(s) and the corresponding-
        author line, i.e. the run of non-empty paragraphs directly after
        the title that are author information. Found by position, so an
        affiliation continuation line with no institution keyword
        ("Selangor, Malaysia") is still part of it. A "Corresponding
        author:" line counts even when the classifier filed it under
        another type (it sometimes comes back as JOURNAL_HEADER).
        """
        title_idx = next(
            (i for i in range(len(paragraphs)) if type_for_index(i) == ParagraphType.PAPER_TITLE),
            None,
        )
        if title_idx is None:
            return []
        block = []
        for i in range(title_idx + 1, len(paragraphs)):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            kind = type_for_index(i)
            if not block and kind == ParagraphType.PAPER_TITLE:
                continue  # a title that wraps onto a second paragraph
            if kind == ParagraphType.AUTHOR_INFO or self._CORRESPONDING_RE.search(text):
                block.append(i)
            else:
                break
        return block

    def _compute_author_block_after(self):
        """Space after (pt) for each paragraph of the author block.

        The author-names line and the last line of the block (normally
        the corresponding-author line) get 8pt; the lines in between -
        the affiliations - get 0, so the affiliations sit under the
        author names' 8pt gap and run straight into the corresponding-
        author line with no space between them. Decided by position
        rather than by role, so it also holds when a line has no role
        keyword, and when a manuscript has no corresponding-author line
        the last affiliation line takes the closing 8pt instead.
        """
        paragraphs = list(self.document.paragraphs)
        block = self._find_author_block(
            paragraphs,
            lambda i: getattr(self._classification_map.get(i), "paragraph_type", None),
        )
        after = {}
        for pos, idx in enumerate(block):
            in_the_middle = 0 < pos < len(block) - 1
            after[idx] = 0.0 if in_the_middle else 8.0
        return after

    def _remove_blank_lines_within_author_block(self):
        """No blank line between the author names, the affiliations and
        the corresponding-author line.

        Removes plain-empty paragraphs sitting between the first and the
        last line of the author block (see _is_plain_empty_paragraph for
        what is never removed). The blank lines before the block (after
        the title) and after it (before Abstract) are handled separately
        by _ensure_blank_lines_between_front_matter_and_headings.
        """
        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        type_by_index = {c.index: c.paragraph_type for c in classifier.classify_document(self.document)}
        block = self._find_author_block(paragraphs, lambda i: type_by_index.get(i))
        if len(block) < 2:
            return
        removed = 0
        for j in range(block[-1] - 1, block[0], -1):
            if self._is_plain_empty_paragraph(paragraphs[j]):
                paragraphs[j]._element.getparent().remove(paragraphs[j]._element)
                removed += 1
        if removed:
            self._add_change_record(
                paragraph_index=block[0],
                location="Author block",
                change_type="blank_line",
                property_name="blank_line",
                current_value=f"{removed} blank line(s) between author names, affiliation and corresponding author",
                target_value="no blank line between author names, affiliation and corresponding author",
                text_preview=truncate_text(paragraphs[block[0]].text, 50),
                paragraph_type=ParagraphType.AUTHOR_INFO.value,
                evidence="Template has no blank line between the author names, the affiliation and the corresponding-author line",
            )

    @staticmethod
    def _is_plain_empty_paragraph(paragraph) -> bool:
        """True only for a paragraph that is empty in every sense, i.e.
        safe to delete: no text, no section break, and nothing but empty
        runs inside it.

        A paragraph with no visible text is *not* necessarily deletable -
        an "empty" paragraph can be what carries a section break
        (<w:sectPr>), an inline/anchored picture (<w:drawing>), a page or
        column break (<w:br>), a tab, a bookmark or a field. Deleting one
        of those because its text is empty would silently destroy the
        section/picture/page break. Anything beyond properties and empty
        text therefore makes the paragraph non-deletable.
        """
        if paragraph.text.strip():
            return False
        p = paragraph._p
        p_pr = p.find(qn("w:pPr"))
        if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
            return False
        for child in p:
            if child.tag == qn("w:pPr"):
                continue
            if child.tag == qn("w:r"):
                for sub in child:
                    if sub.tag == qn("w:rPr"):
                        continue
                    if sub.tag == qn("w:t") and not (sub.text or "").strip():
                        continue
                    return False
                continue
            return False
        return True

    def _ensure_blank_lines_between_front_matter_and_headings(self):
        """Ensure exactly one blank paragraph appears: after the paper
        title (before the author names), after the corresponding-author
        line (before Abstract), and before every main section heading;
        subheadings get none (any blank paragraphs above one are
        removed) - matching the template's own spacing
        convention, which a font/spacing rule on any single paragraph
        can't express since it's about whether a *separate* blank
        paragraph exists between two others.

        Run last, after every other structural fix, since inserting a
        paragraph shifts every later paragraph's index and would
        invalidate the classification-index lookups every fix above
        this one depends on.
        """
        paragraphs = list(self.document.paragraphs)

        def is_blank(p):
            return not p.text.strip()

        # 1) Paper title -> blank line -> author names.
        for i, para in enumerate(paragraphs):
            classification = self._classification_map.get(i)
            if classification and classification.paragraph_type == ParagraphType.PAPER_TITLE:
                if i + 1 < len(paragraphs) and not is_blank(paragraphs[i + 1]):
                    self._insert_blank_paragraph_after(para)
                    self._add_change_record(
                        paragraph_index=i,
                        location="Title",
                        change_type="blank_line",
                        property_name="blank_line",
                        current_value="no blank line after title",
                        target_value="blank line after title",
                        text_preview=truncate_text(para.text, 50),
                        paragraph_type=ParagraphType.PAPER_TITLE.value,
                        evidence="Template requires a blank line between the paper title and author names",
                    )
                break

        # 2) Corresponding-author line -> blank line -> Abstract.
        # Re-classified fresh (not self._classification_map) since the
        # insertion in step 1 above shifts every later paragraph's
        # index by one, which would otherwise point this lookup at the
        # wrong paragraph.
        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        classifications_after_step1 = classifier.classify_document(self.document)
        type_by_index_step2 = {c.index: c.paragraph_type for c in classifications_after_step1}
        for i, para in enumerate(paragraphs):
            if type_by_index_step2.get(i) != ParagraphType.ABSTRACT_LABEL:
                continue
            if i > 0 and not is_blank(paragraphs[i - 1]):
                self._insert_blank_paragraph_after(paragraphs[i - 1])
                self._add_change_record(
                    paragraph_index=i,
                    location="Abstract",
                    change_type="blank_line",
                    property_name="blank_line",
                    current_value="no blank line before Abstract",
                    target_value="blank line before Abstract",
                    text_preview=truncate_text(para.text, 50),
                    paragraph_type=ParagraphType.ABSTRACT_LABEL.value,
                    evidence="Template requires a blank line between author/affiliation info and Abstract",
                )
            elif i > 1:
                # Exactly one blank line, not "at least one": collapse any
                # extra plain-empty paragraphs directly above the one
                # already there.
                extra = []
                j = i - 2
                while j >= 0 and self._is_plain_empty_paragraph(paragraphs[j]):
                    extra.append(paragraphs[j])
                    j -= 1
                for blank_p in extra:
                    blank_p._element.getparent().remove(blank_p._element)
                if extra:
                    self._add_change_record(
                        paragraph_index=i,
                        location="Abstract",
                        change_type="blank_line",
                        property_name="blank_line",
                        current_value=f"{len(extra) + 1} consecutive blank lines before Abstract",
                        target_value="one blank line before Abstract",
                        text_preview=truncate_text(para.text, 50),
                        paragraph_type=ParagraphType.ABSTRACT_LABEL.value,
                        evidence="Template uses exactly one blank line before Abstract",
                    )
            break

        # 3) Exactly one blank line before every main section heading,
        # and *no* blank line before any subheading (see the subheading
        # branch below). Not applied to the very first paragraph.
        # Re-classified fresh here rather than using
        # self._classification_map: the insertions in steps 1-2 above
        # shift every later paragraph's index, which would otherwise
        # make that map's indices point at the wrong paragraph for
        # everything from the Abstract onward.
        paragraphs = list(self.document.paragraphs)
        classifier = ParagraphClassifier(profile=self.rules.get("_profile", {}))
        fresh_classifications = classifier.classify_document(self.document)
        type_by_index = {c.index: c.paragraph_type for c in fresh_classifications}
        for i in range(len(paragraphs) - 1, 0, -1):
            para_type = type_by_index.get(i)
            if para_type != ParagraphType.SECTION_HEADING:
                continue
            text = paragraphs[i].text.strip()
            if not text:
                continue
            # _heading_rules_for_text has three return paths: the main
            # heading rule object itself (numbered "N.0" headings), a
            # *copy* of it (declaration headings like "Funding", and
            # "Abstract"), or the subheading rule object (numbered
            # "N.M" headings, M != 0). A blank line goes before the
            # first two (both are main headings) but not the third, so
            # this checks specifically for "is this the subheading
            # rule" rather than "is this the exact heading rule object"
            # - the latter would wrongly treat the declaration-heading
            # copy as a subheading purely because dict-copying gives it
            # a different identity than self.rules.get("heading", {}).
            heading_rules = self._heading_rules_for_text(text)
            subheading_rules = self.rules.get("subheading")
            if subheading_rules is not None and heading_rules is subheading_rules:
                # A subheading (e.g. "3.1 Data Collection") has *no* blank
                # line before it. Authors are inconsistent about this -
                # some leave an empty paragraph above a subheading, some
                # don't - so every plain-empty paragraph directly above it
                # is removed (see _is_plain_empty_paragraph for what is
                # deliberately never removed).
                removed = 0
                j = i - 1
                while j >= 0 and self._is_plain_empty_paragraph(paragraphs[j]):
                    paragraphs[j]._element.getparent().remove(paragraphs[j]._element)
                    removed += 1
                    j -= 1
                if removed:
                    self._add_change_record(
                        paragraph_index=i,
                        location=f"Subheading: {truncate_text(text, 30)}",
                        change_type="blank_line",
                        property_name="blank_line",
                        current_value=f"{removed} blank line(s) before subheading",
                        target_value="no blank line before subheading",
                        text_preview=truncate_text(text, 50),
                        paragraph_type=ParagraphType.SECTION_HEADING.value,
                        evidence="Template has no blank line before a subheading",
                    )
                continue
            if not is_blank(paragraphs[i - 1]):
                self._insert_blank_paragraph_after(paragraphs[i - 1])
                self._add_change_record(
                    paragraph_index=i,
                    location=f"Heading: {truncate_text(text, 30)}",
                    change_type="blank_line",
                    property_name="blank_line",
                    current_value="no blank line before heading",
                    target_value="blank line before heading",
                    text_preview=truncate_text(text, 50),
                    paragraph_type=ParagraphType.SECTION_HEADING.value,
                    evidence="Template requires a blank line before each main section heading",
                )
            else:
                # Exactly one blank line, not "at least one" - collapse
                # any extra consecutive blank paragraphs immediately
                # before the heading down to the single one already
                # confirmed present.
                extra_blanks = []
                j = i - 2
                while j >= 0 and self._is_plain_empty_paragraph(paragraphs[j]):
                    extra_blanks.append(paragraphs[j])
                    j -= 1
                if extra_blanks:
                    for blank_p in extra_blanks:
                        blank_p._element.getparent().remove(blank_p._element)
                    self._add_change_record(
                        paragraph_index=i,
                        location=f"Heading: {truncate_text(text, 30)}",
                        change_type="blank_line",
                        property_name="blank_line",
                        current_value=f"{len(extra_blanks) + 1} consecutive blank lines before heading",
                        target_value="one blank line before heading",
                        text_preview=truncate_text(text, 50),
                        paragraph_type=ParagraphType.SECTION_HEADING.value,
                        evidence="Template uses exactly one blank line before each main section heading",
                    )

    @staticmethod
    def _first_row_is_column_header(table) -> bool:
        """True if the first row of the table looks like a row of column
        headers rather than content.

        A header row needs a table to have headers *for*: at least two
        rows and two columns (a 1x1 or single-row table is a text box /
        callout / banner, not a data table). Its cells are short labels,
        so any cell over 80 characters means the row is prose, and a row
        made up mostly of numbers is data (a continuation of a table that
        was split across pages, for instance), not labels. Plain years
        ("2019", "2020") do count as labels - they are common column
        headers.
        """
        rows = table.rows
        if len(rows) < 2 or len(table.columns) < 2:
            return False
        seen, texts = set(), []
        for cell in rows[0].cells:
            if id(cell._tc) in seen:   # a merged cell shows up once per column it spans
                continue
            seen.add(id(cell._tc))
            text = cell.text.strip()
            if text:
                texts.append(text)
        if not texts:
            return False
        if any(len(text) > 80 for text in texts):
            return False
        def numeric_data(text):
            if re.fullmatch(r"(19|20)\d{2}", text):
                return False
            return bool(re.fullmatch(r"[\d.,%+\-()\s]+", text)) and bool(re.search(r"\d", text))
        if sum(1 for text in texts if numeric_data(text)) * 2 > len(texts):
            return False
        return True

    def _fix_table_content(self):
        """Fix the font of every data table's own cell content (not its
        caption, which _fix_caption already handles separately).

        Always applied unconditionally, for the same reason as the other
        core font fixes throughout this class: table cell font has no
        legitimate reason to vary from the template's own convention,
        and there was previously no check or fix for it at all anywhere
        in the pipeline.
        """
        table_rules = self.rules.get("table_content", {})
        expected_font = table_rules.get("font_name")
        expected_size = table_rules.get("font_size")
        if not expected_font and not expected_size:
            return

        for table in self.document.tables:
            table_text = " ".join(
                cell.text for row in table.rows for cell in row.cells
            ).lower()
            # A journal masthead/branding table (logo, DOI, copyright,
            # issue/ISSN line) carries its own distinct, intentional
            # formatting - not the data-table convention this fix is for.
            if "doi:" in table_text or "mmu" in table_text or "eissn" in table_text:
                continue

            first_row_is_header = self._first_row_is_column_header(table)
            for row_index, row in enumerate(table.rows):
                # The first row of a data table is its column headers -
                # bold, unlike the rest of the table's ordinary data rows,
                # which are left at whatever bold state they already
                # have (only font/size is normalized for them). Only when
                # the first row really looks like a header row, though
                # (see _first_row_is_column_header): a table used as a
                # text box or callout holds ordinary prose in its first
                # row, and bolding that bolds the whole box.
                row_bold = True if (row_index == 0 and first_row_is_header) else None
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if not run.text.strip():
                                continue
                            self._fix_run_formatting(
                                run,
                                expected_font,
                                expected_size,
                                row_bold,
                                allowed_properties=None,
                            )

    def _is_journal_title_header(self, text: str) -> bool:
        """Return True for journal title lines, excluding volume and ISSN metadata."""
        text_lower = text.lower()
        return "journal of" in text_lower or "web engineering" in text_lower
    
    def _fix_margins(self):
        """Fix page margins with structured change records."""
        expected_margins = self.rules.get("margins", {})

        for section in self.document.sections:
            margin_specs = [
                ("left_margin", "left", "Left Margin"),
                ("right_margin", "right", "Right Margin"),
                ("top_margin", "top", "Top Margin"),
                ("bottom_margin", "bottom", "Bottom Margin"),
            ]

            for attr_name, rule_name, label in margin_specs:
                current_margin = getattr(section, attr_name)
                if not current_margin:
                    continue

                current_value = current_margin.inches
                expected_value = expected_margins.get(rule_name, 1.0)
                if abs(current_value - expected_value) <= 0.05:
                    continue

                setattr(section, attr_name, Inches(expected_value))
                self._add_change_record(
                    paragraph_index=-1,
                    location=label,
                    change_type="margins",
                    property_name=rule_name,
                    current_value=f"{current_value:.2f} in",
                    target_value=f"{expected_value:.2f} in",
                    text_preview="Document margins adjusted",
                    paragraph_type="document",
                    evidence="Detected margin issue"
                )

    def _fix_layout(self):
        """Fix page size, orientation, and column count."""
        rules = self.rules.get("layout", {})
        page_sizes = {
            "Letter": (8.5, 11.0),
            "A4": (8.27, 11.69),
        }
        expected_page_size = rules.get("page_size")
        expected_orientation = rules.get("orientation")
        expected_columns = rules.get("columns")

        for section_index, section in enumerate(self.document.sections):
            if expected_page_size in page_sizes:
                width, height = page_sizes[expected_page_size]
                if expected_orientation == "LANDSCAPE":
                    width, height = height, width
                section.page_width = Inches(width)
                section.page_height = Inches(height)

            if expected_orientation:
                section.orientation = (
                    WD_ORIENT.LANDSCAPE
                    if expected_orientation == "LANDSCAPE"
                    else WD_ORIENT.PORTRAIT
                )

            if expected_columns is not None:
                cols = section._sectPr.find(qn("w:cols"))
                if cols is None:
                    cols = OxmlElement("w:cols")
                    section._sectPr.append(cols)
                cols.set(qn("w:num"), str(int(expected_columns)))

        self._add_change_record(
            paragraph_index=-1,
            location="Page Layout",
            change_type="layout",
            text_preview="",
            property_name="layout",
            current_value="Non-matching layout",
            target_value=f"{expected_page_size}, {expected_orientation}, {expected_columns} column(s)",
            paragraph_type="document",
            evidence="Detected page layout did not match target rules",
        )

    def _fix_page_header_layout(self):
        """Normalize page header tab spacing to reduce Word wrapping."""
        for section_index, section in enumerate(self.document.sections):
            headers = [
                ("Page Header", section.header),
                ("First Page Header", section.first_page_header),
                ("Even Page Header", section.even_page_header),
            ]
            usable_width = section.page_width - section.left_margin - section.right_margin

            for label, header in headers:
                for paragraph in header.paragraphs:
                    original_text = paragraph.text
                    if not self._has_unstable_header_tabs(original_text):
                        continue
                    normalized_text = self._normalize_tabbed_header_text(original_text)
                    if normalized_text == original_text:
                        continue

                    self._replace_paragraph_text_preserving_first_run(paragraph, normalized_text)
                    try:
                        tab_stops = paragraph.paragraph_format.tab_stops
                        tab_stops.clear_all()
                        tab_stops.add_tab_stop(usable_width, WD_TAB_ALIGNMENT.RIGHT)
                    except Exception:
                        pass

                    self._add_change_record(
                        paragraph_index=-1,
                        location=f"{label} (Section {section_index + 1})",
                        change_type="page_header",
                        property_name="manual_tabs",
                        current_value=truncate_text(original_text, 80),
                        target_value=truncate_text(normalized_text, 80),
                        text_preview=truncate_text(normalized_text, 80),
                        paragraph_type="document_header",
                        evidence="Detected page header manual tabs/spaces that may wrap in Word",
                    )

    def _has_unstable_header_tabs(self, text: str) -> bool:
        """Return True for tab spacing patterns that are likely to wrap in Word."""
        return "\t\t" in text or bool(re.search(r"\t\s{2,}", text))

    def _normalize_tabbed_header_text(self, text: str) -> str:
        """Collapse unstable tab runs into one left/right tab-separated header line."""
        if "\t" not in text:
            return text
        parts = [re.sub(r"\s+", " ", part).strip() for part in text.split("\t") if part.strip()]
        if len(parts) < 2:
            return text
        return f"{parts[0]}\t{parts[-1]}"

    def _trim_paragraph_edge_whitespace(self, paragraph):
        """Trim paragraph edge spaces without replacing runs and losing formatting."""
        text = paragraph.text
        leading_count = len(text) - len(text.lstrip(" \t"))
        trailing_count = len(text) - len(text.rstrip(" \t"))

        remaining = leading_count
        for run in paragraph.runs:
            if remaining <= 0:
                break
            run_text = run.text
            if len(run_text) <= remaining:
                run.text = ""
                remaining -= len(run_text)
            else:
                run.text = run_text[remaining:]
                remaining = 0

        remaining = trailing_count
        for run in reversed(paragraph.runs):
            if remaining <= 0:
                break
            run_text = run.text
            if len(run_text) <= remaining:
                run.text = ""
                remaining -= len(run_text)
            else:
                run.text = run_text[:-remaining]
                remaining = 0

    def _replace_paragraph_text_preserving_first_run(self, paragraph, text: str):
        """Replace paragraph text while preserving the first visible run formatting."""
        first_index = next(
            (index for index, run in enumerate(paragraph.runs) if run.text),
            None,
        )
        if first_index is None:
            paragraph.add_run(text)
            return

        first_run = paragraph.runs[first_index]
        first_run.text = text
        for index, run in enumerate(paragraph.runs):
            if index != first_index:
                run.text = ""

    def _add_change_record(
        self,
        paragraph_index: int,
        location: str,
        change_type: str,
        property_name: str,
        current_value: str,
        target_value: str,
        text_preview: str,
        paragraph_type: str,
        evidence: str,
    ):
        """Append one structured change record."""
        self.changes.append(ChangeRecord(
            paragraph_index=paragraph_index,
            location=location,
            change_type=change_type,
            before=f"{property_name}: {current_value}",
            after=f"{property_name}: {target_value}",
            text_preview=text_preview,
            property_name=property_name,
            current_value=current_value,
            target_value=target_value,
            paragraph_type=paragraph_type,
            evidence=evidence,
        ))

    def _add_property_changes(
        self,
        paragraph_index: int,
        location: str,
        change_type: str,
        details: List[Dict[str, str]],
        text_preview: str,
        paragraph_type: str,
    ):
        """Add de-duplicated property changes for one paragraph."""
        seen = set()
        for detail in details:
            key = (
                detail.get("property_name", "formatting"),
                detail.get("current_value", ""),
                detail.get("target_value", ""),
            )
            if key in seen:
                continue
            seen.add(key)
            self._add_change_record(
                paragraph_index=paragraph_index,
                location=location,
                change_type=change_type,
                property_name=detail.get("property_name", "formatting"),
                current_value=detail.get("current_value", ""),
                target_value=detail.get("target_value", ""),
                text_preview=text_preview,
                paragraph_type=paragraph_type,
                evidence=detail.get("evidence", "Detected formatting difference"),
            )

    def _fix_journal_header(self, paragraph, index: int):
        """Fix journal title/header formatting without using paper title rules."""
        header_rules = self.rules.get("journal_header", {})
        changes = []
        current_alignment = get_paragraph_alignment(paragraph)
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["journal_header"],
            fallback=["font_name", "font_size", "bold", "alignment", "manual_tabs"],
        )

        expected_font = header_rules.get("font_name", "Palatino Linotype")
        expected_size = header_rules.get("font_size", 24)
        expected_bold = header_rules.get("bold", True)
        expected_alignment = header_rules.get("alignment", "CENTER")

        original_text = paragraph.text
        cleaned_text = original_text.strip(" \t")
        if (
            self._property_allowed("manual_tabs", allowed_properties)
            and cleaned_text
            and cleaned_text != original_text
        ):
            self._trim_paragraph_edge_whitespace(paragraph)
            changes.append({
                "property_name": "manual_tabs",
                "current_value": "Manual tabs/spaces",
                "target_value": "No leading/trailing manual tabs/spaces",
                "evidence": "Journal header contained manual indentation that can shift layout",
            })

        if self._property_allowed("alignment", allowed_properties) and current_alignment != expected_alignment:
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 1)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Journal header alignment did not match target rule",
            })

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_strike=False,
                    allowed_properties=allowed_properties,
                ))

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Journal Header",
                change_type="journal_header",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 50),
                paragraph_type=ParagraphType.JOURNAL_HEADER.value,
            )

    def _fix_run_formatting(self, run, expected_font: str, expected_size: float,
                           expected_bold: Optional[bool] = None,
                           expected_italic: Optional[bool] = None,
                           expected_strike: Optional[bool] = None,
                           allowed_properties: Optional[set] = None) -> List[Dict[str, str]]:
        """Fix a run and return structured property changes."""
        changes = []
        font = run.font

        preserve_italic = font.italic
        preserve_underline = font.underline
        preserve_subscript = font.subscript
        preserve_superscript = font.superscript
        preserve_strike = font.strike

        current_font = font.name
        if expected_font is not None and self._property_allowed("font_name", allowed_properties):
            if current_font is None or not is_font_equivalent(current_font, expected_font):
                font.name = expected_font
                # font.name only sets w:rFonts' ascii/hAnsi attributes -
                # eastAsia/cs and a leftover hint="eastAsia" are untouched
                # by it, and pdf2docx (or an author's own East-Asian IME)
                # commonly leaves w:eastAsia set to a CJK font (e.g.
                # "Microsoft YaHei") with hint="eastAsia" on runs that are
                # actually plain Latin text. That hint tells the renderer
                # to prefer the eastAsia font for this run, so the text
                # visually renders in the CJK font's own letter-spacing
                # and glyph widths (looking like "wrong spacing" even
                # though ascii/hAnsi now correctly say Times New Roman)
                # rather than the Latin font this run's rPr otherwise
                # claims. Every rule in this system targets a Western
                # font, so eastAsia/cs are aligned to the same name and
                # the misleading hint is dropped whenever the font gets
                # corrected.
                #
                # Only when the run holds none of that script's own
                # characters, though: a run that really contains Chinese/
                # Japanese/Korean text keeps its East-Asian font (SimSun,
                # MS Mincho, ...) - Times New Roman has no CJK glyphs, so
                # overwriting it would make Word substitute an arbitrary
                # fallback and lose the author's chosen font - and a run
                # with Thai/Arabic/Hebrew/Indic text keeps its complex-
                # script font for the same reason.
                rpr = run._r.get_or_add_rPr()
                rfonts = rpr.find(qn("w:rFonts"))
                if rfonts is not None:
                    run_text = run.text or ""
                    has_cjk = contains_east_asian_text(run_text)
                    has_complex = bool(_COMPLEX_SCRIPT_RE.search(run_text))
                    if not has_cjk:
                        if rfonts.get(qn("w:eastAsia")) is not None:
                            rfonts.set(qn("w:eastAsia"), expected_font)
                        if rfonts.get(qn("w:hint")) == "eastAsia":
                            del rfonts.attrib[qn("w:hint")]
                    if not has_complex:
                        if rfonts.get(qn("w:cs")) is not None:
                            rfonts.set(qn("w:cs"), expected_font)
                changes.append({
                    "property_name": "font_name",
                    "current_value": current_font or "(inherited)",
                    "target_value": expected_font,
                    "evidence": "Run font did not match target rule",
                })

        current_size = font.size.pt if font.size else None
        if expected_size is not None and self._property_allowed("font_size", allowed_properties):
            if current_size is None or abs(current_size - expected_size) > 0.5:
                font.size = Pt(expected_size)
                changes.append({
                    "property_name": "font_size",
                    "current_value": f"{current_size} pt" if current_size is not None else "(inherited)",
                    "target_value": f"{expected_size} pt",
                    "evidence": "Run font size did not match target rule",
                })

        if expected_bold is not None and self._property_allowed("bold", allowed_properties):
            current_bold = font.bold
            current_bold_value = bool(current_bold) if current_bold is not None else False
            if current_bold_value != expected_bold:
                font.bold = expected_bold
                changes.append({
                    "property_name": "bold",
                    "current_value": "Bold" if current_bold_value else "Not Bold",
                    "target_value": "Bold" if expected_bold else "Not Bold",
                    "evidence": "Run bold setting did not match target rule",
                })

        if expected_italic is not None and self._property_allowed("italic", allowed_properties):
            current_italic = font.italic
            current_italic_value = bool(current_italic) if current_italic is not None else False
            if current_italic_value != expected_italic:
                font.italic = expected_italic
                changes.append({
                    "property_name": "italic",
                    "current_value": "Italic" if current_italic_value else "Not Italic",
                    "target_value": "Italic" if expected_italic else "Not Italic",
                    "evidence": "Run italic setting did not match target rule",
                })
        else:
            font.italic = preserve_italic

        if expected_strike is not None and self._property_allowed("strike", allowed_properties):
            if font.strike != expected_strike:
                font.strike = expected_strike
        else:
            font.strike = preserve_strike

        font.underline = preserve_underline
        font.subscript = preserve_subscript
        font.superscript = preserve_superscript

        return changes

    def _fix_title(self, paragraph, index: int, force: bool = False):
        """Fix paper title formatting with structured change records."""
        title_rules = self.rules.get("title", {})
        changes = []
        current_alignment = get_paragraph_alignment(paragraph)

        # A continuation paragraph is part of the same logical title even
        # when the checker did not attach a separate issue to its index.
        allowed_properties = (
            None
            if force
            else self._allowed_properties_for(
                index,
                categories=["title"],
                fallback=["font_name", "font_size", "bold", "italic", "alignment"],
            )
        )

        expected_font = title_rules.get("font_name", "Times New Roman")
        expected_size = title_rules.get("font_size", 24)
        expected_bold = title_rules.get("bold", None)
        expected_italic = title_rules.get("italic", None)
        expected_alignment = title_rules.get("alignment", "CENTER")

        if self._property_allowed("alignment", allowed_properties) and current_alignment != expected_alignment:
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 1)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Paragraph alignment did not match target rule",
            })

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_italic,
                    expected_strike=False,
                    allowed_properties=allowed_properties,
                ))

        # Applied unconditionally (not gated behind allowed_properties) -
        # these three target values are always correct for the title
        # regardless of what else was specifically flagged for it.
        expected_space_before = title_rules.get("space_before", 0)
        current_space_before = get_space_before_pt(paragraph)
        if (
            current_space_before is None
            or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_before_pt(paragraph, float(expected_space_before))
            changes.append({
                "property_name": "space_before",
                "current_value": f"{current_space_before}pt" if current_space_before is not None else "(inherited)",
                "target_value": f"{expected_space_before}pt",
                "evidence": "Title spacing before did not match target rule",
            })

        expected_space_after = title_rules.get("space_after", 8)
        current_space_after = get_space_after_pt(paragraph)
        if (
            current_space_after is None
            or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_after_pt(paragraph, float(expected_space_after))
            changes.append({
                "property_name": "space_after",
                "current_value": f"{current_space_after}pt" if current_space_after is not None else "(inherited)",
                "target_value": f"{expected_space_after}pt",
                "evidence": "Title spacing after did not match target rule",
            })

        expected_line_spacing = title_rules.get("line_spacing", 1.0)
        current_line_spacing = paragraph.paragraph_format.line_spacing
        current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
        effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
        needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
        if abs(float(expected_line_spacing) - 1.0) < 0.001 and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
            needs_fix = True
        if needs_fix:
            paragraph.paragraph_format.line_spacing = expected_line_spacing
            changes.append({
                "property_name": "line_spacing",
                "current_value": (
                    f"{current_line_spacing_rule} ({current_line_spacing})"
                    if current_line_spacing is not None
                    else str(current_line_spacing_rule)
                ),
                "target_value": str(expected_line_spacing),
                "evidence": "Title line spacing did not match target rule",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Paper Title",
                change_type="title",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 50),
                paragraph_type=ParagraphType.PAPER_TITLE.value,
            )

    def _fix_author_info(self, paragraph, index: int):
        """Fix front-matter author information without changing its text."""
        role = classify_author_info_role(get_paragraph_text(paragraph), self.rules)
        rules = self.rules.get(role, self.rules.get("author", {}))
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["author_info"],
            fallback=["font_name", "font_size", "bold", "italic", "alignment"],
        )
        changes = []
        expected_size = rules.get("font_size")
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            changes.extend(self._fix_run_formatting(
                run,
                rules.get("font_name"),
                expected_size,
                rules.get("bold"),
                rules.get("italic"),
                allowed_properties=allowed_properties,
            ))

            # An author-number/corresponding-author marker ("1", "2",
            # "1*", "*") attached to a name or affiliation is meant to be
            # true superscript - raised and visually smaller through that
            # formatting property - not a separately shrunk explicit font
            # size standing in for it. The latter is what was actually
            # happening here: the marker kept its own smaller font size
            # (e.g. 7pt next to 11pt body text) with no superscript flag
            # at all, which the general per-run fix above never corrects,
            # since the paragraph's overall font size already reads as
            # "matching" from its non-marker runs. Fixed independently of
            # that gate, since a marker needing this is unambiguous and
            # always correct regardless of what else was flagged.
            if re.fullmatch(r"\d{1,2}\*?|\*", run.text.strip()):
                marker_changed = False
                if run.font.superscript is not True:
                    run.font.superscript = True
                    marker_changed = True
                if expected_size and run.font.size != Pt(expected_size):
                    run.font.size = Pt(expected_size)
                    marker_changed = True
                if marker_changed:
                    changes.append({
                        "property_name": "superscript_marker",
                        "current_value": "Manually shrunk font size, not true superscript",
                        "target_value": f"True superscript at {expected_size}pt",
                        "evidence": "Author/affiliation marker should use superscript formatting, not a separately reduced font size",
                    })

        # Applied unconditionally, same reasoning as elsewhere in this
        # class - always correct regardless of what the checker
        # specifically flagged for this paragraph.
        expected_alignment = rules.get("alignment")
        current_alignment = get_paragraph_alignment(paragraph)
        if expected_alignment and current_alignment != expected_alignment:
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 1)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Author-information alignment did not match target rule",
            })

        # Author name, affiliation, and corresponding-author lines all
        # use the same standard spacing as ordinary body text (0/8pt,
        # single) rather than whatever pdf2docx's own line-layout
        # measurements happened to leave behind for each of them. Applied
        # unconditionally (not gated behind a checker-detected issue),
        # since these three target values are always correct regardless
        # of what else was flagged for this paragraph.
        expected_space_before = rules.get("space_before", 0)
        current_space_before = get_space_before_pt(paragraph)
        if (
            current_space_before is None
            or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_before_pt(paragraph, float(expected_space_before))
            changes.append({
                "property_name": "space_before",
                "current_value": (
                    f"{current_space_before}pt" if current_space_before is not None else "(inherited)"
                ),
                "target_value": f"{expected_space_before}pt",
                "evidence": "Author-information spacing before did not match target rule",
            })

        # Space after depends on where the line sits in the author block
        # (see _compute_author_block_after), not on its role keyword.
        expected_space_after = getattr(self, "_author_block_after", {}).get(
            index, rules.get("space_after", 8)
        )
        current_space_after = get_space_after_pt(paragraph)
        if (
            current_space_after is None
            or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_after_pt(paragraph, float(expected_space_after))
            changes.append({
                "property_name": "space_after",
                "current_value": (
                    f"{current_space_after}pt" if current_space_after is not None else "(inherited)"
                ),
                "target_value": f"{expected_space_after}pt",
                "evidence": "Author-information spacing after did not match target rule",
            })

        expected_line_spacing = rules.get("line_spacing", 1.0)
        current_line_spacing = paragraph.paragraph_format.line_spacing
        current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
        effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
        target_is_single = abs(float(expected_line_spacing) - 1.0) < 0.001
        needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
        if target_is_single and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
            needs_fix = True
        if needs_fix:
            paragraph.paragraph_format.line_spacing = expected_line_spacing
            changes.append({
                "property_name": "line_spacing",
                "current_value": (
                    f"{current_line_spacing_rule} ({current_line_spacing})"
                    if current_line_spacing is not None
                    else str(current_line_spacing_rule)
                ),
                "target_value": str(expected_line_spacing),
                "evidence": "Author-information line spacing did not match target rule",
            })

        # Left/right indent applied unconditionally, same reasoning as
        # the spacing block above - pdf2docx sometimes leaves a leftover
        # indent (commonly on the right side) from the source PDF's own
        # line-wrap measurements, and author-info lines should sit flush
        # with both margins like any other centered front-matter line.
        pf = paragraph.paragraph_format
        if pf.left_indent is not None and pf.left_indent != Emu(0):
            current_left = pf.left_indent
            pf.left_indent = Emu(0)
            changes.append({
                "property_name": "left_indent",
                "current_value": f"{current_left}",
                "target_value": "0",
                "evidence": "Author-information left indent did not match target rule",
            })
        if pf.right_indent is not None and pf.right_indent != Emu(0):
            current_right = pf.right_indent
            pf.right_indent = Emu(0)
            changes.append({
                "property_name": "right_indent",
                "current_value": f"{current_right}",
                "target_value": "0",
                "evidence": "Author-information right indent did not match target rule",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location=role.replace("_", " ").title(),
                change_type="author_info",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 50),
                paragraph_type=ParagraphType.AUTHOR_INFO.value,
            )

    def _fix_front_matter_spacing(self, paragraph, index: int) -> None:
        """Apply the standard 0/8pt/single spacing to a front-matter line
        that doesn't have its own dedicated fixer (e.g. "Corresponding
        author: ..." - classified as JOURNAL_HEADER but not actually the
        journal title/issue line, so _fix_journal_header never runs for
        it). Applied unconditionally, since these target values are
        always correct for this kind of line regardless of what else was
        or wasn't flagged for it.
        """
        changes = []
        current_space_before = get_space_before_pt(paragraph)
        if current_space_before is None or abs(float(current_space_before)) > 0.5 or has_lines_based_spacing(paragraph):
            set_space_before_pt(paragraph, 0.0)
            changes.append({
                "property_name": "space_before",
                "current_value": f"{current_space_before}pt" if current_space_before is not None else "(inherited)",
                "target_value": "0pt",
                "evidence": "Front-matter line spacing before did not match target rule",
            })

        expected_after = float(getattr(self, "_author_block_after", {}).get(index, 8.0))
        current_space_after = get_space_after_pt(paragraph)
        if current_space_after is None or abs(float(current_space_after) - expected_after) > 0.5 or has_lines_based_spacing(paragraph):
            set_space_after_pt(paragraph, expected_after)
            changes.append({
                "property_name": "space_after",
                "current_value": f"{current_space_after}pt" if current_space_after is not None else "(inherited)",
                "target_value": f"{expected_after:g}pt",
                "evidence": "Front-matter line spacing after did not match target rule",
            })

        current_line_spacing = paragraph.paragraph_format.line_spacing
        current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
        effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
        needs_fix = abs(effective_line_spacing - 1.0) > 0.01 or current_line_spacing_rule != WD_LINE_SPACING.SINGLE
        if needs_fix:
            paragraph.paragraph_format.line_spacing = 1.0
            changes.append({
                "property_name": "line_spacing",
                "current_value": (
                    f"{current_line_spacing_rule} ({current_line_spacing})"
                    if current_line_spacing is not None
                    else str(current_line_spacing_rule)
                ),
                "target_value": "1.0",
                "evidence": "Front-matter line spacing did not match target rule",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Front Matter",
                change_type="journal_header",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 50),
                paragraph_type=ParagraphType.JOURNAL_HEADER.value,
            )

    def _fix_heading(self, paragraph, index: int):
        """Fix section heading formatting with structured change records."""
        heading_rules = self._heading_rules_for_text(get_paragraph_text(paragraph))
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["headings"],
            fallback=[
                "font_name",
                "font_size",
                "bold",
                "italic",
                "capitalization",
                "number_font_name",
                "number_font_size",
                "number_bold",
                "alignment",
                "space_before",
                "space_after",
                "line_spacing",
            ],
        )

        expected_font = heading_rules.get("font_name", "Times New Roman")
        expected_size = heading_rules.get("font_size", 10)
        expected_bold = heading_rules.get("bold", None)
        expected_italic = heading_rules.get("italic", None)
        expected_all_caps = heading_rules.get("all_caps")
        expected_alignment = heading_rules.get("alignment")

        changes.extend(self._fix_numbering_formatting(
            paragraph,
            expected_font,
            expected_size,
            expected_bold,
            allowed_properties,
        ))

        current_alignment = get_paragraph_alignment(paragraph)
        if (
            expected_alignment
            and self._property_allowed("alignment", allowed_properties)
            and current_alignment != expected_alignment
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 0)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Heading alignment did not match target rule",
            })

        # All three applied unconditionally (not gated behind
        # allowed_properties) - these target values are always correct
        # for any heading now, regardless of what the checker specifically
        # flagged for this paragraph.
        expected_space_before = heading_rules.get("space_before", self._consistent_heading_space_before)
        if expected_space_before is not None:
            current_space_before = get_space_before_pt(paragraph)
            if (
                current_space_before is None
                or abs(float(current_space_before) - float(expected_space_before)) > 1.0
            or has_lines_based_spacing(paragraph)
            ):
                set_space_before_pt(paragraph, float(expected_space_before))
                changes.append({
                    "property_name": "space_before",
                    "current_value": (
                        f"{current_space_before}pt"
                        if current_space_before is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_before}pt",
                    "evidence": "Heading spacing before did not match target rule",
                })

        expected_space_after = heading_rules.get("space_after")
        if expected_space_after is not None:
            current_space_after = get_space_after_pt(paragraph)
            if (
                current_space_after is None
                or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_after_pt(paragraph, float(expected_space_after))
                changes.append({
                    "property_name": "space_after",
                    "current_value": (
                        f"{current_space_after}pt"
                        if current_space_after is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_after}pt",
                    "evidence": "Heading spacing after did not match target rule",
                })

        expected_line_spacing = heading_rules.get("line_spacing")
        if expected_line_spacing is not None:
            current_line_spacing = paragraph.paragraph_format.line_spacing
            current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
            effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
            target_is_single = abs(float(expected_line_spacing) - 1.0) < 0.001
            needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
            if target_is_single and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
                needs_fix = True
            if needs_fix:
                paragraph.paragraph_format.line_spacing = expected_line_spacing
                changes.append({
                    "property_name": "line_spacing",
                    "current_value": (
                        f"{current_line_spacing_rule} ({current_line_spacing})"
                        if current_line_spacing is not None
                        else str(current_line_spacing_rule)
                    ),
                    "target_value": str(expected_line_spacing),
                    "evidence": "Heading line spacing did not match target rule",
                })

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_italic,
                    expected_strike=False,
                    allowed_properties=allowed_properties,
                ))

        current_text = get_paragraph_text(paragraph)
        if (
            expected_all_caps
            and self._property_allowed("capitalization", allowed_properties)
            and not self._heading_text_is_all_caps(current_text)
        ):
            target_text = current_text.upper()
            self._replace_paragraph_text_preserving_first_run(paragraph, target_text)
            changes.append({
                "property_name": "capitalization",
                "current_value": current_text,
                "target_value": target_text,
                "evidence": "Heading capitalization did not match target rule",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Section Heading",
                change_type="heading",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.SECTION_HEADING.value,
            )

    def _heading_rules_for_text(self, text: str) -> Dict[str, Any]:
        """Return the correct heading rule for main headings or subheadings."""
        stripped = re.sub(r"\s+", " ", text.strip())
        # Main heading: "1.0", "2.0", ... - the decimal part is exactly
        # zero. Checked before the subheading pattern below, since "1.0"
        # would otherwise also match "^\d+\.\d+" and incorrectly be
        # treated as a subheading (picking up subheading's italic, for
        # instance). Space after the number is optional - pdf2docx (or
        # an author's own typing) sometimes drops it entirely ("1.0Introduction"),
        # and the heading is no less a main heading for that.
        if re.match(r"^\d+\.0(?!\d)\s*", stripped):
            return self.rules.get("heading", {})
        if re.match(r"^\d+\.\d+", stripped):
            return self.rules.get("subheading", self.rules.get("heading", {}))
        base_rules = dict(self.rules.get("heading", {}))
        if stripped.lower() == "abstract":
            # Abstract is a one-off heading - it isn't part of either the
            # numbered-section cluster or the declaration-heading cluster
            # (Acknowledgements, References, ...), so it shouldn't be
            # pulled into either of those groups' own spacing convention.
            base_rules["space_before"] = 0
            base_rules["space_after"] = 8
            base_rules["line_spacing"] = 1.0
        return base_rules

    @staticmethod
    def _heading_text_is_all_caps(text: str) -> bool:
        """Return True when heading words are uppercase after removing numbering."""
        stripped = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s+", "", text or "").strip()
        letters = [char for char in stripped if char.isalpha()]
        return bool(letters) and all(char.isupper() for char in letters)

    def _get_numbering_level(self, paragraph):
        """Return the numbering level XML element for a numbered paragraph."""
        p_pr = paragraph._p.pPr
        if p_pr is None or p_pr.numPr is None:
            return None

        num_id_element = p_pr.numPr.numId
        ilvl_element = p_pr.numPr.ilvl
        if num_id_element is None:
            return None

        num_id = num_id_element.val
        ilvl = str(ilvl_element.val if ilvl_element is not None else 0)
        numbering = self.document.part.numbering_part.element

        abstract_num_id = None
        for num in numbering.findall(qn("w:num")):
            if num.get(qn("w:numId")) == str(num_id):
                abstract_node = num.find(qn("w:abstractNumId"))
                if abstract_node is not None:
                    abstract_num_id = abstract_node.get(qn("w:val"))
                break

        if abstract_num_id is None:
            return None

        for abstract_num in numbering.findall(qn("w:abstractNum")):
            if abstract_num.get(qn("w:abstractNumId")) != str(abstract_num_id):
                continue
            for level in abstract_num.findall(qn("w:lvl")):
                if level.get(qn("w:ilvl")) == ilvl:
                    return level

        return None

    def _get_or_create_child(self, parent, tag_name: str):
        """Return a child XML node, creating it when missing."""
        child = parent.find(qn(tag_name))
        if child is None:
            child = OxmlElement(tag_name)
            parent.append(child)
        return child

    def _numbering_bold_value(self, r_pr) -> Optional[bool]:
        """Read numbering bold value from a numbering rPr node."""
        bold = r_pr.find(qn("w:b")) if r_pr is not None else None
        if bold is None:
            return None
        value = bold.get(qn("w:val"))
        return value not in {"0", "false", "False", "off"}

    def _fix_numbering_formatting(
        self,
        paragraph,
        expected_font: str,
        expected_size: float,
        expected_bold: Optional[bool],
        allowed_properties: Optional[set] = None,
    ) -> List[Dict[str, str]]:
        """Fix Word list numbering formatting for numbered headings."""
        level = self._get_numbering_level(paragraph)
        if level is None:
            return []

        changes = []
        r_pr = self._get_or_create_child(level, "w:rPr")

        if expected_font and self._property_allowed("number_font_name", allowed_properties):
            r_fonts = self._get_or_create_child(r_pr, "w:rFonts")
            current_font = r_fonts.get(qn("w:ascii")) or r_fonts.get(qn("w:hAnsi"))
            if current_font and not is_font_equivalent(current_font, expected_font):
                changes.append({
                    "property_name": "number_font_name",
                    "current_value": current_font,
                    "target_value": expected_font,
                    "evidence": "Heading number font did not match target rule",
                })
            r_fonts.set(qn("w:ascii"), expected_font)
            r_fonts.set(qn("w:hAnsi"), expected_font)

        if expected_size and self._property_allowed("number_font_size", allowed_properties):
            size_node = self._get_or_create_child(r_pr, "w:sz")
            current_size = None
            if size_node.get(qn("w:val")):
                current_size = int(size_node.get(qn("w:val"))) / 2
            if current_size is not None and abs(current_size - expected_size) > 0.5:
                changes.append({
                    "property_name": "number_font_size",
                    "current_value": f"{current_size:.1f} pt",
                    "target_value": f"{expected_size} pt",
                    "evidence": "Heading number font size did not match target rule",
                })
            size_node.set(qn("w:val"), str(int(float(expected_size) * 2)))

        if expected_bold is not None and self._property_allowed("number_bold", allowed_properties):
            current_bold = self._numbering_bold_value(r_pr)
            if current_bold != expected_bold:
                changes.append({
                    "property_name": "number_bold",
                    "current_value": "Bold" if current_bold else "Not Bold",
                    "target_value": "Bold" if expected_bold else "Not Bold",
                    "evidence": "Heading number bold setting did not match target rule",
                })
            bold_node = self._get_or_create_child(r_pr, "w:b")
            bold_node.set(qn("w:val"), "1" if expected_bold else "0")

        return changes

    def _insert_tab_after_marker(self, paragraph, marker_end: int) -> bool:
        """Insert a tab right after a list-item marker if none exists there.

        A manually-lettered/numbered list item ("a. To quantify...") is
        meant to have a tab between the marker and the item text, so the
        text lines up at the paragraph's tab stop the same way the
        template's list items do. When that tab is missing - which
        happens when pdf2docx's conversion drops it, or an earlier
        paragraph-splitting repair doesn't carry it through - the marker
        and text run together with no visible gap at all ("a.To
        quantify..."). Returns True if a tab was actually inserted.
        """
        cursor = 0
        for run in paragraph.runs:
            run_text = run.text or ""
            run_start, run_end = cursor, cursor + len(run_text)
            cursor = run_end
            if not (run_start <= marker_end <= run_end):
                continue

            offset_in_run = marker_end - run_start
            if offset_in_run < len(run_text) and run_text[offset_in_run] in ("\t", " "):
                return False  # already separated, nothing to do

            r_element = run._r
            parent = r_element.getparent()
            insert_index = list(parent).index(r_element)

            before_text = run_text[:offset_in_run]
            after_text = run_text[offset_in_run:]

            before_el = deepcopy(r_element)
            self._set_run_element_text(before_el, before_text)

            tab_el = deepcopy(r_element)
            for t_el in tab_el.findall(qn("w:t")):
                tab_el.remove(t_el)
            for existing_tab in tab_el.findall(qn("w:tab")):
                tab_el.remove(existing_tab)
            tab_el.append(OxmlElement("w:tab"))

            after_el = deepcopy(r_element)
            self._set_run_element_text(after_el, after_text)

            new_elements = [tab_el]
            if before_text:
                new_elements.insert(0, before_el)
            if after_text:
                new_elements.append(after_el)

            for offset, element in enumerate(new_elements):
                parent.insert(insert_index + offset, element)
            parent.remove(r_element)
            return True
        return False

    def _fix_body_text(self, paragraph, index: int):
        """Fix body text formatting with structured change records."""
        body_rules = self.rules.get("body", {})
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["body_text", "line_spacing"],
            fallback=["font_name", "font_size", "bold", "line_spacing", "alignment", "space_after"],
        )

        expected_font = body_rules.get("font_name", "Times New Roman")
        expected_size = body_rules.get("font_size", 12)
        # Bold in body text is the author's own inline emphasis (a bold
        # label such as "Objective 1:" followed by ordinary text, a bold
        # key term, ...) and is left exactly as written - only font,
        # size, spacing, alignment and the like are normalised. Passing
        # None for expected_bold below makes _fix_run_formatting skip the
        # bold property entirely instead of forcing every run to the
        # rule's bold=False, which silently erased such labels.
        expected_bold = None
        expected_alignment = body_rules.get("alignment")
        expected_space_after = body_rules.get("space_after")

        # A paragraph with a real w:numPr is a genuine Word numbered list
        # item (converted from manually-typed "a."/"1." text by the
        # repair stage) - its indentation is governed by the numbering
        # definition itself, and directly setting left/right indent here
        # would override and break that. Left alone entirely below.
        #
        # A paragraph whose text still starts with "a."/"1." despite
        # that conversion never runs here anymore - it used to get a
        # cosmetic hanging-indent/tab treatment that made it look
        # reasonably formatted while still not being real numbering,
        # which made a failed conversion easy to miss. Now it's treated
        # as ordinary body text like anything else, so a conversion
        # failure stays visibly obvious (flush left, no hanging indent)
        # instead of being quietly disguised.
        has_real_numbering = (
            paragraph._p.pPr is not None
            and paragraph._p.pPr.find(qn("w:numPr")) is not None
        )

        # font_name/font_size/bold are checked and fixed per run,
        # unconditionally (not gated behind allowed_properties) - the
        # checker's own detection for these compares only the
        # paragraph's *most common* run value against the target, so a
        # minority-mismatched run within an otherwise-correct paragraph
        # (e.g. one fragment still at the wrong size while the rest of
        # the paragraph was already fixed) never gets flagged as an
        # issue at all, and the gate would silently leave it uncorrected
        # forever. These three are core, unambiguous formatting
        # requirements with no legitimate reason to vary within a body
        # paragraph, so enforcing them on every run is always correct
        # regardless of whether the checker's paragraph-level summary
        # happened to miss the inconsistency.
        core_font_properties = {"font_name", "font_size", "bold"}
        run_allowed_properties = (
            set(allowed_properties) | core_font_properties
            if allowed_properties is not None
            else None
        )

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    allowed_properties=run_allowed_properties,
                ))

        expected_spacing = body_rules.get("line_spacing")
        if expected_spacing and self._property_allowed("line_spacing", allowed_properties):
            paragraph_format = paragraph.paragraph_format
            if paragraph_format.line_spacing != expected_spacing:
                current_spacing = paragraph_format.line_spacing
                paragraph_format.line_spacing = expected_spacing
                changes.append({
                    "property_name": "line_spacing",
                    "current_value": str(current_spacing) if current_spacing is not None else "(inherited)",
                    "target_value": str(expected_spacing),
                    "evidence": "Paragraph line spacing did not match target rule",
                })

        if (
            expected_space_after is not None
            and self._property_allowed("space_after", allowed_properties)
        ):
            current_space_after = get_space_after_pt(paragraph)
            if (
                current_space_after is None
                or abs(float(current_space_after) - float(expected_space_after)) > 0.05
            or has_lines_based_spacing(paragraph)
            ):
                set_space_after_pt(paragraph, float(expected_space_after))
                changes.append({
                    "property_name": "space_after",
                    "current_value": (
                        f"{current_space_after}pt"
                        if current_space_after is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_after}pt",
                    "evidence": "Body paragraph spacing after did not match target rule",
                })

        # Space before is virtually always 0 for body text (space_after
        # alone controls the gap between paragraphs) - pdf2docx sometimes
        # leaves a leftover "before" value from the source PDF's own line
        # spacing instead, which then stacks with space_after and throws
        # off vertical rhythm. Reset unconditionally rather than only
        # when the template specifies a nonzero value, since a nonzero
        # leftover is never actually intentional here. A tight 0.05pt
        # tolerance (not the usual 0.5pt) catches small leftover
        # artifacts like 0.4pt/0.5pt that a looser check let slip through
        # - these are exactly the kind of "almost zero" residue pdf2docx
        # leaves, not an intentional value anyone would set on purpose.
        expected_space_before = body_rules.get("space_before", 0)
        if (
            expected_space_before is not None
            and self._property_allowed("space_after", allowed_properties)
        ):
            current_space_before = get_space_before_pt(paragraph)
            if (
                current_space_before is None
                or abs(float(current_space_before) - float(expected_space_before)) > 0.05
            or has_lines_based_spacing(paragraph)
            ):
                set_space_before_pt(paragraph, float(expected_space_before))
                changes.append({
                    "property_name": "space_before",
                    "current_value": (
                        f"{current_space_before}pt"
                        if current_space_before is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_before}pt",
                    "evidence": "Body paragraph spacing before did not match target rule",
                })

        # Left/right indent of a body paragraph is deliberately NOT
        # normalised. Input is native Word documents, where an indent is
        # the author's own intentional layout (e.g. the one-tab indent in
        # front of "Objective 1: ..." lines), not a leftover from PDF
        # conversion - the earlier "reset every body indent to 0" logic
        # assumed the latter and wiped such indents out. Indentation is
        # still normalised where the template actually specifies it
        # (author/affiliation lines, abstract, references, lists).

        # Outline level: body text should be Word's "Level 1" (w:outlineLvl
        # val="0") rather than the default "Body Text" level (no
        # w:outlineLvl at all), which is what pdf2docx output leaves it
        # as - this affects how the paragraph appears in Word's Navigation
        # pane and outline view. A real numbered-list paragraph is left at
        # the default "Body Text" level - it isn't a structural unit the
        # way a real body paragraph is, so it shouldn't be promoted.
        ppr = paragraph._p.get_or_add_pPr()
        outline_element = ppr.find(qn("w:outlineLvl"))
        if has_real_numbering:
            if outline_element is not None:
                ppr.remove(outline_element)
                changes.append({
                    "property_name": "outline_level",
                    "current_value": "Level 1",
                    "target_value": "Body Text",
                    "evidence": "List items stay at Body Text outline level, not Level 1",
                })
        else:
            if outline_element is None or outline_element.get(qn("w:val")) != "0":
                if outline_element is None:
                    outline_element = OxmlElement("w:outlineLvl")
                    ppr.append(outline_element)
                outline_element.set(qn("w:val"), "0")
                changes.append({
                    "property_name": "outline_level",
                    "current_value": "Body Text",
                    "target_value": "Level 1",
                    "evidence": "Body paragraph outline level did not match target rule",
                })

        # The "Received:", "Accepted:", "Published:" labels on the
        # submission-dates line are each individually bold in the
        # template; the dates themselves are not. This line is otherwise
        # ordinary body text (no dedicated paragraph type), so it's
        # handled here rather than by a separate classification.
        # Applied unconditionally (not gated behind allowed_properties) -
        # the core per-run font fix above unconditionally normalizes
        # bold to body text's own False for every run first (correctly,
        # since that's what body text should be), which would otherwise
        # permanently erase these labels' own bold=True with no later
        # step allowed to restore it if this were still gated behind
        # whatever bold mismatch the checker's own paragraph-level
        # majority vote happened to detect (or, as here, didn't).
        raw_text = "".join(run.text or "" for run in paragraph.runs)
        if re.search(r"^received\s*:", raw_text.strip(), re.IGNORECASE):
            label_bolded = False
            for label_match in re.finditer(r"(?:,\s*)?(?:received|accepted|published)\s*:\s*", raw_text, re.IGNORECASE):
                if self._apply_bold_span(paragraph, label_match.start(), label_match.end()):
                    label_bolded = True
            if label_bolded:
                changes.append({
                    "property_name": "submission_dates_label_bold",
                    "current_value": "Received/Accepted/Published labels not bold",
                    "target_value": "Received/Accepted/Published labels bold",
                    "evidence": "Submission date labels should be bold per template",
                })

        current_alignment = get_paragraph_alignment(paragraph)
        if (
            expected_alignment
            and self._property_allowed("alignment", allowed_properties)
            and current_alignment != expected_alignment
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 3)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Body text alignment did not match target rule",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location=f"Body Text (Para {index + 1})",
                change_type="body",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.BODY.value,
            )

    def _fix_abstract(self, paragraph, index: int):
        """Fix abstract formatting with structured change records."""
        abstract_rules = self.rules.get("abstract", {})
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["body_text", "line_spacing"],
            fallback=["font_name", "font_size", "bold", "line_spacing", "alignment", "space_after"],
        )

        expected_font = abstract_rules.get("font_name", "Times New Roman")
        expected_size = abstract_rules.get("font_size", 9)
        expected_bold = abstract_rules.get("bold", None)
        # Some profiles do not explicitly repeat the body alignment in the
        # abstract rules. Abstract body text should still follow the same
        # paragraph alignment convention as the body; otherwise a converted
        # manuscript can remain LEFT even when the profile requires JUSTIFY.
        expected_alignment = abstract_rules.get(
            "alignment",
            self.rules.get("body", {}).get("alignment", "JUSTIFY"),
        )
        expected_space_after = abstract_rules.get(
            "space_after",
            self.rules.get("body", {}).get("space_after"),
        )

        # Unconditional per-run font check - same reasoning as
        # _fix_body_text: the checker's paragraph-level "most common
        # run value" detection can miss a minority-mismatched run.
        core_font_properties = {"font_name", "font_size", "bold"}
        run_allowed_properties = (
            set(allowed_properties) | core_font_properties
            if allowed_properties is not None
            else None
        )
        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    allowed_properties=run_allowed_properties,
                ))

        expected_spacing = abstract_rules.get("line_spacing", self.rules.get("body", {}).get("line_spacing"))
        if expected_spacing and self._property_allowed("line_spacing", allowed_properties):
            paragraph_format = paragraph.paragraph_format
            if paragraph_format.line_spacing != expected_spacing:
                current_spacing = paragraph_format.line_spacing
                paragraph_format.line_spacing = expected_spacing
                changes.append({
                    "property_name": "line_spacing",
                    "current_value": str(current_spacing) if current_spacing is not None else "(inherited)",
                    "target_value": str(expected_spacing),
                    "evidence": "Abstract line spacing did not match target rule",
                })

        if (
            expected_space_after is not None
            and self._property_allowed("space_after", allowed_properties)
        ):
            current_space_after = get_space_after_pt(paragraph)
            if (
                current_space_after is None
                or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_after_pt(paragraph, float(expected_space_after))
                changes.append({
                    "property_name": "space_after",
                    "current_value": (
                        f"{current_space_after}pt"
                        if current_space_after is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_after}pt",
                    "evidence": "Abstract paragraph spacing after did not match target rule",
                })

        # Space before and outline level are applied unconditionally
        # (not gated behind allowed_properties) - pdf2docx commonly
        # leaves a leftover "space before" value from the source PDF's
        # own line layout on abstract-body paragraphs, and outline level
        # was never checked/fixed for the abstract body at all before.
        expected_space_before = abstract_rules.get("space_before", 0)
        current_space_before = get_space_before_pt(paragraph)
        if (
            current_space_before is None
            or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_before_pt(paragraph, float(expected_space_before))
            changes.append({
                "property_name": "space_before",
                "current_value": (
                    f"{current_space_before}pt" if current_space_before is not None else "(inherited)"
                ),
                "target_value": f"{expected_space_before}pt",
                "evidence": "Abstract paragraph spacing before did not match target rule",
            })

        ppr = paragraph._p.get_or_add_pPr()
        outline_element = ppr.find(qn("w:outlineLvl"))
        if outline_element is None or outline_element.get(qn("w:val")) != "0":
            if outline_element is None:
                outline_element = OxmlElement("w:outlineLvl")
                ppr.append(outline_element)
            outline_element.set(qn("w:val"), "0")
            changes.append({
                "property_name": "outline_level",
                "current_value": "Body Text",
                "target_value": "Level 1",
                "evidence": "Abstract paragraph outline level did not match target rule",
            })

        # Left/right indent applied unconditionally - pdf2docx sometimes
        # leaves a leftover indent from the source PDF's own line-wrap
        # measurements, and the abstract body should sit flush with both
        # margins like ordinary body text.
        pf = paragraph.paragraph_format
        if pf.left_indent is not None and pf.left_indent != Emu(0):
            current_left = pf.left_indent
            pf.left_indent = Emu(0)
            changes.append({
                "property_name": "left_indent",
                "current_value": f"{current_left}",
                "target_value": "0",
                "evidence": "Abstract paragraph left indent did not match target rule",
            })
        if pf.right_indent is not None and pf.right_indent != Emu(0):
            current_right = pf.right_indent
            pf.right_indent = Emu(0)
            changes.append({
                "property_name": "right_indent",
                "current_value": f"{current_right}",
                "target_value": "0",
                "evidence": "Abstract paragraph right indent did not match target rule",
            })

        # Abstract body alignment is a structural template requirement.
        # Apply it directly instead of depending on whether the checker
        # happened to report an alignment issue for this paragraph.
        current_alignment = get_paragraph_alignment(paragraph)
        if expected_alignment and current_alignment != expected_alignment:
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 3)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Abstract body alignment should match the template body alignment",
            })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Abstract",
                change_type="abstract",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.ABSTRACT_CONTENT.value,
            )

    def _fix_keywords(self, paragraph, index: int):
        """Fix keyword formatting with structured change records."""
        keywords_rules = self.rules.get("keywords", self.rules.get("abstract", {}))
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["body_text"],
            fallback=["font_name", "font_size", "bold", "italic", "capitalization"],
        )

        expected_font = keywords_rules.get("font_name", "Times New Roman")
        expected_size = keywords_rules.get("font_size", 9)
        expected_bold = keywords_rules.get("bold", None)
        expected_italic = keywords_rules.get("italic", None)
        capitalize_first_letter = keywords_rules.get("capitalize_first_letter")

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_italic,
                    allowed_properties=allowed_properties,
                ))

        current_text = get_paragraph_text(paragraph)
        if (
            capitalize_first_letter
            and self._property_allowed("capitalization", allowed_properties)
        ):
            target_text = self._capitalize_keyword_text(current_text)
            if target_text != current_text:
                self._replace_paragraph_text_preserving_first_run(paragraph, target_text)
                changes.append({
                    "property_name": "capitalization",
                    "current_value": current_text,
                    "target_value": target_text,
                    "evidence": "Keyword capitalization did not match target rule",
                })

        # The "Keywords:" label itself is bold in the template; the actual
        # keyword list that follows is not. Applied as a span so only the
        # label - not the whole line - gets bolded. Matched against the
        # raw run concatenation (not get_paragraph_text, which may
        # normalize whitespace) since _apply_bold_span's offsets are
        # computed the same way.
        raw_text = "".join(run.text or "" for run in paragraph.runs)
        label_match = re.match(r"^\s*(?:keywords?|key\s+words?)\s*:?\s*", raw_text, re.IGNORECASE)
        if label_match and self._property_allowed("bold", allowed_properties):
            if self._apply_bold_span(paragraph, label_match.start(), label_match.end()):
                changes.append({
                    "property_name": "keywords_label_bold",
                    "current_value": "Keywords label not bold",
                    "target_value": "Keywords label bold",
                    "evidence": "Keywords label should be bold per template",
                })

        # Applied unconditionally - same reasoning as title/abstract/etc:
        # these three target values are always correct regardless of
        # what else was specifically flagged for this paragraph.
        expected_space_before = keywords_rules.get("space_before", 0)
        current_space_before = get_space_before_pt(paragraph)
        if (
            current_space_before is None
            or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_before_pt(paragraph, float(expected_space_before))
            changes.append({
                "property_name": "space_before",
                "current_value": f"{current_space_before}pt" if current_space_before is not None else "(inherited)",
                "target_value": f"{expected_space_before}pt",
                "evidence": "Keywords spacing before did not match target rule",
            })

        expected_space_after = keywords_rules.get("space_after", 8)
        current_space_after = get_space_after_pt(paragraph)
        if (
            current_space_after is None
            or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
        ):
            set_space_after_pt(paragraph, float(expected_space_after))
            changes.append({
                "property_name": "space_after",
                "current_value": f"{current_space_after}pt" if current_space_after is not None else "(inherited)",
                "target_value": f"{expected_space_after}pt",
                "evidence": "Keywords spacing after did not match target rule",
            })

        expected_line_spacing = keywords_rules.get("line_spacing", 1.0)
        current_line_spacing = paragraph.paragraph_format.line_spacing
        current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
        effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
        needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
        if abs(float(expected_line_spacing) - 1.0) < 0.001 and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
            needs_fix = True
        if needs_fix:
            paragraph.paragraph_format.line_spacing = expected_line_spacing
            changes.append({
                "property_name": "line_spacing",
                "current_value": (
                    f"{current_line_spacing_rule} ({current_line_spacing})"
                    if current_line_spacing is not None
                    else str(current_line_spacing_rule)
                ),
                "target_value": str(expected_line_spacing),
                "evidence": "Keywords line spacing did not match target rule",
            })

        # Applied unconditionally, same reasoning as the spacing fixes
        # above - always correct regardless of what the checker flagged.
        expected_alignment = keywords_rules.get("alignment")
        if expected_alignment:
            current_alignment = get_paragraph_alignment(paragraph)
            if current_alignment != expected_alignment:
                paragraph.alignment = WD_ALIGN_PARAGRAPH(
                    ALIGNMENT_REVERSE_MAP.get(expected_alignment, 3)
                )
                changes.append({
                    "property_name": "alignment",
                    "current_value": current_alignment,
                    "target_value": expected_alignment,
                    "evidence": "Keywords alignment did not match target rule",
                })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Keywords",
                change_type="keywords",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.KEYWORDS_CONTENT.value,
            )

    @staticmethod
    def _capitalize_keyword_text(text: str) -> str:
        """Capitalize the first letter of each existing keyword item."""
        match = re.match(
            r"^(\s*(?:keywords?|key\s+words?)\s*(?:-|\u2013|\u2014|:)?\s*)(.*)$",
            text or "",
            flags=re.IGNORECASE,
        )
        if not match:
            return text

        prefix, keyword_text = match.groups()
        parts = re.split(r"([,;])", keyword_text)
        for index in range(0, len(parts), 2):
            parts[index] = re.sub(
                r"^(\s*)([A-Za-z])",
                lambda item: item.group(1) + item.group(2).upper(),
                parts[index],
                count=1,
            )
        return prefix + "".join(parts)

    def _fix_caption(self, paragraph, index: int):
        """Fix caption formatting with structured change records."""
        caption_rules = self.rules.get("caption", {})
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["figures", "tables"],
            fallback=["font_name", "font_size", "italic", "space_after", "capitalization"],
        )

        expected_font = caption_rules.get("font_name", "Times New Roman")
        expected_size = caption_rules.get("font_size", 10)
        expected_bold = caption_rules.get("bold", None)
        expected_italic = caption_rules.get("italic", None)
        expected_space_after = caption_rules.get("space_after")
        expected_space_before = caption_rules.get("space_before")
        expected_line_spacing = caption_rules.get("line_spacing")
        title_case_required = bool(caption_rules.get("title_case"))

        for run in paragraph.runs:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_italic,
                    allowed_properties=allowed_properties,
                ))

        # space_before/after/line_spacing applied unconditionally - same
        # reasoning as elsewhere: always correct regardless of what the
        # checker specifically flagged for this paragraph.
        if expected_space_before is not None:
            current_space_before = get_space_before_pt(paragraph)
            if (
                current_space_before is None
                or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_before_pt(paragraph, float(expected_space_before))
                changes.append({
                    "property_name": "space_before",
                    "current_value": (
                        f"{current_space_before}pt" if current_space_before is not None else "(inherited)"
                    ),
                    "target_value": f"{expected_space_before}pt",
                    "evidence": "Caption paragraph spacing before did not match target rule",
                })

        if expected_space_after is not None:
            current_space_after = get_space_after_pt(paragraph)
            if (
                current_space_after is None
                or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_after_pt(paragraph, float(expected_space_after))
                changes.append({
                    "property_name": "space_after",
                    "current_value": (
                        f"{current_space_after}pt"
                        if current_space_after is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_after}pt",
                    "evidence": "Caption paragraph spacing after did not match target rule",
                })

        if expected_line_spacing is not None:
            current_line_spacing = paragraph.paragraph_format.line_spacing
            current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
            effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
            needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
            if abs(float(expected_line_spacing) - 1.0) < 0.001 and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
                needs_fix = True
            if needs_fix:
                paragraph.paragraph_format.line_spacing = expected_line_spacing
                changes.append({
                    "property_name": "line_spacing",
                    "current_value": (
                        f"{current_line_spacing_rule} ({current_line_spacing})"
                        if current_line_spacing is not None
                        else str(current_line_spacing_rule)
                    ),
                    "target_value": str(expected_line_spacing),
                    "evidence": "Caption line spacing did not match target rule",
                })

        current_text = get_paragraph_text(paragraph)
        if (
            title_case_required
            and self._property_allowed("capitalization", allowed_properties)
        ):
            target_text = to_journal_caption_title_case(current_text)
            if target_text != current_text:
                self._replace_paragraph_text_preserving_first_run(paragraph, target_text)
                changes.append({
                    "property_name": "capitalization",
                    "current_value": current_text,
                    "target_value": target_text,
                    "evidence": "Caption capitalization did not match target rule",
                })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location="Caption",
                change_type="caption",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.CAPTION.value,
            )

    @staticmethod
    def _set_run_element_italic(r_element, value: bool) -> None:
        """Set or clear italic on a raw <w:r> element's run properties."""
        r_pr = r_element.find(qn("w:rPr"))
        if value:
            if r_pr is None:
                r_pr = OxmlElement("w:rPr")
                r_element.insert(0, r_pr)
            i_el = r_pr.find(qn("w:i"))
            if i_el is None:
                r_pr.append(OxmlElement("w:i"))
            elif i_el.get(qn("w:val")) is not None:
                # An existing <w:i w:val="0"/> (explicit off) would
                # otherwise silently survive untouched, since merely
                # checking "does a <w:i> element exist" doesn't
                # distinguish an explicit off from an implicit on.
                del i_el.attrib[qn("w:val")]
        elif r_pr is not None:
            i_el = r_pr.find(qn("w:i"))
            if i_el is not None:
                r_pr.remove(i_el)

    @staticmethod
    def _set_run_element_text(r_element, text: str) -> None:
        """Set a raw <w:r> element's visible text, collapsing to one <w:t>."""
        t_elements = r_element.findall(qn("w:t"))
        if not t_elements:
            t_element = OxmlElement("w:t")
            t_element.set(qn("xml:space"), "preserve")
            t_element.text = text
            r_element.append(t_element)
            return
        t_elements[0].text = text
        t_elements[0].set(qn("xml:space"), "preserve")
        for extra in t_elements[1:]:
            r_element.remove(extra)

    def _apply_italic_span(self, paragraph, start: int, end: int) -> bool:
        """Italicize exactly the [start, end) character span of a paragraph's
        text, splitting the underlying runs as needed so every other
        character keeps its existing formatting untouched. Returns True if
        anything was changed.
        """
        if start >= end:
            return False

        changed_any = False
        cursor = 0
        for run in list(paragraph.runs):
            run_text = run.text or ""
            run_start = cursor
            run_end = cursor + len(run_text)
            cursor = run_end

            overlap_start = max(start, run_start)
            overlap_end = min(end, run_end)
            if overlap_start >= overlap_end:
                continue

            if overlap_start == run_start and overlap_end == run_end:
                if not run.font.italic:
                    run.font.italic = True
                    changed_any = True
                continue

            before_text = run_text[: overlap_start - run_start]
            middle_text = run_text[overlap_start - run_start : overlap_end - run_start]
            after_text = run_text[overlap_end - run_start :]

            r_element = run._r
            parent = r_element.getparent()
            insert_index = list(parent).index(r_element)

            new_elements = []
            for piece_text, make_italic in (
                (before_text, False),
                (middle_text, True),
                (after_text, False),
            ):
                if not piece_text:
                    continue
                piece_element = deepcopy(r_element)
                self._set_run_element_text(piece_element, piece_text)
                if make_italic:
                    self._set_run_element_italic(piece_element, True)
                new_elements.append(piece_element)

            for offset, element in enumerate(new_elements):
                parent.insert(insert_index + offset, element)
            parent.remove(r_element)
            changed_any = True

        return changed_any

    @staticmethod
    def _set_run_element_bold(r_element, value: bool) -> None:
        """Set or clear bold on a raw <w:r> element's run properties."""
        r_pr = r_element.find(qn("w:rPr"))
        if value:
            if r_pr is None:
                r_pr = OxmlElement("w:rPr")
                r_element.insert(0, r_pr)
            b_el = r_pr.find(qn("w:b"))
            if b_el is None:
                r_pr.append(OxmlElement("w:b"))
            elif b_el.get(qn("w:val")) is not None:
                # An existing <w:b w:val="0"/> (explicit off) would
                # otherwise silently survive untouched, since merely
                # checking "does a <w:b> element exist" doesn't
                # distinguish an explicit off from an implicit on.
                del b_el.attrib[qn("w:val")]
        elif r_pr is not None:
            b_el = r_pr.find(qn("w:b"))
            if b_el is not None:
                r_pr.remove(b_el)

    def _apply_bold_span(self, paragraph, start: int, end: int) -> bool:
        """Bold exactly the [start, end) character span of a paragraph's
        text, splitting the underlying runs as needed so every other
        character keeps its existing formatting untouched. Returns True if
        anything was changed. Mirrors _apply_italic_span for bold instead.
        """
        if start >= end:
            return False

        changed_any = False
        cursor = 0
        for run in list(paragraph.runs):
            run_text = run.text or ""
            run_start = cursor
            run_end = cursor + len(run_text)
            cursor = run_end

            overlap_start = max(start, run_start)
            overlap_end = min(end, run_end)
            if overlap_start >= overlap_end:
                continue

            if overlap_start == run_start and overlap_end == run_end:
                if not run.font.bold:
                    run.font.bold = True
                    changed_any = True
                continue

            before_text = run_text[: overlap_start - run_start]
            middle_text = run_text[overlap_start - run_start : overlap_end - run_start]
            after_text = run_text[overlap_end - run_start :]

            r_element = run._r
            parent = r_element.getparent()
            insert_index = list(parent).index(r_element)

            new_elements = []
            for piece_text, make_bold in (
                (before_text, False),
                (middle_text, True),
                (after_text, False),
            ):
                if not piece_text:
                    continue
                piece_element = deepcopy(r_element)
                self._set_run_element_text(piece_element, piece_text)
                if make_bold:
                    self._set_run_element_bold(piece_element, True)
                new_elements.append(piece_element)

            for offset, element in enumerate(new_elements):
                parent.insert(insert_index + offset, element)
            parent.remove(r_element)
            changed_any = True

        return changed_any


    def _fix_reference_publication_italic(self, paragraph) -> bool:
        """Italicize the journal/periodical name and volume in a reference.

        Targets only the common APA7 "...(Year). Title. *Journal Name,
        Volume*(Issue), pages." shape. References that don't match this
        pattern (books, edited chapters, conference proceedings without a
        volume number, etc.) are left exactly as they are - guessing wrong
        here would corrupt an entry that might already be correct.
        """
        full_text = "".join(run.text or "" for run in paragraph.runs)
        year_match = _REFERENCE_YEAR_RE.search(full_text)
        search_start = year_match.end() if year_match else 0
        source_match = _REFERENCE_SOURCE_RE.search(full_text, search_start)
        if source_match is None:
            return False

        # italic = journal name + volume number (the "(issue)" part stays upright)
        span_start, span_end = source_match.start(), source_match.end("volume")
        changed = self._capitalise_source_span(paragraph, source_match.start("source"), source_match.end("source"))

        # Already italic across the whole span? Nothing more to do.
        cursor = 0
        already_italic = True
        for run in paragraph.runs:
            run_text = run.text or ""
            run_start = cursor
            run_end = cursor + len(run_text)
            cursor = run_end
            overlap_start = max(span_start, run_start)
            overlap_end = min(span_end, run_end)
            if overlap_start < overlap_end and not run.font.italic:
                already_italic = False
                break
        if already_italic:
            return changed

        return self._apply_italic_span(paragraph, span_start, span_end) or changed

    _MINOR_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "nor", "of", "on", "or",
                    "the", "to", "with", "via", "per", "vs", "de", "la", "le", "du", "des", "et", "und", "der", "die", "das"}

    def _capitalise_source_span(self, paragraph, start: int, end: int) -> bool:
        """Journal name: capitalise the first letter of every content word (lower-case words only;
        acronyms and words that already have capitals are left alone). Length never changes."""
        full = "".join(run.text or "" for run in paragraph.runs)
        seg = full[start:end]
        new_chars = list(seg)
        words = list(re.finditer(r"[A-Za-z][A-Za-z'’]*", seg))
        for n, m in enumerate(words):
            w = m.group(0)
            if not w.islower():
                continue
            if n > 0 and w in self._MINOR_WORDS:
                continue
            # skip words glued to digits / symbols like "e-learning" prefixes are still words; keep simple
            new_chars[m.start()] = w[0].upper()
        new = "".join(new_chars)
        if new == seg:
            return False
        cursor = 0
        for run in paragraph.runs:
            t = run.text or ""
            r0, r1 = cursor, cursor + len(t)
            cursor = r1
            o0, o1 = max(start, r0), min(end, r1)
            if o0 >= o1:
                continue
            a, b = o0 - r0, o1 - r0
            run.text = t[:a] + new[o0 - start:o1 - start] + t[b:]
        return True

    def _fix_reference(self, paragraph, index: int, location: str = "Reference"):
        """Fix reference formatting with structured change records."""
        reference_rules = self.rules.get("reference", {})
        changes = []
        allowed_properties = self._allowed_properties_for(
            index,
            categories=["references"],
            fallback=[
                "font_name",
                "font_size",
                "bold",
                "alignment",
                "line_spacing",
                "space_before",
                "space_after",
                "hanging_indent",
            ],
        )

        expected_font = reference_rules.get("font_name", "Times New Roman")
        expected_size = reference_rules.get("font_size", 9)
        expected_bold = reference_rules.get("bold", None)
        expected_alignment = reference_rules.get("alignment")
        expected_line_spacing = reference_rules.get("line_spacing")
        expected_space_before = reference_rules.get("space_before")
        expected_space_after = reference_rules.get("space_after")
        expected_hanging_indent = reference_rules.get("hanging_indent")

        # Fix every direct run AND every run nested inside a <w:hyperlink>
        # (e.g. a DOI link). paragraph.runs only returns direct <w:r>
        # children - a run properly wrapped in <w:hyperlink> (the correct,
        # valid OOXML structure) is invisible to it, so without this a
        # hyperlinked DOI silently keeps whatever font the "Hyperlink"
        # character style specifies (often a theme font like Cambria)
        # instead of the reference font every other run gets fixed to.
        runs_to_fix = list(paragraph.runs)
        for hyperlink in paragraph._p.findall(qn("w:hyperlink")):
            for run_element in hyperlink.findall(qn("w:r")):
                runs_to_fix.append(Run(run_element, paragraph))

        # Unconditional per-run font check - same reasoning as
        # _fix_body_text: the checker's paragraph-level "most common
        # run value" detection can miss a minority-mismatched run (e.g.
        # a DOI or a fragment left at a different size than the rest of
        # an otherwise-already-fixed reference entry).
        core_font_properties = {"font_name", "font_size", "bold"}
        run_allowed_properties = (
            set(allowed_properties) | core_font_properties
            if allowed_properties is not None
            else None
        )
        for run in runs_to_fix:
            if run.text.strip():
                changes.extend(self._fix_run_formatting(
                    run,
                    expected_font,
                    expected_size,
                    expected_bold,
                    allowed_properties=run_allowed_properties,
                ))

        current_alignment = get_paragraph_alignment(paragraph)
        if (
            expected_alignment
            and self._property_allowed("alignment", allowed_properties)
            and current_alignment != expected_alignment
        ):
            paragraph.alignment = WD_ALIGN_PARAGRAPH(
                ALIGNMENT_REVERSE_MAP.get(expected_alignment, 3)
            )
            changes.append({
                "property_name": "alignment",
                "current_value": current_alignment,
                "target_value": expected_alignment,
                "evidence": "Reference alignment did not match target rule",
            })

        if (
            expected_line_spacing is not None
            and self._property_allowed("line_spacing", allowed_properties)
        ):
            current_line_spacing = paragraph.paragraph_format.line_spacing
            current_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
            effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
            target_is_single = abs(float(expected_line_spacing) - 1.0) < 0.001
            # A tolerance of 0.05 let pdf2docx's "Multiple 0.97" (an actual
            # MULTIPLE rule, not true Single) through unnoticed since 0.97
            # is numerically close to 1.0 - close enough to pass the old
            # check, but still visibly "Multiple" rather than "Single" in
            # Word. Tightened the tolerance and, when the target is
            # single spacing, also check the rule itself rather than only
            # the numeric ratio.
            needs_fix = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
            if target_is_single and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
                needs_fix = True
            if needs_fix:
                paragraph.paragraph_format.line_spacing = expected_line_spacing
                changes.append({
                    "property_name": "line_spacing",
                    "current_value": (
                        f"{current_line_spacing_rule} ({current_line_spacing})"
                        if current_line_spacing is not None
                        else str(current_line_spacing_rule)
                    ),
                    "target_value": str(expected_line_spacing),
                    "evidence": "Reference line spacing did not match target rule",
                })

        if (
            expected_space_before is not None
            and self._property_allowed("space_before", allowed_properties)
        ):
            current_space_before = get_space_before_pt(paragraph)
            if (
                current_space_before is None
                or abs(float(current_space_before) - float(expected_space_before)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_before_pt(paragraph, float(expected_space_before))
                changes.append({
                    "property_name": "space_before",
                    "current_value": (
                        f"{current_space_before}pt"
                        if current_space_before is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_before}pt",
                    "evidence": "Reference paragraph spacing before did not match target rule",
                })

        if (
            expected_space_after is not None
            and self._property_allowed("space_after", allowed_properties)
        ):
            current_space_after = get_space_after_pt(paragraph)
            if (
                current_space_after is None
                or abs(float(current_space_after) - float(expected_space_after)) > 0.5
            or has_lines_based_spacing(paragraph)
            ):
                set_space_after_pt(paragraph, float(expected_space_after))
                changes.append({
                    "property_name": "space_after",
                    "current_value": (
                        f"{current_space_after}pt"
                        if current_space_after is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_space_after}pt",
                    "evidence": "Reference paragraph spacing after did not match target rule",
                })

        if (
            expected_hanging_indent is not None
            and self._property_allowed("hanging_indent", allowed_properties)
        ):
            # PDF conversion gives each reference paragraph its own,
            # inconsistent left indent (calculated from wherever that
            # specific line happened to sit on the original PDF page).
            # Hanging indent only controls the *relative* offset of the
            # first line via a negative first_line_indent - if the base
            # left_indent itself still varies reference-to-reference, both
            # the first line and the wrapped lines land at a different
            # horizontal position for every entry, which is exactly the
            # "references don't line up" misalignment this produces.
            # Reset the baseline so hanging indent is measured from the
            # same starting point for every reference. The correct
            # baseline for a hanging indent is the hanging amount itself,
            # not 0 - with left=0 and hanging=0.5, the math (first line =
            # left - hanging, wrapped lines = left) puts the first line at
            # -0.5" (sticking out past the margin) and wrapped lines at 0"
            # (no indent at all) - backwards, and exactly what produces a
            # reference list that doesn't line up with the rest of the
            # document. left = hanging = 0.5" is what actually gives
            # "first line flush at the margin, wrapped lines indented."
            target_left_indent = Inches(float(expected_hanging_indent))
            current_left_indent = paragraph.paragraph_format.left_indent
            if current_left_indent is None or current_left_indent != target_left_indent:
                paragraph.paragraph_format.left_indent = target_left_indent
                changes.append({
                    "property_name": "left_indent",
                    "current_value": (
                        f"{current_left_indent.inches:.2f}in"
                        if current_left_indent is not None
                        else "(inherited)"
                    ),
                    "target_value": f"{expected_hanging_indent:.2f}in",
                    "evidence": "Reference left indent was inherited from PDF conversion and varied between entries",
                })

            current_hanging_indent = get_hanging_indent_inches(paragraph)
            if (
                current_hanging_indent is None
                or abs(float(current_hanging_indent) - float(expected_hanging_indent)) > 0.03
            ):
                paragraph.paragraph_format.first_line_indent = Inches(-float(expected_hanging_indent))
                changes.append({
                    "property_name": "hanging_indent",
                    "current_value": (
                        f"{current_hanging_indent:.2f}in"
                        if current_hanging_indent is not None
                        else "No hanging indent"
                    ),
                    "target_value": f"{expected_hanging_indent}in",
                    "evidence": "Reference hanging indent did not match target rule",
                })

        publication_italic_required = reference_rules.get("publication_italic_required")
        if publication_italic_required is None:
            profile_name = str(self.rules.get("_profile", {}).get("name", "")).lower()
            publication_italic_required = profile_name == "jclc"

        if (
            publication_italic_required
            and (self._property_allowed("italic", allowed_properties)
                 or self._property_allowed("font_name", allowed_properties))
        ):
            if self._fix_reference_publication_italic(paragraph):
                changes.append({
                    "property_name": "publication_italic",
                    "current_value": "Journal/periodical name and volume not italicized",
                    "target_value": "Journal/periodical name and volume italicized",
                    "evidence": "APA7 requires the periodical name and volume to be italicized",
                })

        if changes:
            self._add_property_changes(
                paragraph_index=index,
                location=location,
                change_type="reference",
                details=changes,
                text_preview=truncate_text(get_paragraph_text(paragraph), 40),
                paragraph_type=ParagraphType.REFERENCE.value,
            )

    def _fix_sdt_references(self):
        """Fix reference paragraphs stored inside Word content controls."""
        for index, paragraph in enumerate(get_sdt_reference_paragraphs(self.document), start=1):
            self._fix_reference(
                paragraph,
                -1,
                location=f"Reference Content Control {index}",
            )

    def get_fixed_document_bytes(self) -> bytes:
        """Get the fixed document as bytes for download"""
        if not self.document:
            raise ValueError("No document to export")
        
        buffer = BytesIO()
        self.document.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()
    
    def get_changes_summary(self) -> Dict[str, Any]:
        """Get summary of all changes made"""
        summary = {
            "total_changes": len(self.changes),
            "changes_by_type": {},
            "changes_list": []
        }
        
        for change in self.changes:
            change_type = change.change_type
            summary["changes_by_type"][change_type] = \
                summary["changes_by_type"].get(change_type, 0) + 1
            
            summary["changes_list"].append({
                "location": change.location,
                "type": change.change_type,
                "before": change.before,
                "after": change.after,
                "property_name": change.property_name,
                "current_value": change.current_value,
                "target_value": change.target_value,
                "paragraph_type": change.paragraph_type,
                "evidence": change.evidence,
                "preview": change.text_preview
            })
        
        return summary
    
    def get_change_records(self) -> List[ChangeRecord]:
        """Get list of all change records"""
        return self.changes

    def _run_matches_change(self, run, change: ChangeRecord) -> bool:
        """Return True when a run likely contains the changed property."""
        property_name = change.property_name
        info = get_run_font_info(run)

        if property_name == "font_name":
            current_font = info.get("font_name")
            target_font = change.target_value
            return current_font is None or not is_font_equivalent(current_font, target_font)

        if property_name == "font_size":
            current_size = info.get("font_size")
            target_match = re.search(r"([\d.]+)", change.target_value)
            if not target_match:
                return True
            target_size = float(target_match.group(1))
            return current_size is None or abs(current_size - target_size) > 0.5

        if property_name == "bold":
            target_bold = change.target_value.lower() == "bold"
            return bool(info.get("bold")) != target_bold

        if property_name == "italic":
            target_italic = change.target_value.lower() == "italic"
            return bool(info.get("italic")) != target_italic

        return True

    def _highlight_page_header_change(self, document: Document, change: ChangeRecord) -> bool:
        """Highlight original page header text for document-level header changes."""
        if change.change_type != "page_header":
            return False

        highlighted = False
        for section in document.sections:
            headers = [
                section.header,
                section.first_page_header,
                section.even_page_header,
            ]
            for header in headers:
                for paragraph in header.paragraphs:
                    if not paragraph.text.strip():
                        continue
                    if not self._has_unstable_header_tabs(paragraph.text):
                        continue
                    for run in paragraph.runs:
                        if run.text.strip():
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            highlighted = True
        return highlighted

    def get_highlighted_document_bytes(self) -> bytes:
        """
        Get the original document with highlighted changed locations.
        Changed runs are highlighted in yellow without adding summary content.
        
        Returns:
            Document bytes with yellow highlighting on original changed locations
        """
        if not self.original_document:
            raise ValueError("No document to export")

        highlighted_document = deepcopy(self.original_document)

        for change in self.changes:
            index = change.paragraph_index
            if index < 0 or index >= len(highlighted_document.paragraphs):
                self._highlight_page_header_change(highlighted_document, change)
                continue

            paragraph = highlighted_document.paragraphs[index]
            matched_any_run = False
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                if self._run_matches_change(run, change):
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    matched_any_run = True

            if not matched_any_run:
                for run in paragraph.runs:
                    if run.text.strip():
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        buffer = BytesIO()
        highlighted_document.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()