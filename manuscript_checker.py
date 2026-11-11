"""
Manuscript Checker Module
Checks manuscript formatting against template rules
"""

from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from collections import Counter
import re
from difflib import SequenceMatcher

from .utils import (
    load_document, get_paragraph_text, get_paragraph_alignment, get_paragraph_font_info,
    get_sdt_reference_paragraphs, get_margins, get_line_spacing,
    get_space_after_pt, get_space_before_pt, get_hanging_indent_inches,
    get_hyperlink_run_font_names, has_run_level_size_mismatch,
    detect_reference_style, calculate_compliance_score, count_columns,
    truncate_text, is_font_equivalent, classify_author_info_role,
    to_journal_caption_title_case,
)
from .paragraph_classifier import ParagraphClassifier, ParagraphType, ClassifiedParagraph
from config import REQUIRED_SECTIONS

@dataclass
class FormatIssue:
    """Represents a formatting issue found in the manuscript"""
    category: str
    location: str
    paragraph_index: int
    description: str
    current_value: str
    expected_value: str
    severity: str = "warning"  # "error", "warning", "info"
    text_preview: str = ""

@dataclass
class CheckResult:
    """Result of manuscript check"""
    is_compliant: bool
    compliance_score: float
    total_issues: int
    issues_by_category: Dict[str, List[FormatIssue]]
    classifications: List[ClassifiedParagraph]
    document_structure: Dict[str, Any]
    statistics: Dict[str, Any]

class ManuscriptChecker:
    """
    Checks manuscript formatting against extracted template rules
    Performs 10 category checks and calculates compliance score
    """
    
    CATEGORIES = [
        "margins",
        "layout",
        "journal_header",
        "title",
        "author_info",
        "body_text",
        "headings",
        "structure",
        "tables",
        "figures",
        "references",
        "line_spacing",
        "other"
    ]

    REFERENCE_SOURCE_KEYWORDS = (
        "journal",
        "proceedings",
        "conference",
        "transactions",
        "communications",
        "informatics",
        "engineering",
        "springer",
        "elsevier",
        "ieee",
        "acm",
        "arxiv",
    )
    
    def __init__(self, rules: Dict[str, Any], llm_integration=None):
        """Initialize checker with template rules"""
        self.rules = rules
        self.llm = llm_integration
        self.document = None
        self.profile = rules.get("_profile", {})
        self.classifier = ParagraphClassifier(llm_integration, self.profile)
        self.issues: Dict[str, List[FormatIssue]] = {cat: [] for cat in self.CATEGORIES}
        self.classifications: List[ClassifiedParagraph] = []
        self.required_sections = self.profile.get("required_sections", REQUIRED_SECTIONS)
        self.required_declarations = self.profile.get("required_declarations", [])
    
    def load_manuscript(self, file_path_or_bytes):
        """Load the manuscript to check"""
        self.document = load_document(file_path_or_bytes)
        return self
    
    def check_all(self) -> CheckResult:
        """Perform all format checks"""
        if not self.document:
            raise ValueError("No manuscript loaded. Call load_manuscript() first.")
        
        # Reset issues
        self.issues = {cat: [] for cat in self.CATEGORIES}
        
        # Classify paragraphs first
        self.classifications = self.classifier.classify_document(self.document)
        
        # Run all checks
        self._check_margins()
        self._check_layout()
        self._check_journal_header()
        self._check_layout_stability()
        self._check_title()
        self._check_author_info()
        self._check_body_text()
        self._check_headings()
        structure = self._check_document_structure()
        self._check_tables()
        self._check_figures()
        self._check_equations()
        self._check_references()
        self._check_line_spacing()
        
        # Calculate compliance score
        score = calculate_compliance_score(self.issues)
        total_issues = sum(len(issues) for issues in self.issues.values())
        
        # Gather statistics
        stats = self._gather_statistics()
        
        return CheckResult(
            is_compliant=total_issues == 0,
            compliance_score=score,
            total_issues=total_issues,
            issues_by_category=self.issues,
            classifications=self.classifications,
            document_structure=structure,
            statistics=stats
        )
    
    def _add_issue(self, category: str, location: str, para_index: int,
                   description: str, current: str, expected: str,
                   severity: str = "warning", text_preview: str = ""):
        """Add a formatting issue"""
        issue = FormatIssue(
            category=category,
            location=location,
            paragraph_index=para_index,
            description=description,
            current_value=current,
            expected_value=expected,
            severity=severity,
            text_preview=text_preview
        )
        self.issues[category].append(issue)
    
    def _check_margins(self):
        """Check page margins against template"""
        current_margins = get_margins(self.document)
        expected_margins = self.rules.get("margins", {})
        
        margin_names = ["left", "right", "top", "bottom"]
        
        for margin in margin_names:
            current = current_margins.get(margin, 1.0)
            expected = expected_margins.get(margin, 1.0)
            
            # Allow small tolerance (0.05 inches)
            if abs(current - expected) > 0.05:
                self._add_issue(
                    category="margins",
                    location=f"{margin.capitalize()} Margin",
                    para_index=-1,
                    description=f"{margin.capitalize()} margin does not match template",
                    current=f"{current:.2f} inches",
                    expected=f"{expected:.2f} inches",
                    severity="error"
                )
    
    def _check_title(self):
        """Check paper title formatting"""
        title_rules = self.rules.get("title", {})
        found_title = False
        
        # Find the paper title
        for cp in self.classifications:
            if cp.paragraph_type == ParagraphType.PAPER_TITLE:
                found_title = True
                font_info = cp.font_info
                alignment = cp.alignment
                
                # Check font name (using font equivalence)
                expected_font = title_rules.get("font_name", "Times New Roman")
                current_font = font_info.get("font_name")
                if current_font and not is_font_equivalent(current_font, expected_font):
                    self._add_issue(
                        category="title",
                        location="Paper Title",
                        para_index=cp.index,
                        description="Title font does not match template",
                        current=current_font,
                        expected=expected_font,
                        severity="error",
                        text_preview=truncate_text(cp.text, 60)
                    )
                
                # Check font size (with 0.5pt tolerance)
                expected_size = title_rules.get("font_size", 24)
                current_size = font_info.get("font_size")
                if current_size and abs(current_size - expected_size) > 0.5:
                    self._add_issue(
                        category="title",
                        location="Paper Title",
                        para_index=cp.index,
                        description="Title font size does not match template",
                        current=f"{current_size}pt",
                        expected=f"{expected_size}pt",
                        severity="error",
                        text_preview=truncate_text(cp.text, 60)
                    )
                
                # Check bold
                expected_bold = title_rules.get("bold")
                current_bold = self._is_paragraph_mostly_bold(cp.index)
                if expected_bold is not None and current_bold != expected_bold:
                    self._add_issue(
                        category="title",
                        location="Paper Title",
                        para_index=cp.index,
                        description="Title bold formatting does not match template",
                        current="Bold" if current_bold else "Not Bold",
                        expected="Bold" if expected_bold else "Not Bold",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 60)
                    )

                expected_italic = title_rules.get("italic")
                current_italic = bool(font_info.get("italic"))
                if expected_italic is not None and current_italic != expected_italic:
                    self._add_issue(
                        category="title",
                        location="Paper Title",
                        para_index=cp.index,
                        description="Title italic formatting does not match template",
                        current="Italic" if current_italic else "Not Italic",
                        expected="Italic" if expected_italic else "Not Italic",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 60)
                    )

                # Check alignment
                expected_align = title_rules.get("alignment", "CENTER")
                if alignment != expected_align:
                    self._add_issue(
                        category="title",
                        location="Paper Title",
                        para_index=cp.index,
                        description="Title alignment does not match template",
                        current=alignment,
                        expected=expected_align,
                        severity="warning",
                        text_preview=truncate_text(cp.text, 60)
                    )              
                break  # Only check the first title

        if not found_title:
            self._add_issue(
                category="title",
                location="Paper Title",
                para_index=-1,
                description="Paper title could not be identified",
                current="Not found",
                expected="One clear paper title near the beginning of the manuscript",
                severity="error"
            )

    def _check_journal_header(self):
        """Check journal header formatting separately from the paper title."""
        header_rules = self.rules.get("journal_header", {})
        expected_font = header_rules.get("font_name", "Palatino Linotype")
        expected_size = header_rules.get("font_size", 24)
        expected_bold = header_rules.get("bold", True)
        expected_align = header_rules.get("alignment", "CENTER")

        header_paragraphs = [
            cp for cp in self.classifications
            if (
                cp.paragraph_type == ParagraphType.JOURNAL_HEADER
                and cp.index <= 1
                and ("journal of" in cp.text.lower() or "web engineering" in cp.text.lower())
            )
        ]

        if not header_paragraphs:
            self._add_issue(
                category="journal_header",
                location="Journal Header",
                para_index=-1,
                description="Journal header could not be identified",
                current="Not found",
                expected="Journal title/header near the top of the manuscript",
                severity="warning"
            )
            return

        for cp in header_paragraphs:
            font_info = cp.font_info
            location = f"Journal Header (Para {cp.index + 1})"
            preview = truncate_text(cp.text, 60)
            raw_text = self.document.paragraphs[cp.index].text

            current_font = font_info.get("font_name")
            if current_font and not is_font_equivalent(current_font, expected_font):
                self._add_issue(
                    category="journal_header",
                    location=location,
                    para_index=cp.index,
                    description="Journal header font does not match template",
                    current=current_font,
                    expected=expected_font,
                    severity="warning",
                    text_preview=preview
                )

            current_size = font_info.get("font_size")
            if current_size and abs(current_size - expected_size) > 0.5:
                self._add_issue(
                    category="journal_header",
                    location=location,
                    para_index=cp.index,
                    description="Journal header font size does not match template",
                    current=f"{current_size}pt",
                    expected=f"{expected_size}pt",
                    severity="warning",
                    text_preview=preview
                )

            current_bold = bool(font_info.get("bold"))
            if expected_bold is not None and current_bold != expected_bold:
                self._add_issue(
                    category="journal_header",
                    location=location,
                    para_index=cp.index,
                    description="Journal header bold formatting does not match template",
                    current="Bold" if current_bold else "Not Bold",
                    expected="Bold" if expected_bold else "Not Bold",
                    severity="warning",
                    text_preview=preview
                )

            if cp.alignment != expected_align:
                self._add_issue(
                    category="journal_header",
                    location=location,
                    para_index=cp.index,
                    description="Journal header alignment does not match template",
                    current=cp.alignment,
                    expected=expected_align,
                    severity="warning",
                    text_preview=preview
                )

            has_manual_indentation = raw_text.startswith((" ", "\t")) or "\t" in raw_text.strip()
            if has_manual_indentation:
                self._add_issue(
                    category="journal_header",
                    location=location,
                    para_index=cp.index,
                    description="Journal header contains manual tab indentation that can shift layout",
                    current="Manual tabs/spaces",
                    expected="No manual indentation before centered journal header text",
                    severity="warning",
                    text_preview=preview
                )

    def _check_layout_stability(self):
        """Check manual page header spacing patterns that commonly wrap in Word."""
        for section_index, section in enumerate(self.document.sections):
            header_parts = [
                ("Page Header", section.header),
                ("First Page Header", section.first_page_header),
                ("Even Page Header", section.even_page_header),
            ]
            for label, header in header_parts:
                for paragraph in header.paragraphs:
                    text = get_paragraph_text(paragraph)
                    if not text.strip():
                        continue
                    has_unstable_tabs = "\t\t" in text or re.search(r"\t\s{2,}", text)
                    if not has_unstable_tabs:
                        continue
                    self._add_issue(
                        category="other",
                        location=f"{label} (Section {section_index + 1})",
                        para_index=-1,
                        description="Page header uses multiple manual tabs/spaces that can wrap in Word",
                        current="Multiple manual tabs/spaces",
                        expected="Single right-aligned tab stop between left and right header text",
                        severity="warning",
                        text_preview=truncate_text(text, 80)
                    )
                    break

    def _check_author_info(self):
        """Check author names, affiliations, and corresponding-author details."""
        abstract_indices = [
            cp.index for cp in self.classifications
            if cp.paragraph_type in {ParagraphType.ABSTRACT_LABEL, ParagraphType.ABSTRACT_CONTENT}
        ]
        front_matter_end = min(abstract_indices) if abstract_indices else 25
        author_paragraphs = [
            cp for cp in self.classifications
            if cp.paragraph_type == ParagraphType.AUTHOR_INFO and cp.index < front_matter_end
        ]

        for cp in author_paragraphs:
            role = classify_author_info_role(cp.text, self.rules)
            rules = self.rules.get(role, self.rules.get("author", {}))
            expected_font = rules.get("font_name")
            expected_size = rules.get("font_size")
            expected_bold = rules.get("bold")
            expected_italic = rules.get("italic")
            expected_alignment = rules.get("alignment")
            mismatches = []

            current_font = cp.font_info.get("font_name")
            if expected_font and current_font and not is_font_equivalent(current_font, expected_font):
                mismatches.append(f"font {current_font}")
            current_size = cp.font_info.get("font_size")
            if expected_size is not None and current_size and abs(current_size - expected_size) > 0.5:
                mismatches.append(f"size {current_size}pt")
            current_bold = self._is_paragraph_mostly_bold(cp.index)
            if expected_bold is not None and current_bold != bool(expected_bold):
                mismatches.append("bold" if current_bold else "not bold")
            current_italic = bool(cp.font_info.get("italic"))
            if expected_italic is not None and current_italic != bool(expected_italic):
                mismatches.append("italic" if current_italic else "not italic")
            if expected_alignment and cp.alignment != expected_alignment:
                mismatches.append(f"alignment {cp.alignment}")

            # An author-number/corresponding-author marker ("1", "2",
            # "1*", "*") is meant to be true superscript formatting, not a
            # separately shrunk font size standing in for it. cp.font_info
            # reflects the paragraph's majority/representative run, so a
            # marker run that differs from the rest of the line is
            # otherwise invisible to this check entirely - checked
            # directly against the document's own runs here instead.
            paragraph = self.document.paragraphs[cp.index]
            for run in paragraph.runs:
                if not re.fullmatch(r"\d{1,2}\*?|\*", run.text.strip() or ""):
                    continue
                marker_size = run.font.size.pt if run.font.size else None
                size_mismatch = (
                    expected_size is not None
                    and marker_size is not None
                    and abs(marker_size - expected_size) > 0.5
                )
                if run.font.superscript is not True or size_mismatch:
                    mismatches.append("marker not true superscript")
                break

            if mismatches:
                role_label = {
                    "author": "Author name",
                    "affiliation": "Affiliation",
                    "corresponding_author": "Corresponding author",
                }[role]
                self._add_issue(
                    category="author_info",
                    location=role_label,
                    para_index=cp.index,
                    description=f"{role_label} formatting does not match template",
                    current=", ".join(mismatches),
                    expected=f"{expected_font or 'Template font'}, {expected_size or 'template'}pt",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 60),
                )

            if role == "corresponding_author":
                valid_email = re.search(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", cp.text)
                valid_orcid = re.search(
                    r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dXx]\b",
                    cp.text,
                )
                if rules.get("email_required") and not valid_email:
                    self._add_issue(
                        category="author_info",
                        location="Corresponding Author",
                        para_index=cp.index,
                        description="Corresponding author email is missing or invalid",
                        current="No valid email detected",
                        expected="A valid corresponding-author email address",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 60),
                    )
                if rules.get("orcid_required") and not valid_orcid:
                    self._add_issue(
                        category="author_info",
                        location="Corresponding Author",
                        para_index=cp.index,
                        description="Corresponding author ORCID is missing or invalid",
                        current="No valid ORCID detected",
                        expected="ORCID format: 0000-0000-0000-0000",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 60),
                    )

    def _check_layout(self):
        """Check page size, orientation, and column count."""
        layout_rules = self.rules.get("layout", {})
        expected_page_size = layout_rules.get("page_size")
        expected_orientation = layout_rules.get("orientation")
        expected_columns = layout_rules.get("columns")
        section = self.document.sections[0]
        width = section.page_width.inches
        height = section.page_height.inches
        short_side, long_side = sorted((width, height))

        if abs(short_side - 8.27) < 0.1 and abs(long_side - 11.69) < 0.1:
            current_page_size = "A4"
        elif abs(short_side - 8.5) < 0.1 and abs(long_side - 11.0) < 0.1:
            current_page_size = "Letter"
        else:
            current_page_size = f"{width:.2f}x{height:.2f}"

        current_orientation = "LANDSCAPE" if width > height else "PORTRAIT"
        current_columns = count_columns(self.document)

        if expected_page_size and current_page_size != expected_page_size:
            self._add_issue(
                "layout", "Page Layout", -1,
                "Page size does not match template",
                current_page_size, expected_page_size, "error",
            )
        if expected_orientation and current_orientation != expected_orientation:
            self._add_issue(
                "layout", "Page Layout", -1,
                "Page orientation does not match template",
                current_orientation, expected_orientation, "warning",
            )
        if expected_columns is not None and current_columns != int(expected_columns):
            self._add_issue(
                "layout", "Page Layout", -1,
                "Column count does not match template",
                str(current_columns), str(expected_columns), "warning",
            )
    
    def _check_body_text(self):
        """Check body text formatting - ENHANCED with tolerance"""
        body_rules = self.rules.get("body", {})
        expected_font = body_rules.get("font_name", "Times New Roman")
        expected_size = body_rules.get("font_size", 12)
        expected_bold = body_rules.get("bold")
        expected_alignment = body_rules.get("alignment")
        expected_space_after = body_rules.get("space_after")
        
        # Font size tolerance (allow +/-1.0pt difference for body text)
        size_tolerance = 0.5
        
        # Common font name variations that are equivalent
        font_equivalents = {
            "times new roman": ["times", "times-roman", "timesnewroman", "times new roman"],
            "arial": ["arial", "arial mt", "arialmt"],
            "calibri": ["calibri", "calibri light"],
        }
        
        def fonts_match(current, expected):
            if not current or not expected:
                return True  # Skip if font not detected
            current_lower = current.lower().strip()
            expected_lower = expected.lower().strip()
            if current_lower == expected_lower:
                return True
            # Check equivalents
            for base, variants in font_equivalents.items():
                if expected_lower in variants and current_lower in variants:
                    return True
            return False
        
        issues_found = 0
        max_issues_to_report = 10  # Limit to avoid too many issues
        alignment_mismatches = []
        space_after_mismatches = []
        space_before_mismatches = []
        expected_space_before = body_rules.get("space_before", 0)
        
        for cp in self.classifications:
            if cp.paragraph_type == ParagraphType.BODY:
                font_info = cp.font_info
                
                # Check font name
                current_font = font_info.get("font_name")
                if current_font and not fonts_match(current_font, expected_font):
                    if issues_found < max_issues_to_report:
                        self._add_issue(
                            category="body_text",
                            location=f"Paragraph {cp.index + 1}",
                            para_index=cp.index,
                            description="Body text font does not match template",
                            current=current_font,
                            expected=expected_font,
                            severity="warning",
                            text_preview=truncate_text(cp.text, 50)
                        )
                    issues_found += 1
                
                # Check font size with tolerance
                current_size = font_info.get("font_size")
                run_mismatch = has_run_level_size_mismatch(
                    self.document.paragraphs[cp.index], expected_size, size_tolerance
                )
                size_issue = (
                    (current_size and abs(current_size - expected_size) > size_tolerance)
                    or run_mismatch
                )
                if size_issue:
                    if issues_found < max_issues_to_report:
                        self._add_issue(
                            category="body_text",
                            location=f"Paragraph {cp.index + 1}",
                            para_index=cp.index,
                            description="Body text font size does not match template",
                            current=(
                                "Mixed sizes within paragraph"
                                if run_mismatch and current_size and abs(current_size - expected_size) <= size_tolerance
                                else f"{current_size}pt"
                            ),
                            expected=f"{expected_size}pt",
                            severity="warning",
                            text_preview=truncate_text(cp.text, 50)
                        )
                    issues_found += 1

                # Body text bold is intentionally not checked: bold in body
                # text is the author's own emphasis and is left as written
                # (see AutoFixer._fix_body_text), so reporting it as a
                # template mismatch would flag something nothing will fix.

                if expected_alignment and cp.alignment != expected_alignment:
                    alignment_mismatches.append(cp)

                if expected_space_after is not None:
                    current_space_after = get_space_after_pt(self.document.paragraphs[cp.index])
                    if current_space_after is None or abs(float(current_space_after) - float(expected_space_after)) > 0.05:
                        space_after_mismatches.append(cp)

                # A separate check from space_after above - a paragraph
                # can have a perfectly correct space_after while still
                # carrying a small leftover space_before artifact from
                # pdf2docx's own line measurements (commonly 0.4-0.5pt,
                # not exactly 0). Checking only space_after would leave
                # that paragraph with no flagged issue at all, which
                # blocks the fixer from ever touching its space_before -
                # this is a real gap, not just a defensive check: some
                # numbered-list-turned-real-numbering paragraphs were
                # observed with correct space_after but a stray nonzero
                # space_before that never got corrected as a result.
                if expected_space_before is not None:
                    current_space_before = get_space_before_pt(self.document.paragraphs[cp.index])
                    if current_space_before is None or abs(float(current_space_before) - float(expected_space_before)) > 0.05:
                        space_before_mismatches.append(cp)

        if alignment_mismatches:
            self._add_issue(
                category="body_text",
                location="Body Paragraphs",
                para_index=-1,
                description="Body text alignment does not match template",
                current=f"{len(alignment_mismatches)} non-{expected_alignment.lower()} paragraphs",
                expected=expected_alignment,
                severity="warning",
                text_preview=truncate_text(alignment_mismatches[0].text, 50)
            )

        if space_after_mismatches:
            current_space_after = get_space_after_pt(self.document.paragraphs[space_after_mismatches[0].index])
            self._add_issue(
                category="body_text",
                location="Body Paragraphs",
                para_index=-1,
                description="Body paragraph spacing after does not match template",
                current=(
                    f"{len(space_after_mismatches)} paragraphs use another after spacing"
                    if current_space_after is not None
                    else "Inherited or missing after spacing"
                ),
                expected=f"{expected_space_after}pt after paragraph spacing",
                severity="warning",
                text_preview=truncate_text(space_after_mismatches[0].text, 50)
            )

        if space_before_mismatches:
            current_space_before = get_space_before_pt(self.document.paragraphs[space_before_mismatches[0].index])
            self._add_issue(
                category="body_text",
                location="Body Paragraphs",
                para_index=-1,
                description="Body paragraph spacing before does not match template",
                current=(
                    f"{len(space_before_mismatches)} paragraphs use another before spacing"
                    if current_space_before is not None
                    else "Inherited or missing before spacing"
                ),
                expected=f"{expected_space_before}pt before paragraph spacing",
                severity="warning",
                text_preview=truncate_text(space_before_mismatches[0].text, 50)
            )
            # Also flagged per-paragraph (not just as one summary issue at
            # para_index=-1), since the fixer's own gating can be keyed to
            # a specific paragraph's own flagged issues rather than only
            # the document-wide category summary.
            for cp in space_before_mismatches:
                self._add_issue(
                    category="body_text",
                    location=f"Paragraph {cp.index + 1}",
                    para_index=cp.index,
                    description="Body paragraph spacing before does not match template",
                    current="Nonzero or missing before spacing",
                    expected=f"{expected_space_before}pt before paragraph spacing",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )
        
        # Add summary if many issues
        if issues_found > max_issues_to_report:
            self._add_issue(
                category="body_text",
                location="Multiple Paragraphs",
                para_index=-1,
                description=f"Additional {issues_found - max_issues_to_report} body text formatting issues found",
                current="Various",
                expected=f"{expected_font} {expected_size}pt",
                severity="info"
            )
        
        # Check abstract content formatting (usually smaller font like 9pt)
        self._check_abstract_content()
        
        # Check keywords content formatting (usually same as abstract)
        self._check_keywords_content()
    
    def _check_abstract_content(self):
        """Check abstract content formatting"""
        abstract_rules = self.rules.get("abstract", {})
        expected_font = abstract_rules.get("font_name", "Times New Roman")
        expected_size = abstract_rules.get("font_size", 9)
        expected_bold = abstract_rules.get("bold")
        expected_alignment = abstract_rules.get("alignment")
        minimum_words = abstract_rules.get("min_words")
        maximum_words = abstract_rules.get("max_words")
        one_paragraph = abstract_rules.get("one_paragraph")
        prohibit_equations = abstract_rules.get("prohibit_equations")
        prohibit_tables = abstract_rules.get("prohibit_tables")
        prohibit_citations = abstract_rules.get("prohibit_citations")
        size_tolerance = 0.5

        abstract_paragraphs = [
            cp for cp in self.classifications
            if cp.paragraph_type == ParagraphType.ABSTRACT_CONTENT
        ]
        for cp in abstract_paragraphs:
            font_info = cp.font_info

            current_font = font_info.get("font_name")
            if current_font and not is_font_equivalent(current_font, expected_font):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=cp.index,
                    description="Abstract font does not match template",
                    current=current_font,
                    expected=expected_font,
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            if expected_alignment and cp.alignment != expected_alignment:
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=cp.index,
                    description="Abstract alignment does not match template",
                    current=cp.alignment,
                    expected=expected_alignment,
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            current_size = font_info.get("font_size")
            run_mismatch = has_run_level_size_mismatch(self.document.paragraphs[cp.index], expected_size, size_tolerance)
            if (
                (current_size and abs(current_size - expected_size) > size_tolerance)
                or run_mismatch
            ):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=cp.index,
                    description="Abstract font size does not match template",
                    current=(
                        "Mixed sizes within paragraph"
                        if run_mismatch and current_size and abs(current_size - expected_size) <= size_tolerance
                        else f"{current_size}pt"
                    ),
                    expected=f"{expected_size}pt",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            current_bold = self._is_paragraph_mostly_bold(cp.index)
            if expected_bold is not None and current_bold != bool(expected_bold):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=cp.index,
                    description="Abstract bold formatting does not match template",
                    current="Bold" if current_bold else "Not Bold",
                    expected="Bold" if expected_bold else "Not Bold",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

        if abstract_paragraphs and (minimum_words is not None or maximum_words is not None):
            abstract_text = " ".join(cp.text for cp in abstract_paragraphs)
            abstract_text = re.sub(
                r"^\s*abstract\s*(?:-|\u2013|\u2014|:)?\s*",
                "",
                abstract_text,
                flags=re.IGNORECASE,
            )
            word_count = len(re.findall(r"\b[\w]+(?:[-'][\w]+)*\b", abstract_text))
            below_minimum = minimum_words is not None and word_count < int(minimum_words)
            above_maximum = maximum_words is not None and word_count > int(maximum_words)
            if below_minimum or above_maximum:
                if minimum_words is not None and maximum_words is not None:
                    expected_range = f"{minimum_words}-{maximum_words} words"
                elif minimum_words is not None:
                    expected_range = f"At least {minimum_words} words"
                else:
                    expected_range = f"No more than {maximum_words} words"
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=abstract_paragraphs[0].index,
                    description="Abstract word count is outside the template limit",
                    current=f"{word_count} words",
                    expected=expected_range,
                    severity="warning",
                    text_preview=truncate_text(abstract_text, 50)
                )

        if one_paragraph and len(abstract_paragraphs) > 1:
            self._add_issue(
                category="body_text",
                location="Abstract",
                para_index=abstract_paragraphs[0].index,
                description="Abstract must be written as one paragraph",
                current=f"{len(abstract_paragraphs)} paragraphs",
                expected="One paragraph",
                severity="warning",
                text_preview=truncate_text(" ".join(cp.text for cp in abstract_paragraphs), 50),
            )

        if abstract_paragraphs:
            abstract_text = " ".join(cp.text for cp in abstract_paragraphs)
            first_index = abstract_paragraphs[0].index
            if prohibit_citations and (
                re.search(r"\[\d+(?:\s*[-,]\s*\d+)*\]", abstract_text)
                or re.search(r"\([A-Z][A-Za-z-]+(?:\s+et\s+al\.)?,?\s+(?:19|20)\d{2}\)", abstract_text)
            ):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=first_index,
                    description="Abstract contains citation or reference markers",
                    current="Citation marker detected",
                    expected="No references or citations in the abstract",
                    severity="warning",
                    text_preview=truncate_text(abstract_text, 50),
                )

            has_math_xml = any(
                self.document.paragraphs[cp.index]._p.find(".//" + qn("m:oMath")) is not None
                or self.document.paragraphs[cp.index]._p.find(".//" + qn("m:oMathPara")) is not None
                for cp in abstract_paragraphs
            )
            has_equation_text = self._has_equation_like_text(abstract_text)
            if prohibit_equations and (has_math_xml or has_equation_text):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=first_index,
                    description="Abstract contains equation-like content",
                    current="Equation or formula detected",
                    expected="No mathematical equations in the abstract",
                    severity="warning",
                    text_preview=truncate_text(abstract_text, 50),
                )

            if prohibit_tables and re.search(r"\b(?:table|figure|fig\.)\s*\d+", abstract_text, re.IGNORECASE):
                self._add_issue(
                    category="body_text",
                    location="Abstract",
                    para_index=first_index,
                    description="Abstract contains table or figure material",
                    current="Table or figure reference detected",
                    expected="No tabular or figure material in the abstract",
                    severity="warning",
                    text_preview=truncate_text(abstract_text, 50),
                )

    @staticmethod
    def _has_equation_like_text(text: str) -> bool:
        """Return True for compact mathematical or statistical formula patterns."""
        patterns = [
            r"\b[\w\u0370-\u03ff]+\s*[=<>]\s*[.\w\u0370-\u03ff]+",
            r"\b[A-Za-z]\s*\([^)]*\)\s*=\s*-?\d+(?:\.\d+)?",
            r"\bp\s*[<>=]\s*\.?\d+",
            r"\b(?:r|r2|r\u00b2|f1|wer|mae|mse|rmse)\s*[=<>]\s*\.?\d+",
        ]
        return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)

    def _check_keywords_content(self):
        """Check keywords content formatting"""
        # Keywords typically use same format as abstract
        keywords_rules = self.rules.get("keywords", self.rules.get("abstract", {}))
        expected_font = keywords_rules.get("font_name", "Times New Roman")
        expected_size = keywords_rules.get("font_size", 9)
        expected_bold = keywords_rules.get("bold")
        expected_italic = keywords_rules.get("italic")
        minimum_count = keywords_rules.get("min_count")
        capitalize_first_letter = keywords_rules.get("capitalize_first_letter")
        size_tolerance = 0.5
        
        for cp in self.classifications:
            if cp.paragraph_type == ParagraphType.KEYWORDS_CONTENT:
                font_info = cp.font_info

                current_font = font_info.get("font_name")
                if current_font and not is_font_equivalent(current_font, expected_font):
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keywords font does not match template",
                        current=current_font,
                        expected=expected_font,
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                # Check font size
                current_size = font_info.get("font_size")
                if current_size and abs(current_size - expected_size) > size_tolerance:
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keywords font size does not match template",
                        current=f"{current_size}pt",
                        expected=f"{expected_size}pt",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )
                #  
                current_bold = self._is_paragraph_mostly_bold(cp.index)
                if expected_bold is not None and current_bold != bool(expected_bold):
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keywords bold formatting does not match template",
                        current="Bold" if current_bold else "Not Bold",
                        expected="Bold" if expected_bold else "Not Bold",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                current_italic = bool(font_info.get("italic"))
                if expected_italic is not None and current_italic != bool(expected_italic):
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keywords italic formatting does not match template",
                        current="Italic" if current_italic else "Not Italic",
                        expected="Italic" if expected_italic else "Not Italic",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                keyword_text = re.sub(
                    r"^\s*(?:keywords?|key\s+words?)\s*(?:-|\u2013|\u2014|:)?\s*",
                    "",
                    cp.text,
                    flags=re.IGNORECASE,
                )
                keyword_text = re.sub(r"\([^)]*\)\s*$", "", keyword_text).strip()
                keywords = [
                    item.strip()
                    for item in re.split(r"[,;]", keyword_text)
                    if item.strip()
                ]
                if minimum_count is not None and len(keywords) < int(minimum_count):
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keyword count is below the template minimum",
                        current=f"{len(keywords)} keywords",
                        expected=f"At least {minimum_count} keywords",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                invalid_capitalization = [
                    keyword for keyword in keywords
                    if keyword and keyword[0].isalpha() and not keyword[0].isupper()
                ]
                if capitalize_first_letter and invalid_capitalization:
                    self._add_issue(
                        category="body_text",
                        location="Keywords",
                        para_index=cp.index,
                        description="Keyword capitalization does not match template",
                        current=", ".join(invalid_capitalization[:5]),
                        expected="Capitalize the first letter of each keyword",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )
    
    def _check_headings(self):
        """Check section heading formatting"""
        # Non-numbered declaration-style headings (Acknowledgements,
        # Funding, References, Author Biography, ...) are consistent with
        # each other in practice even though the template doesn't specify
        # an explicit space-before value to check against directly.
        # Numbered section/subsection headings ("1.0 Introduction",
        # "4.1.2 Determiners") are excluded from this vote - they form
        # their own, considerably messier cluster of values that
        # shouldn't dilute (or be judged against) the declaration
        # headings' otherwise-consistent spacing.
        declaration_space_before_values = []
        for cp in self.classifications:
            if cp.paragraph_type != ParagraphType.SECTION_HEADING:
                continue
            if re.match(r"^\s*\d", cp.text or ""):
                continue
            if cp.text.strip().lower() == "abstract":
                continue
            space_before = self.document.paragraphs[cp.index].paragraph_format.space_before
            if space_before is not None:
                declaration_space_before_values.append(round(space_before.pt, 1))
        consistent_space_before = (
            Counter(declaration_space_before_values).most_common(1)[0][0]
            if declaration_space_before_values
            else None
        )

        for cp in self.classifications:
            if cp.paragraph_type in (ParagraphType.SECTION_HEADING, ParagraphType.ABSTRACT_LABEL):
                heading_rules = self._heading_rules_for_text(cp.text)
                expected_font = heading_rules.get("font_name", "Times New Roman")
                expected_size = heading_rules.get("font_size", 10)
                expected_bold = heading_rules.get("bold")
                expected_italic = heading_rules.get("italic")
                font_info = cp.font_info
                
                # Check font
                current_font = font_info.get("font_name")
                if current_font and not is_font_equivalent(current_font, expected_font):
                    self._add_issue(
                        category="headings",
                        location=f"Heading: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading font does not match template",
                        current=current_font,
                        expected=expected_font,
                        severity="warning",
                        text_preview=cp.text
                    )
                
                # Check size with 1.0pt tolerance
                current_size = font_info.get("font_size")
                if current_size and abs(current_size - expected_size) > 1.0:
                    self._add_issue(
                        category="headings",
                        location=f"Heading: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading font size does not match template",
                        current=f"{current_size}pt",
                        expected=f"{expected_size}pt",
                        severity="warning",
                        text_preview=cp.text
                    )
                
                # Check bold - handle None values (None means not bold)
                current_bold = bool(font_info.get("bold"))
                if expected_bold is not None and current_bold != bool(expected_bold):
                    self._add_issue(
                        category="headings",
                        location=f"Heading: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading bold formatting does not match template",
                        current="Bold" if current_bold else "Not Bold",
                        expected="Bold" if expected_bold else "Not Bold",
                        severity="warning",
                        text_preview=cp.text
                    )

                # Check space-before consistency, restricted to the same
                # non-numbered declaration-style headings the vote above
                # was computed from - comparing a numbered section heading
                # against this value would be comparing it to a different
                # convention entirely.
                # All headings (numbered sections, declaration headings,
                # Abstract) now share one explicit space-before rule
                # rather than needing a self-consistency vote among a
                # subset of them - heading_rules already resolves to the
                # right value for whichever kind cp.text is.
                expected_heading_space_before = heading_rules.get("space_before")
                if expected_heading_space_before is not None:
                    current_heading_space_before = get_space_before_pt(self.document.paragraphs[cp.index])
                    if (
                        current_heading_space_before is None
                        or abs(float(current_heading_space_before) - float(expected_heading_space_before)) > 1.0
                    ):
                        self._add_issue(
                            category="headings",
                            location=f"Heading: {truncate_text(cp.text, 30)}",
                            para_index=cp.index,
                            description="Heading spacing before did not match target rule",
                            current=(
                                f"{current_heading_space_before}pt"
                                if current_heading_space_before is not None
                                else "(inherited)"
                            ),
                            expected=f"{expected_heading_space_before}pt",
                            severity="warning",
                            text_preview=cp.text
                        )

                expected_heading_space_after = heading_rules.get("space_after")
                if expected_heading_space_after is not None:
                    current_heading_space_after = get_space_after_pt(self.document.paragraphs[cp.index])
                    if (
                        current_heading_space_after is None
                        or abs(float(current_heading_space_after) - float(expected_heading_space_after)) > 0.5
                    ):
                        self._add_issue(
                            category="headings",
                            location=f"Heading: {truncate_text(cp.text, 30)}",
                            para_index=cp.index,
                            description="Heading spacing after did not match target rule",
                            current=(
                                f"{current_heading_space_after}pt"
                                if current_heading_space_after is not None
                                else "(inherited)"
                            ),
                            expected=f"{expected_heading_space_after}pt",
                            severity="warning",
                            text_preview=cp.text
                        )

                expected_heading_line_spacing = heading_rules.get("line_spacing")
                if expected_heading_line_spacing is not None:
                    heading_paragraph = self.document.paragraphs[cp.index]
                    current_heading_line_spacing = heading_paragraph.paragraph_format.line_spacing
                    current_heading_line_spacing_rule = heading_paragraph.paragraph_format.line_spacing_rule
                    effective_heading_line_spacing = (
                        1.0 if current_heading_line_spacing is None else float(current_heading_line_spacing)
                    )
                    heading_target_is_single = abs(float(expected_heading_line_spacing) - 1.0) < 0.001
                    heading_needs_fix = (
                        abs(effective_heading_line_spacing - float(expected_heading_line_spacing)) > 0.01
                    )
                    if heading_target_is_single and current_heading_line_spacing_rule != WD_LINE_SPACING.SINGLE:
                        heading_needs_fix = True
                    if heading_needs_fix:
                        self._add_issue(
                            category="headings",
                            location=f"Heading: {truncate_text(cp.text, 30)}",
                            para_index=cp.index,
                            description="Heading line spacing did not match target rule",
                            current=(
                                f"{current_heading_line_spacing_rule} ({current_heading_line_spacing})"
                                if current_heading_line_spacing is not None
                                else str(current_heading_line_spacing_rule)
                            ),
                            expected=str(expected_heading_line_spacing),
                            severity="warning",
                            text_preview=cp.text
                        )

                current_italic = bool(font_info.get("italic"))
                if expected_italic is not None and current_italic != bool(expected_italic):
                    self._add_issue(
                        category="headings",
                        location=f"Heading: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading italic formatting does not match template",
                        current="Italic" if current_italic else "Not Italic",
                        expected="Italic" if expected_italic else "Not Italic",
                        severity="warning",
                        text_preview=cp.text
                    )

                expected_all_caps = heading_rules.get("all_caps")
                if expected_all_caps and not self._heading_text_is_all_caps(cp.text):
                    self._add_issue(
                        category="headings",
                        location=f"Heading: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading capitalization does not match template",
                        current=cp.text,
                        expected="All capital letters",
                        severity="warning",
                        text_preview=cp.text
                    )

                paragraph = self.document.paragraphs[cp.index]
                numbering_bold = self._get_numbering_bold(paragraph)
                if expected_bold is not None and numbering_bold is not None and numbering_bold != bool(expected_bold):
                    self._add_issue(
                        category="headings",
                        location=f"Heading Number: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Heading number bold formatting does not match template",
                        current="Bold" if numbering_bold else "Not Bold",
                        expected="Bold" if expected_bold else "Not Bold",
                        severity="warning",
                        text_preview=cp.text
                    )

    def _heading_rules_for_text(self, text: str) -> Dict[str, Any]:
        stripped = re.sub(r"\s+", " ", text.strip())
        # 主标题：1.0、2.0 等（小数部分严格为0；数字后的空格可有可无——
        # pdf2docx 或作者手打有时会漏掉空格，比如"1.0Introduction"，
        # 这仍然是主标题，不该因为没空格就被下面的子标题判断接住）
        if re.match(r"^\d+\.0(?!\d)\s*", stripped):
            return self.rules.get("heading", {})
        # 子标题：1.1、2.1.1 等
        if re.match(r"^\d+\.\d+", stripped):
            return self.rules.get("subheading", self.rules.get("heading", {}))
        base_rules = dict(self.rules.get("heading", {}))
        if stripped.lower() == "abstract":
            # Abstract is a one-off heading - it isn't part of either the
            # numbered-section cluster or the declaration-heading cluster,
            # so it shouldn't be pulled into either group's own spacing.
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
    
    def _check_document_structure(self) -> Dict[str, Any]:
        """Check required sections, order, and heading-role confidence."""
        sections = {
            section: {"found": False, "index": None, "format_status": "not_checked"}
            for section in self.required_sections
        }

        def normalize_heading(text: str) -> str:
            """Normalize a heading for structure matching.

            Examples:
            "2.0 Literature Review" -> "literature review"
            "3.0 Methods"           -> "methods"
            "4.0 Results"           -> "results"
            "5.0 Discussion"        -> "discussion"
            "6.0 Conclusion"        -> "conclusion"
            """
            normalized = re.sub(
                r"^\s*(?:\d+(?:\.\d+)*\.?|[IVXLC]+\.?|[A-Z]\.)\s+",
                "",
                text or "",
                flags=re.IGNORECASE
            )

            normalized = normalized.replace("&", "and")
            normalized = re.sub(r"\s+", " ", normalized)
            normalized = normalized.strip(" .:-–—\t")
            return normalized.lower()

        section_aliases = {

            "abstract": {
                "abstract"
            },

            "keywords": {
                "keywords",
                "keyword",
                "key words"
            },

            "introduction": {
                "introduction",
                "background"
            },

            "literature review": {
                "literature review",
                "literature",
                "related work",
                "related literature"
            },

            # -----------------------------------------------------
            # METHODS
            # -----------------------------------------------------
            "methods": {
                "methods",
                "method",
                "methodology",
                "research methodology",
                "research method",
                "materials and methods",
                "materials and method",
                "materials & methods",
                "experimental methods",
                "experimental methodology"
            },

            # -----------------------------------------------------
            # RESULTS
            # -----------------------------------------------------
            "results": {
                "results",
                "result",
                "findings",

                # Combined sections
                "results and discussion",
                "results & discussion",

                "discussion and results",
                "discussion & results",

                "findings and discussion",
                "findings & discussion",

                "discussion and findings",
                "discussion & findings",

                "results and findings",
                "results & findings",

                "findings and results",
                "findings & results"
            },

            # -----------------------------------------------------
            # DISCUSSION
            # -----------------------------------------------------
            "discussion": {
                "discussion",

                # Combined sections
                "results and discussion",
                "results & discussion",

                "discussion and results",
                "discussion & results",

                "findings and discussion",
                "findings & discussion",

                "discussion and findings",
                "discussion & findings"
            },

            # -----------------------------------------------------
            # CONCLUSION
            # -----------------------------------------------------
            "conclusion": {
                "conclusion",
                "conclusions",
                "conclusion and future work",
                "conclusions and future work",
                "summary and conclusion",
                "summary and conclusions"
            },

            # -----------------------------------------------------
            # REFERENCES
            # -----------------------------------------------------
            "references": {
                "references",
                "reference",
                "bibliography",
                "works cited"
            }
        }

        # Normalize aliases once
        normalized_aliases = {}

        for section, aliases in section_aliases.items():
            normalized_aliases[section] = {
                normalize_heading(alias)
                for alias in aliases
            }

        def heading_matches_section(
            heading: str,
            required_section: str
        ) -> bool:

            normalized_heading = normalize_heading(
                heading
            )

            normalized_required = normalize_heading(
                required_section
            )

            aliases = normalized_aliases.get(
                normalized_required,
                {normalized_required}
            )

            return normalized_heading in aliases

        def get_matching_sections(
            heading: str
        ) -> list:

            normalized_heading = normalize_heading(
                heading
            )

            matched_sections = []

            for required_section in self.required_sections:

                normalized_required = normalize_heading(
                    required_section
                )

                aliases = normalized_aliases.get(
                    normalized_required,
                    {normalized_required}
                )

                if normalized_heading in aliases:
                    matched_sections.append(
                        required_section
                    )

            return matched_sections

        def mark_section(section_name: str, cp: ClassifiedParagraph, format_status: str = "valid"):
            if section_name not in sections:
                return
            existing = sections[section_name]
            if existing["found"] and not (
                existing["format_status"] != "valid" and format_status == "valid"
            ):
                return
            sections[section_name] = {
                "found": True,
                "index": cp.index,
                "format_status": format_status,
            }

        for cp in self.classifications:
            text = cp.text.strip()
            text_lower = text.lower().strip()
            normalized_heading = normalize_heading(text)

            if not normalized_heading:
                continue

            if "abstract" in sections:
                if cp.paragraph_type in {
                    ParagraphType.ABSTRACT_LABEL,
                    ParagraphType.ABSTRACT_CONTENT
                }:
                    mark_section(
                        "abstract",
                        cp,
                        "valid"
                    )

            if "keywords" in sections:
                if cp.paragraph_type in {
                    ParagraphType.KEYWORDS_LABEL,
                    ParagraphType.KEYWORDS_CONTENT
                }:
                    mark_section(
                        "keywords",
                        cp,
                        "valid"
                    )

            if cp.paragraph_type == ParagraphType.SECTION_HEADING:
                matched_sections = get_matching_sections(
                    text
                )

                for matched_section in matched_sections:
                    mark_section(
                        matched_section,
                        cp,
                        "valid"
                    )

            if "references" in sections:

                if normalized_heading in {
                    "references",
                    "bibliography",
                    "works cited"
                }:
                    mark_section(
                        "references",
                        cp,
                        "valid"
                    )

            if "introduction" in sections and "introduction" in text_lower and cp.paragraph_type != ParagraphType.BODY:
                status = "valid" if cp.paragraph_type == ParagraphType.SECTION_HEADING else "weak"
                mark_section("introduction", cp, status)

            if "conclusion" in sections and "conclusion" in text_lower and cp.paragraph_type != ParagraphType.BODY:
                status = "valid" if cp.paragraph_type == ParagraphType.SECTION_HEADING else "weak"
                mark_section("conclusion", cp, status)

            if "references" in sections:
                text_lower.startswith(("references", "bibliography", "works cited"))
                status = "valid" if cp.paragraph_type == ParagraphType.SECTION_HEADING else "weak"
                mark_section("references", cp, status)
            elif cp.paragraph_type == ParagraphType.REFERENCE:
                mark_section("references", cp, "weak")

        self._check_required_heading_typos(sections)
        declarations = self._check_required_declarations()

        for section, details in sections.items():
            if not details["found"]:
                self._add_issue(
                    category="structure",
                    location="Document Structure",
                    para_index=-1,
                    description=f"Missing required section: {section.capitalize()}",
                    current="Not Found",
                    expected=f"{section.capitalize()} section",
                    severity="error"
                )
            elif details["format_status"] == "weak":
                self._add_issue(
                    category="structure",
                    location=f"{section.capitalize()} Section",
                    para_index=details["index"],
                    description="Section was found but its heading role is not confidently detected",
                    current="Weak section evidence",
                    expected="Recognized section heading or section label",
                    severity="warning"
                )

        found_positions = [
            sections[section]["index"]
            for section in self.required_sections
            if sections[section]["found"] and sections[section]["index"] is not None
        ]
        order_correct = found_positions == sorted(found_positions)
        if len(found_positions) > 1 and not order_correct:
            self._add_issue(
                category="structure",
                location="Document Section Order",
                para_index=-1,
                description="Required sections are not in the expected order",
                current="Out of order",
                expected=", ".join(section.title() for section in self.required_sections),
                severity="warning"
            )

        return {
            "sections": sections,
            "declarations": declarations,
            "order_correct": order_correct,
            "expected_order": self.required_sections,
        }

    def _check_required_declarations(self) -> Dict[str, Dict[str, Any]]:
        """Check mandatory JCLC declaration headings without generating content."""
        aliases = {
            "acknowledgement": ("acknowledgement", "acknowledgements"),
            "funding": ("funding", "funding statement"),
            "conflict_of_interest": (
                "conflict of interest",
                "conflict of interests",
                "competing interests",
                "competing interest",
            ),
            "author_contribution": (
                "author contribution",
                "author contributions",
                "authors contribution",
                "authors contributions",
                "author contribution statement",
                "authors contribution statement",
            ),
            "ethics_statement": (
                "ethics statement",
                "ethical statement",
                "ethics approval",
                "ethics declaration",
            ),
            "ai_declaration": (
                "ai declaration",
                "declaration of ai use",
                "declaration of generative ai use",
                "use of ai statement",
                "generative ai statement",
                "generative ai and ai assisted technologies statement",
            ),
            "data_access_statement": (
                "data access statement",
                "data availability statement",
                "data availability",
            ),
            "author_biography": (
                "author biography",
                "author biographies",
                "authors biography",
                "authors biographies",
            ),
        }
        results = {}
        normalized_paragraphs = []
        for cp in self.classifications:
            text = re.sub(r"\([^)]*\)", "", cp.text)
            text = re.sub(r"^\d+(?:\.\d+)*\.?\s+", "", text)
            normalized = re.sub(r"[^a-z ]", " ", text.lower())
            normalized = re.sub(r"\s+", " ", normalized).strip()
            normalized_paragraphs.append((cp, normalized))

        for declaration in self.required_declarations:
            found = None
            for cp, normalized in normalized_paragraphs:
                if normalized in aliases.get(declaration, (declaration.replace("_", " "),)):
                    found = cp
                    break
            results[declaration] = {
                "found": found is not None,
                "index": found.index if found else None,
            }
            if found is None:
                label = declaration.replace("_", " ").title()
                self._add_issue(
                    category="structure",
                    location="Declaration Sections",
                    para_index=-1,
                    description=f"Missing required declaration section: {label}",
                    current="Not Found",
                    expected=label,
                    severity="warning",
                )
        return results

    def _check_required_heading_typos(self, sections: Dict[str, Dict[str, Any]]) -> None:
        """Report likely misspellings of required section headings."""
        reported_sections = set()
        for cp in self.classifications:
            candidate = self._normalize_heading_candidate(cp.text)
            if not candidate:
                continue
            for section in self.required_sections:
                if sections.get(section, {}).get("found") or section in reported_sections:
                    continue
                ratio = SequenceMatcher(None, candidate, section).ratio()
                if ratio >= 0.82:
                    reported_sections.add(section)
                    self._add_issue(
                        category="structure",
                        location=f"Possible {section.capitalize()} Heading",
                        para_index=cp.index,
                        description="Possible misspelled required section heading",
                        current=cp.text,
                        expected=section.upper(),
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

    def _normalize_heading_candidate(self, text: str) -> str:
        """Normalize short heading-like text for typo detection."""
        stripped = re.sub(r"\s+", " ", text.strip())
        if not stripped or len(stripped) > 80:
            return ""
        if not (stripped.isupper() or re.match(r"^\d+(\.\d+)*\.?\s+\S+", stripped)):
            return ""
        stripped = re.sub(r"^\d+(\.\d+)*\.?\s+", "", stripped)
        stripped = re.sub(r"\([^)]*\)", "", stripped)
        normalized = re.sub(r"[^a-z ]", "", stripped.lower()).strip()
        if len(normalized.split()) > 4:
            return ""
        return normalized
    
    def _check_three_line_borders(self):
        from .table_borders import check_table
        for n, tbl in enumerate(self.document.element.body.findall(qn("w:tbl")), 1):
            probs = check_table(tbl)
            if probs:
                self._add_issue(
                    category="tables", location=f"Table {n}", para_index=-1,
                    description="Table must have exactly three horizontal lines (top, under header, bottom)",
                    current="; ".join(probs[:2]), expected="top, header and bottom lines only", severity="warning")

    def _check_tables(self):
        """Check tables in the document"""
        self._check_three_line_borders()
        tables = self.document.tables
        
        if not tables:
            # Check if tables are mentioned in text but not found
            for cp in self.classifications:
                if re.search(r'table\s*\d+', cp.text.lower()):
                    self._add_issue(
                        category="tables",
                        location="Tables",
                        para_index=-1,
                        description="Table reference found but no tables in document",
                        current="No tables",
                        expected="Tables matching references",
                        severity="warning"
                    )
                    break
        else:
            # Check table captions
            table_captions = [
                cp for cp in self.classifications
                if cp.paragraph_type == ParagraphType.CAPTION
                and re.match(r"^table\s*\d+\b", cp.text.lower())
            ]
            caption_count = len(table_captions)
            expected_caption_count = self._count_captionable_tables()
            
            if caption_count < expected_caption_count:
                self._add_issue(
                    category="tables",
                    location="Table Captions",
                    para_index=-1,
                    description="Some tables may be missing captions",
                    current=f"{caption_count} captions",
                    expected=f"{expected_caption_count} captions",
                    severity="warning"
                )

            caption_rules = self.rules.get("caption", {})
            expected_font = caption_rules.get("font_name", "Times New Roman")
            expected_size = caption_rules.get("font_size", 10)
            expected_italic = caption_rules.get("italic")
            expected_space_after = caption_rules.get("space_after")
            title_case_required = bool(caption_rules.get("title_case"))

            for cp in table_captions:
                current_font = cp.font_info.get("font_name")
                current_size = cp.font_info.get("font_size")
                current_italic = bool(cp.font_info.get("italic"))

                if current_font and not is_font_equivalent(current_font, expected_font):
                    self._add_issue(
                        category="tables",
                        location=f"Caption: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Table caption font does not match template",
                        current=current_font,
                        expected=expected_font,
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                if current_size and abs(current_size - expected_size) > 0.5:
                    self._add_issue(
                        category="tables",
                        location=f"Caption: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Table caption size does not match template",
                        current=f"{current_size}pt",
                        expected=f"{expected_size}pt",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                if expected_italic is not None and current_italic != bool(expected_italic):
                    self._add_issue(
                        category="tables",
                        location=f"Caption: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Table caption italic formatting does not match template",
                        current="Italic" if current_italic else "Not Italic",
                        expected="Italic" if expected_italic else "Not Italic",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

                if expected_space_after is not None:
                    current_space_after = get_space_after_pt(self.document.paragraphs[cp.index])
                    if (
                        current_space_after is None
                        or abs(float(current_space_after) - float(expected_space_after)) > 0.5
                    ):
                        self._add_issue(
                            category="tables",
                            location=f"Caption: {truncate_text(cp.text, 30)}",
                            para_index=cp.index,
                            description="Table caption paragraph spacing after does not match template",
                            current=(
                                f"{current_space_after}pt"
                                if current_space_after is not None
                                else "Inherited or missing after spacing"
                            ),
                            expected=f"{expected_space_after}pt",
                            severity="warning",
                            text_preview=truncate_text(cp.text, 50)
                        )

                if title_case_required:
                    expected_caption = to_journal_caption_title_case(cp.text)
                    if expected_caption != cp.text:
                        self._add_issue(
                            category="tables",
                            location=f"Caption: {truncate_text(cp.text, 30)}",
                            para_index=cp.index,
                            description="Table caption capitalization does not match template",
                            current=truncate_text(cp.text, 80),
                            expected=truncate_text(expected_caption, 80),
                            severity="warning",
                            text_preview=truncate_text(cp.text, 50)
                        )

        self._check_caption_order_and_numbering("table")
    
    def _check_figures(self):
        """Check figures in the document"""
        # Count figure captions
        figure_captions = [
            cp for cp in self.classifications
            if cp.paragraph_type == ParagraphType.CAPTION and
            re.match(r'^(figure|fig\.?)\s*\d+\b', cp.text.lower())
        ]
        
        image_count = self._count_body_images_before_back_matter()
        
        if image_count > 0 and len(figure_captions) < image_count:
            self._add_issue(
                category="figures",
                location="Figure Captions",
                para_index=-1,
                description="Some figures may be missing captions",
                current=f"{len(figure_captions)} captions",
                expected=f"{image_count} captions (for {image_count} images)",
                severity="warning"
            )

        # Check caption formatting
        caption_rules = self.rules.get("caption", {})
        expected_font = caption_rules.get("font_name", "Times New Roman")
        expected_size = caption_rules.get("font_size", 10)
        expected_italic = caption_rules.get("italic")
        expected_space_after = caption_rules.get("space_after")
        title_case_required = bool(caption_rules.get("title_case"))
        
        for cp in figure_captions:
            current_font = cp.font_info.get("font_name")
            current_size = cp.font_info.get("font_size")
            current_italic = bool(cp.font_info.get("italic"))
            
            if current_font and not is_font_equivalent(current_font, expected_font):
                self._add_issue(
                    category="figures",
                    location=f"Caption: {truncate_text(cp.text, 30)}",
                    para_index=cp.index,
                    description="Figure caption font does not match template",
                    current=current_font,
                    expected=expected_font,
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )
            
            if current_size and current_size != expected_size:
                self._add_issue(
                    category="figures",
                    location=f"Caption: {truncate_text(cp.text, 30)}",
                    para_index=cp.index,
                    description="Figure caption size does not match template",
                    current=f"{current_size}pt",
                    expected=f"{expected_size}pt",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            if expected_italic is not None and current_italic != bool(expected_italic):
                self._add_issue(
                    category="figures",
                    location=f"Caption: {truncate_text(cp.text, 30)}",
                    para_index=cp.index,
                    description="Figure caption italic formatting does not match template",
                    current="Italic" if current_italic else "Not Italic",
                    expected="Italic" if expected_italic else "Not Italic",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            if expected_space_after is not None:
                current_space_after = get_space_after_pt(self.document.paragraphs[cp.index])
                if (
                    current_space_after is None
                    or abs(float(current_space_after) - float(expected_space_after)) > 0.5
                ):
                    self._add_issue(
                        category="figures",
                        location=f"Caption: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Figure caption paragraph spacing after does not match template",
                        current=(
                            f"{current_space_after}pt"
                            if current_space_after is not None
                            else "Inherited or missing after spacing"
                        ),
                        expected=f"{expected_space_after}pt",
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

            if title_case_required:
                expected_caption = to_journal_caption_title_case(cp.text)
                if expected_caption != cp.text:
                    self._add_issue(
                        category="figures",
                        location=f"Caption: {truncate_text(cp.text, 30)}",
                        para_index=cp.index,
                        description="Figure caption capitalization does not match template",
                        current=truncate_text(cp.text, 80),
                        expected=truncate_text(expected_caption, 80),
                        severity="warning",
                        text_preview=truncate_text(cp.text, 50)
                    )

        self._check_caption_order_and_numbering("figure")
    
    def _check_references(self):
        """Check references section formatting"""
        reference_rules = self.rules.get("reference", {})
        expected_font = reference_rules.get("font_name", "Times New Roman")
        expected_size = reference_rules.get("font_size", 9)
        expected_bold = reference_rules.get("bold")
        expected_alignment = reference_rules.get("alignment")
        expected_line_spacing = reference_rules.get("line_spacing")
        expected_space_before = reference_rules.get("space_before")
        expected_space_after = reference_rules.get("space_after")
        expected_hanging_indent = reference_rules.get("hanging_indent")
        publication_italic_required = reference_rules.get("publication_italic_required")
        if publication_italic_required is None:
            publication_italic_required = self.profile.get("name", "").lower() == "jclc"
        
        references = [
            cp for cp in self.classifications
            if cp.paragraph_type == ParagraphType.REFERENCE
        ]
        
        if not references:
            sdt_references = get_sdt_reference_paragraphs(self.document)
            if sdt_references:
                self._check_ieee_reference_consistency()
                self._check_sdt_references(
                    sdt_references,
                    expected_font,
                    expected_size,
                    expected_bold,
                    expected_alignment,
                    expected_line_spacing,
                    expected_space_after,
                    expected_hanging_indent,
                    publication_italic_required,
                )
                return
            self._add_issue(
                category="references",
                location="References Section",
                para_index=-1,
                description="No references found in document",
                current="0 references",
                expected="At least 1 reference",
                severity="error"
            )
            return
        
        # Detect reference style
        ref_texts = [cp.text for cp in references]
        detected_style = detect_reference_style(ref_texts)
        self._check_ieee_reference_consistency()
        
        # Check for style consistency
        ieee_pattern = r'^\[\d+\]'
        apa_pattern = r'\(\d{4}\)'
        
        ieee_count = sum(1 for ref in ref_texts if re.match(ieee_pattern, ref.strip()))
        apa_count = sum(1 for ref in ref_texts if re.search(apa_pattern, ref))
        
        if ieee_count > 0 and apa_count > 0:
            self._add_issue(
                category="references",
                location="Reference Style",
                para_index=-1,
                description="Mixed reference styles detected (IEEE and APA)",
                current=f"IEEE: {ieee_count}, APA: {apa_count}",
                expected="Consistent style throughout",
                severity="warning"
            )
        
        # Check font formatting for every reference entry.
        alignment_mismatches = []
        line_spacing_mismatches = []
        space_before_mismatches = []
        space_after_mismatches = []
        hanging_indent_mismatches = []
        doi_font_mismatches = []
        for i, cp in enumerate(references):
            current_font = cp.font_info.get("font_name")
            current_size = cp.font_info.get("font_size")
            
            if current_font and not is_font_equivalent(current_font, expected_font):
                self._add_issue(
                    category="references",
                    location=f"Reference {i + 1}",
                    para_index=cp.index,
                    description="Reference font does not match template",
                    current=current_font,
                    expected=expected_font,
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            # A DOI/URL properly wrapped in <w:hyperlink> is invisible to
            # cp.font_info (built from paragraph.runs, which only sees
            # direct <w:r> children), so its font is otherwise never
            # checked at all. A hyperlink run commonly has no explicit
            # font at all (relying on Word's built-in "Hyperlink" style,
            # which may not even be defined in this document's own
            # styles.xml) - treat that as needing a fix in its own right,
            # not just an explicit-but-wrong font.
            hyperlink_fonts = get_hyperlink_run_font_names(self.document.paragraphs[cp.index])
            if any(
                font is None or not is_font_equivalent(font, expected_font)
                for font in hyperlink_fonts
            ):
                doi_font_mismatches.append(cp)

            run_mismatch = has_run_level_size_mismatch(self.document.paragraphs[cp.index], expected_size, 0.5)
            if (
                (current_size and current_size != expected_size)
                or run_mismatch
            ):
                self._add_issue(
                    category="references",
                    location=f"Reference {i + 1}",
                    para_index=cp.index,
                    description="Reference font size does not match template",
                    current=(
                        "Mixed sizes within paragraph"
                        if run_mismatch and current_size == expected_size
                        else f"{current_size}pt"
                    ),
                    expected=f"{expected_size}pt",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            current_bold = self._is_paragraph_mostly_bold(cp.index)
            if expected_bold is not None and current_bold != bool(expected_bold):
                self._add_issue(
                    category="references",
                    location=f"Reference {i + 1}",
                    para_index=cp.index,
                    description="Reference bold formatting does not match template",
                    current="Bold" if current_bold else "Not Bold",
                    expected="Bold" if expected_bold else "Not Bold",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

            if expected_alignment and cp.alignment != expected_alignment:
                alignment_mismatches.append(cp)

            if expected_line_spacing is not None:
                current_line_spacing = get_line_spacing(self.document.paragraphs[cp.index])
                current_line_spacing_rule = self.document.paragraphs[cp.index].paragraph_format.line_spacing_rule
                effective_line_spacing = 1.0 if current_line_spacing is None else float(current_line_spacing)
                target_is_single = abs(float(expected_line_spacing) - 1.0) < 0.001
                # A tolerance of 0.05 let pdf2docx's "Multiple 0.97" - a
                # genuine MULTIPLE rule, not true Single - through
                # unnoticed, since 0.97 is numerically close enough to 1.0
                # to pass. Tightened the tolerance and, when the target is
                # single spacing, also check the rule itself so a
                # near-1.0 MULTIPLE value still gets flagged.
                is_mismatch = abs(effective_line_spacing - float(expected_line_spacing)) > 0.01
                if target_is_single and current_line_spacing_rule != WD_LINE_SPACING.SINGLE:
                    is_mismatch = True
                if is_mismatch:
                    line_spacing_mismatches.append(cp)

            if expected_space_before is not None:
                current_space_before = get_space_before_pt(self.document.paragraphs[cp.index])
                if (
                    current_space_before is None
                    or abs(float(current_space_before) - float(expected_space_before)) > 0.5
                ):
                    space_before_mismatches.append(cp)

            if expected_space_after is not None:
                current_space_after = get_space_after_pt(self.document.paragraphs[cp.index])
                if (
                    current_space_after is None
                    or abs(float(current_space_after) - float(expected_space_after)) > 0.5
                ):
                    space_after_mismatches.append(cp)

            if expected_hanging_indent is not None:
                current_hanging_indent = get_hanging_indent_inches(self.document.paragraphs[cp.index])
                if (
                    current_hanging_indent is None
                    or abs(float(current_hanging_indent) - float(expected_hanging_indent)) > 0.03
                ):
                    hanging_indent_mismatches.append(cp)

            if (
                publication_italic_required
                and self._reference_likely_has_publication_source(cp.text)
                and not self._paragraph_has_italic_publication_source(cp.index)
            ):
                self._add_issue(
                    category="references",
                    location=f"Reference {i + 1}",
                    para_index=cp.index,
                    description="Reference publication source may need italic formatting",
                    current="No italic publication source segment detected",
                    expected="Italic journal, conference, book, or proceedings source segment",
                    severity="warning",
                    text_preview=truncate_text(cp.text, 50)
                )

        if alignment_mismatches:
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference alignment does not match template",
                current=f"{len(alignment_mismatches)} non-{expected_alignment.lower()} references",
                expected=expected_alignment,
                severity="warning",
                text_preview=truncate_text(alignment_mismatches[0].text, 50)
            )

        if line_spacing_mismatches:
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference line spacing does not match template",
                current=f"{len(line_spacing_mismatches)} references use another spacing",
                expected=str(expected_line_spacing),
                severity="warning",
                text_preview=truncate_text(line_spacing_mismatches[0].text, 50)
            )

        if space_before_mismatches:
            current_space_before = get_space_before_pt(self.document.paragraphs[space_before_mismatches[0].index])
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference paragraph spacing before does not match template",
                current=(
                    f"{len(space_before_mismatches)} references use another before spacing"
                    if current_space_before is not None
                    else "Inherited or missing before spacing"
                ),
                expected=f"{expected_space_before}pt",
                severity="warning",
                text_preview=truncate_text(space_before_mismatches[0].text, 50)
            )

        if space_after_mismatches:
            current_space_after = get_space_after_pt(self.document.paragraphs[space_after_mismatches[0].index])
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference paragraph spacing after does not match template",
                current=(
                    f"{len(space_after_mismatches)} references use another after spacing"
                    if current_space_after is not None
                    else "Inherited or missing after spacing"
                ),
                expected=f"{expected_space_after}pt",
                severity="warning",
                text_preview=truncate_text(space_after_mismatches[0].text, 50)
            )

        if doi_font_mismatches:
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference DOI/hyperlink font does not match template",
                current=f"{len(doi_font_mismatches)} references have a mismatched DOI/hyperlink font",
                expected=expected_font,
                severity="warning",
                text_preview=truncate_text(doi_font_mismatches[0].text, 50)
            )

        if hanging_indent_mismatches:
            current_hanging_indent = get_hanging_indent_inches(
                self.document.paragraphs[hanging_indent_mismatches[0].index]
            )
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference hanging indent does not match template",
                current=(
                    f"{current_hanging_indent:.2f}in"
                    if current_hanging_indent is not None
                    else "No hanging indent"
                ),
                expected=f"{expected_hanging_indent}in",
                severity="warning",
                text_preview=truncate_text(hanging_indent_mismatches[0].text, 50)
            )

    def _check_equations(self) -> None:
        """Check display-equation numbering and explicit equation citations."""
        equation_numbers = []
        unnumbered_count = 0
        citation_numbers = set()
        in_references = False
        paragraphs = [
            paragraph
            for paragraph in self.document.element.body.iter()
            if paragraph.tag == qn("w:p")
        ]

        for index, paragraph in enumerate(paragraphs):
            text = self._get_equation_paragraph_text(paragraph)
            normalized = re.sub(r"\s+", " ", text).strip()
            lower_text = normalized.lower()
            if lower_text.startswith(("references", "bibliography", "works cited")):
                in_references = True

            has_display_math = any(
                node.tag == qn("m:oMathPara")
                for node in paragraph.iter()
            )
            has_math = any(
                node.tag == qn("m:oMath")
                for node in paragraph.iter()
            )
            if has_math:
                math_text = "".join(
                    text_node.text or ""
                    for text_node in paragraph.iter()
                    if text_node.tag == qn("m:t")
                )
                number = self._extract_equation_label(math_text)
                if number is None:
                    number = self._extract_equation_label(normalized)
                if number is None and has_display_math:
                    row = paragraph.getparent()
                    while row is not None and row.tag not in {qn("w:tr"), qn("w:body")}:
                        row = row.getparent()
                    if row is not None and row.tag == qn("w:tr"):
                        row_text = "".join(
                            text_node.text or ""
                            for text_node in row.iter()
                            if text_node.tag in {qn("w:t"), qn("m:t")}
                        )
                        number = self._extract_equation_label(row_text)
                if number is None and has_display_math:
                    for adjacent_index in (index - 1, index + 1):
                        if not 0 <= adjacent_index < len(paragraphs):
                            continue
                        if any(
                            node.tag == qn("m:oMath")
                            for node in paragraphs[adjacent_index].iter()
                        ):
                            continue
                        adjacent_text = self._get_equation_paragraph_text(
                            paragraphs[adjacent_index]
                        )
                        adjacent_match = re.fullmatch(
                            r"\s*\((\d+)\)\s*",
                            adjacent_text,
                        )
                        if adjacent_match:
                            number = int(adjacent_match.group(1))
                            break
                if number is not None:
                    equation_numbers.append(number)
                elif has_display_math:
                    unnumbered_count += 1

            if not in_references:
                citation_numbers.update(self._extract_equation_citation_numbers(normalized))

        if unnumbered_count and equation_numbers and citation_numbers:
            self._add_issue(
                category="other",
                location="Equation Numbering",
                para_index=-1,
                description="Display equation may be missing a number",
                current=f"{unnumbered_count} unnumbered display equations",
                expected="Numbered display equations such as (1), (2), and (3)",
                severity="warning",
            )

        equation_numbers = list(dict.fromkeys(equation_numbers))
        expected_numbers = list(range(1, len(equation_numbers) + 1))
        if equation_numbers and equation_numbers != expected_numbers:
            self._add_issue(
                category="other",
                location="Equation Numbering",
                para_index=-1,
                description="Equation numbering is not continuous",
                current=", ".join(f"({number})" for number in equation_numbers),
                expected=", ".join(f"({number})" for number in expected_numbers),
                severity="warning",
            )

        missing_equations = sorted(citation_numbers - set(equation_numbers))
        if equation_numbers and missing_equations:
            self._add_issue(
                category="other",
                location="Equation Citations",
                para_index=-1,
                description="Equation citation has no matching equation",
                current=", ".join(f"({number})" for number in missing_equations),
                expected="Every explicit equation citation matches a numbered equation",
                severity="error",
            )

    @staticmethod
    def _get_equation_paragraph_text(paragraph) -> str:
        """Return visible Word and OMML text from a paragraph."""
        return "".join(
            text_node.text or ""
            for text_node in paragraph.iter()
            if text_node.tag in {qn("w:t"), qn("m:t")}
        )

    @staticmethod
    def _extract_equation_label(text: str) -> Optional[int]:
        """Extract a trailing equation label such as (3) or Word's #3 form."""
        match = re.search(r"#\s*(\d+)", text)
        if match:
            return int(match.group(1))
        match = re.search(r"^\s*\((\d+)\)|\((\d+)\)\s*$", text)
        if not match:
            return None
        return int(match.group(1) or match.group(2))

    @staticmethod
    def _extract_equation_citation_numbers(text: str) -> set:
        """Extract numbers from explicit Equation or Eq. citations."""
        numbers = set()
        range_pattern = (
            r"\b(?:equations?|eqs?\.?)\s*\(?(\d+)\)?"
            r"\s*[-\u2013\u2014]\s*\(?(\d+)\)?"
        )
        for match in re.finditer(range_pattern, text, re.IGNORECASE):
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                numbers.update(range(start, end + 1))

        single_pattern = r"\b(?:equations?|eqs?\.?)\s*\(?(\d+)\)?"
        numbers.update(
            int(match.group(1))
            for match in re.finditer(single_pattern, text, re.IGNORECASE)
        )
        return numbers

    def _check_sdt_references(
        self,
        references,
        expected_font,
        expected_size,
        expected_bold,
        expected_alignment,
        expected_line_spacing,
        expected_space_after,
        expected_hanging_indent,
        publication_italic_required,
    ):
        """Check references stored inside Word content controls."""
        alignment_mismatches = []
        line_spacing_mismatches = []
        space_after_mismatches = []
        hanging_indent_mismatches = []

        for index, paragraph in enumerate(references):
            text = get_paragraph_text(paragraph)
            font_info = get_paragraph_font_info(paragraph)
            current_font = font_info.get("font_name")
            current_size = font_info.get("font_size")
            current_bold = bool(font_info.get("bold"))

            if current_font and not is_font_equivalent(current_font, expected_font):
                self._add_issue(
                    category="references",
                    location=f"Reference Content Control {index + 1}",
                    para_index=-1,
                    description="Reference font does not match template",
                    current=current_font,
                    expected=expected_font,
                    severity="warning",
                    text_preview=truncate_text(text, 50),
                )

            if current_size and current_size != expected_size:
                self._add_issue(
                    category="references",
                    location=f"Reference Content Control {index + 1}",
                    para_index=-1,
                    description="Reference font size does not match template",
                    current=f"{current_size}pt",
                    expected=f"{expected_size}pt",
                    severity="warning",
                    text_preview=truncate_text(text, 50),
                )

            if expected_bold is not None and current_bold != bool(expected_bold):
                self._add_issue(
                    category="references",
                    location=f"Reference Content Control {index + 1}",
                    para_index=-1,
                    description="Reference bold formatting does not match template",
                    current="Bold" if current_bold else "Not Bold",
                    expected="Bold" if expected_bold else "Not Bold",
                    severity="warning",
                    text_preview=truncate_text(text, 50),
                )

            current_alignment = get_paragraph_alignment(paragraph)
            if expected_alignment and current_alignment != expected_alignment:
                alignment_mismatches.append(paragraph)

            if expected_line_spacing is not None:
                current_spacing = get_line_spacing(paragraph)
                effective_spacing = 1.0 if current_spacing is None else float(current_spacing)
                if abs(effective_spacing - float(expected_line_spacing)) > 0.05:
                    line_spacing_mismatches.append(paragraph)

            if expected_space_after is not None:
                current_space_after = get_space_after_pt(paragraph)
                if (
                    current_space_after is None
                    or abs(float(current_space_after) - float(expected_space_after)) > 0.5
                ):
                    space_after_mismatches.append(paragraph)

            if expected_hanging_indent is not None:
                current_hanging_indent = get_hanging_indent_inches(paragraph)
                if (
                    current_hanging_indent is None
                    or abs(float(current_hanging_indent) - float(expected_hanging_indent)) > 0.03
                ):
                    hanging_indent_mismatches.append(paragraph)

            if (
                publication_italic_required
                and self._reference_likely_has_publication_source(text)
                and not self._paragraph_object_has_italic_publication_source(paragraph)
            ):
                self._add_issue(
                    category="references",
                    location=f"Reference Content Control {index + 1}",
                    para_index=-1,
                    description="Reference publication source may need italic formatting",
                    current="No italic publication source segment detected",
                    expected="Italic journal, conference, book, or proceedings source segment",
                    severity="warning",
                    text_preview=truncate_text(text, 50),
                )

        if alignment_mismatches:
            self._add_issue(
                category="references",
                location="Reference Content Controls",
                para_index=-1,
                description="Reference alignment does not match template",
                current=f"{len(alignment_mismatches)} non-{expected_alignment.lower()} reference paragraphs",
                expected=expected_alignment,
                severity="warning",
                text_preview=truncate_text(get_paragraph_text(alignment_mismatches[0]), 50),
            )

        if line_spacing_mismatches:
            self._add_issue(
                category="references",
                location="Reference Content Controls",
                para_index=-1,
                description="Reference line spacing does not match template",
                current=f"{len(line_spacing_mismatches)} references use another spacing",
                expected=str(expected_line_spacing),
                severity="warning",
                text_preview=truncate_text(get_paragraph_text(line_spacing_mismatches[0]), 50),
            )

        if space_after_mismatches:
            current_space_after = get_space_after_pt(space_after_mismatches[0])
            self._add_issue(
                category="references",
                location="Reference Content Controls",
                para_index=-1,
                description="Reference paragraph spacing after does not match template",
                current=(
                    f"{len(space_after_mismatches)} references use another after spacing"
                    if current_space_after is not None
                    else "Inherited or missing after spacing"
                ),
                expected=f"{expected_space_after}pt",
                severity="warning",
                text_preview=truncate_text(get_paragraph_text(space_after_mismatches[0]), 50),
            )

        if hanging_indent_mismatches:
            current_hanging_indent = get_hanging_indent_inches(hanging_indent_mismatches[0])
            self._add_issue(
                category="references",
                location="Reference Content Controls",
                para_index=-1,
                description="Reference hanging indent does not match template",
                current=(
                    f"{current_hanging_indent:.2f}in"
                    if current_hanging_indent is not None
                    else "No hanging indent"
                ),
                expected=f"{expected_hanging_indent}in",
                severity="warning",
                text_preview=truncate_text(get_paragraph_text(hanging_indent_mismatches[0]), 50),
            )

    def _is_paragraph_mostly_bold(self, paragraph_index: int) -> bool:
        """Return True when most visible paragraph text is explicitly bold."""
        if paragraph_index < 0 or paragraph_index >= len(self.document.paragraphs):
            return False
        bold_chars = 0
        total_chars = 0
        for run in self.document.paragraphs[paragraph_index].runs:
            text_length = len(run.text.strip())
            if not text_length:
                continue
            total_chars += text_length
            if bool(run.font.bold):
                bold_chars += text_length
        return total_chars > 0 and (bold_chars / total_chars) >= 0.8

    def _paragraph_has_italic_text(self, paragraph_index: int) -> bool:
        """Return True when a paragraph contains any explicitly italic visible text."""
        if paragraph_index < 0 or paragraph_index >= len(self.document.paragraphs):
            return False
        for run in self.document.paragraphs[paragraph_index].runs:
            if run.text.strip() and bool(run.font.italic):
                return True
        return False

    def _paragraph_has_italic_publication_source(self, paragraph_index: int) -> bool:
        """Return True when italic text looks like a publication source."""
        if paragraph_index < 0 or paragraph_index >= len(self.document.paragraphs):
            return False
        return self._paragraph_object_has_italic_publication_source(
            self.document.paragraphs[paragraph_index],
        )

    def _paragraph_object_has_italic_publication_source(self, paragraph) -> bool:
        """Return True when a paragraph has an italic source-like run."""
        full_text = "".join(run.text or "" for run in paragraph.runs)
        span = self._reference_source_span(re.sub(r"\s+", " ", full_text)) if full_text == re.sub(r"\s+", " ", full_text) else self._reference_source_span(full_text)
        if span is not None:
            # structural journal+volume span found: it must be italic (journal name and volume)
            cur = 0
            for run in paragraph.runs:
                t = run.text or ""
                a, b = cur, cur + len(t)
                cur = b
                if t.strip() and max(a, span[0]) < min(b, span[1]) and not run.font.italic:
                    return False
            return True
        cursor = 0
        for run in paragraph.runs:
            run_text = run.text or ""
            start = cursor
            end = cursor + len(run_text)
            cursor = end
            if (
                run_text.strip()
                and bool(run.font.italic)
                and (
                    self._text_looks_like_publication_source(run_text)
                    or self._span_looks_like_publication_source_zone(
                        full_text,
                        start,
                        end,
                    )
                )
            ):
                return True
        return False

    def _text_looks_like_publication_source(self, text: str) -> bool:
        """Return True when text resembles a journal or proceedings source."""
        normalized = re.sub(r"\s+", " ", text or "").strip().lower()
        if len(normalized) < 5:
            return False
        return any(
            keyword in normalized
            for keyword in self.REFERENCE_SOURCE_KEYWORDS
        )

    @staticmethod
    def _span_looks_like_publication_source_zone(
        full_text: str,
        start: int,
        end: int,
    ) -> bool:
        """Return True when an italic span is in the usual IEEE source position."""
        prefix = full_text[:start]
        suffix = full_text[end:]
        has_title_before = bool(re.search(r"[\"“”].+[\"“”]\s*,?\s*$", prefix))
        has_reference_detail_after = bool(
            re.search(
                r"\b(vol\.|volume|no\.|pp\.|pages?|doi|https?://)\b|\b\d{4}\b",
                suffix,
                re.IGNORECASE,
            )
        )
        return has_title_before and has_reference_detail_after

    @staticmethod
    def _reference_source_span(text: str):
        """(start, end, source) of the 'Journal Name, Volume' part of an APA reference, or None."""
        from .auto_fixer import _REFERENCE_YEAR_RE, _REFERENCE_SOURCE_RE
        y = _REFERENCE_YEAR_RE.search(text)
        m = _REFERENCE_SOURCE_RE.search(text, y.end() if y else 0)
        if m is None:
            return None
        return m.start("source"), m.end("volume"), m.group("source")

    def _reference_likely_has_publication_source(self, text: str) -> bool:
        """Return True for references likely to contain an italic publication source."""
        normalized = re.sub(r"\s+", " ", text or "").strip()
        if len(normalized) < 80:
            return False

        if self._reference_source_span(normalized) is not None:
            return True

        lower_text = normalized.lower()
        detail_patterns = (
            r"\bvol\.",
            r"\bvolume\b",
            r"\bno\.",
            r"\bpp\.",
            r"\bdoi\b",
            r"\(\d{4}\)",
            r"\".+?\"",
            r"\u201c.+?\u201d",
        )

        has_source_keyword = any(
            keyword in lower_text
            for keyword in self.REFERENCE_SOURCE_KEYWORDS
        )
        has_bibliographic_detail = any(re.search(pattern, normalized, re.IGNORECASE) for pattern in detail_patterns)
        return has_source_keyword and has_bibliographic_detail

    def _count_captionable_tables(self) -> int:
        """Count tables that are likely manuscript tables needing captions."""
        count = 0
        for table in self.document.tables:
            text = "\n".join(cell.text.strip() for row in table.rows for cell in row.cells)
            normalized = re.sub(r"\s+", " ", text.lower()).strip()
            first_cell = ""
            if table.rows and table.rows[0].cells:
                first_cell = table.rows[0].cells[0].text.strip().lower()

            if not normalized:
                continue
            if normalized.startswith("received:") and "published:" in normalized:
                continue
            if first_cell.startswith("algorithm"):
                continue
            if " is a " in normalized and "research" in normalized and len(table.columns) <= 2:
                continue

            count += 1
        return count

    def _check_caption_order_and_numbering(self, caption_type: str) -> None:
        """Check JCLC caption placement and numbering in body XML order."""
        blocks = self._get_body_blocks()
        if caption_type == "table":
            caption_pattern = re.compile(r"^table\s*(\d+)\b", re.IGNORECASE)
            expected_side = "above"
            object_kind = "table"
            category = "tables"
            label = "Table"
        else:
            caption_pattern = re.compile(r"^(?:figure|fig\.?)\s*(\d+)\b", re.IGNORECASE)
            # APA7 (JCLC's reference style) places the figure number and
            # title above the image, exactly like a table caption - any
            # "Note." text goes below, but that's not the caption itself.
            expected_side = "above"
            object_kind = "drawing"
            category = "figures"
            label = "Figure"

        captions = []
        for index, block in enumerate(blocks):
            match = caption_pattern.match(block["text"])
            if not match:
                continue
            # A real caption line is just "Table 1" / "Figure 1" (maybe
            # with trailing punctuation) - its own short paragraph, not a
            # full sentence. Without this, an ordinary sentence that
            # happens to start with "Table 1 displays..." or "Figure 2
            # shows..." (an extremely common way to refer back to a table
            # or figure in body text) gets miscounted as a second caption
            # for the same number, which is exactly what was producing
            # phantom duplicates like "1, 1, 2, 2".
            trailing = block["text"][match.end():].strip()
            if len(trailing) > 3:
                continue
            captions.append((index, block, int(match.group(1))))

        # JCLC: a table / figure caption does not end with a full stop.
        full_re = re.compile(r"^(?:table|figure|fig\.?)\s*\d+\s*[.:]\s*\S", re.IGNORECASE)
        for index, block in enumerate(blocks):
            text = block["text"]
            m_label = caption_pattern.match(text)
            if not m_label or block["kind"] != "paragraph":
                continue
            cap_block = None
            if len(text[m_label.end():].strip()) <= 3:
                nxt = self._nearest_content_block(blocks, index, 1)       # title line under a label-only line
                if nxt and nxt["kind"] == "paragraph" and 0 < len(nxt["text"]) <= 250 \
                        and not caption_pattern.match(nxt["text"]):
                    cap_block = nxt
            elif full_re.match(text) and len(text) <= 250:
                cap_block = block
            if cap_block and re.search(r"(?<!\.)\.\s*$", cap_block["text"]):
                self._add_issue(
                    category=category,
                    location=f"Caption: {truncate_text(cap_block['text'], 30)}",
                    para_index=cap_block["paragraph_index"],
                    description=f"{label} caption must not end with a full stop",
                    current="Ends with a full stop",
                    expected="No full stop at the end of the caption",
                    severity="warning",
                    text_preview=truncate_text(cap_block["text"], 50),
                )

        numbers = [number for _, _, number in captions]
        expected_numbers = list(range(1, len(numbers) + 1))
        if numbers and numbers != expected_numbers:
            self._add_issue(
                category=category,
                location=f"{label} Captions",
                para_index=-1,
                description=f"{label} numbering is not continuous",
                current=", ".join(str(number) for number in numbers),
                expected=", ".join(str(number) for number in expected_numbers),
                severity="warning",
            )

        for index, block, _ in captions:
            previous_block = self._nearest_content_block(blocks, index, -1)
            next_block = self._nearest_content_block(blocks, index, 1)
            reversed_position = (
                expected_side == "above"
                and previous_block
                and previous_block["kind"] == object_kind
            ) or (
                expected_side == "below"
                and next_block
                and next_block["kind"] == object_kind
            )
            if reversed_position:
                self._add_issue(
                    category=category,
                    location=f"Caption: {truncate_text(block['text'], 30)}",
                    para_index=block["paragraph_index"],
                    description=f"{label} caption should appear {expected_side} the {caption_type}",
                    current=f"Caption appears on the wrong side of the {caption_type}",
                    expected=f"Caption immediately {expected_side} the {caption_type}",
                    severity="warning",
                    text_preview=truncate_text(block["text"], 50),
                )

    def _get_body_blocks(self) -> List[Dict[str, Any]]:
        """Return body paragraphs, tables, and drawings in document order."""
        blocks = []
        paragraph_index = -1
        for child in self.document.element.body.iterchildren():
            if child.tag == qn("w:p"):
                paragraph_index += 1
                text = "".join(
                    text_node.text or ""
                    for text_node in child.iter()
                    if text_node.tag == qn("w:t")
                )
                has_drawing = any(
                    node.tag == qn("w:drawing")
                    for node in child.iter()
                )
                blocks.append({
                    "kind": "drawing" if has_drawing else "paragraph",
                    "text": re.sub(r"\s+", " ", text).strip(),
                    "paragraph_index": paragraph_index,
                })
            elif child.tag == qn("w:tbl"):
                blocks.append({
                    "kind": "table",
                    "text": "",
                    "paragraph_index": -1,
                })
        return blocks

    @staticmethod
    def _nearest_content_block(
        blocks: List[Dict[str, Any]],
        start_index: int,
        direction: int,
    ) -> Optional[Dict[str, Any]]:
        """Find the nearest non-empty paragraph, table, or drawing."""
        index = start_index + direction
        while 0 <= index < len(blocks):
            block = blocks[index]
            if block["kind"] != "paragraph" or block["text"]:
                return block
            index += direction
        return None

    def _count_body_images_before_back_matter(self) -> int:
        """Count body images while ignoring author photos in biography sections."""
        count = 0
        seen_figure_caption = False
        for child in self.document.element.body.iterchildren():
            block_text = " ".join(
                text_node.text or ""
                for text_node in child.iter()
                if text_node.tag == qn("w:t")
            )
            normalized = re.sub(r"\s+", " ", block_text.lower()).strip()
            if normalized.startswith(("biographies of authors", "appendix")):
                break
            if re.match(r"^(figure|fig\.?)\s*\d+\b", normalized):
                seen_figure_caption = True
            if seen_figure_caption:
                count += sum(1 for node in child.iter() if node.tag == qn("w:drawing"))
        return count

    def _extract_reference_texts_from_body_xml(self) -> List[str]:
        """Extract references from body XML, including Word content controls."""
        chunks = []
        in_references = False

        for child in self.document.element.body.iterchildren():
            block_text = "\n".join(
                text_node.text or ""
                for text_node in child.iter()
                if text_node.tag == qn("w:t")
            )
            normalized = re.sub(r"\s+", " ", block_text.lower()).strip()
            if not normalized:
                continue
            if normalized.startswith(("biographies of authors", "appendix")):
                break
            if normalized.startswith(("references", "bibliography", "works cited")):
                in_references = True
                block_text = re.sub(r"^\s*(references|bibliography|works cited)\s*", "", block_text, flags=re.IGNORECASE)
            if in_references and block_text.strip():
                chunks.append(block_text)

        reference_text = "\n".join(chunks)
        markers = list(re.finditer(r"(?m)^\s*\[\d+\]\s*$", reference_text))
        if markers:
            references = []
            for index, marker in enumerate(markers):
                start = marker.end()
                end = markers[index + 1].start() if index + 1 < len(markers) else len(reference_text)
                entry = reference_text[start:end].strip()
                if entry:
                    references.append(entry)
            return references

        return [
            line.strip()
            for line in reference_text.splitlines()
            if "doi" in line.lower() or re.search(r"\b(19|20)\d{2}\b", line)
        ]

    def _check_ieee_reference_consistency(self) -> None:
        """Check reference numbering and links between citations and entries."""
        reference_numbers = []
        citation_numbers = set()
        in_references = False

        for child in self.document.element.body.iterchildren():
            paragraph_texts = [
                "".join(
                    text_node.text or ""
                    for text_node in paragraph.iter()
                    if text_node.tag == qn("w:t")
                )
                for paragraph in child.iter()
                if paragraph.tag == qn("w:p")
            ]
            for block_text in paragraph_texts:
                normalized = re.sub(r"\s+", " ", block_text).strip()
                lower_text = normalized.lower()
                if not normalized:
                    continue

                if lower_text.startswith(("references", "bibliography", "works cited")):
                    in_references = True
                    block_text = re.sub(
                        r"^\s*(references|bibliography|works cited)\s*",
                        "",
                        block_text,
                        flags=re.IGNORECASE,
                    )
                elif in_references and lower_text.startswith(("biographies of authors", "appendix")):
                    break

                if in_references:
                    match = re.match(r"^\s*\[(\d+)\]", block_text)
                    if match:
                        reference_numbers.append(int(match.group(1)))
                else:
                    citation_numbers.update(self._extract_citation_numbers(block_text))

        if not reference_numbers:
            return

        citation_numbers.discard(0)
        highest_plausible_citation = max(reference_numbers) + 1
        citation_numbers = {
            number
            for number in citation_numbers
            if number <= highest_plausible_citation
        }

        expected_numbers = list(range(1, len(reference_numbers) + 1))
        if reference_numbers != expected_numbers:
            self._add_issue(
                category="references",
                location="Reference Numbering",
                para_index=-1,
                description="Reference numbering is not continuous",
                current=", ".join(f"[{number}]" for number in reference_numbers),
                expected=", ".join(f"[{number}]" for number in expected_numbers),
                severity="warning",
            )

        reference_number_set = set(reference_numbers)
        missing_references = sorted(citation_numbers - reference_number_set)
        if missing_references:
            self._add_issue(
                category="references",
                location="In-text Citations",
                para_index=-1,
                description="In-text citation has no matching reference",
                current=", ".join(f"[{number}]" for number in missing_references),
                expected="Every in-text citation matches a numbered reference entry",
                severity="error",
            )

        uncited_references = sorted(reference_number_set - citation_numbers)
        if uncited_references:
            self._add_issue(
                category="references",
                location="Reference Entries",
                para_index=-1,
                description="Reference is not cited in the manuscript",
                current=", ".join(f"[{number}]" for number in uncited_references),
                expected="Every reference entry is cited in the manuscript",
                severity="warning",
            )

    @staticmethod
    def _extract_citation_numbers(text: str) -> set:
        """Extract IEEE citation numbers, including comma lists and ranges."""
        numbers = set()
        for match in re.finditer(r"\[(\d+)\]\s*[-\u2013\u2014]\s*\[(\d+)\]", text):
            start, end = int(match.group(1)), int(match.group(2))
            if start <= end:
                numbers.update(range(start, end + 1))

        for match in re.finditer(
            r"\[(\d+(?:\s*(?:,|[-\u2013\u2014])\s*\d+)*)\]",
            text,
        ):
            content = match.group(1)
            for part in re.split(r"\s*,\s*", content):
                range_match = re.fullmatch(r"(\d+)\s*[-\u2013\u2014]\s*(\d+)", part)
                if range_match:
                    start, end = int(range_match.group(1)), int(range_match.group(2))
                    if start <= end:
                        numbers.update(range(start, end + 1))
                else:
                    numbers.add(int(part))
        return numbers

    def _get_numbering_level(self, paragraph):
        """Return the numbering level XML element for a numbered paragraph."""
        p_pr = paragraph._p.pPr
        if p_pr is None or p_pr.numPr is None or p_pr.numPr.numId is None:
            return None

        num_id = p_pr.numPr.numId.val
        ilvl = str(p_pr.numPr.ilvl.val if p_pr.numPr.ilvl is not None else 0)
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

    def _get_numbering_bold(self, paragraph) -> Optional[bool]:
        """Read bold formatting from a paragraph numbering level."""
        level = self._get_numbering_level(paragraph)
        if level is None:
            return None
        r_pr = level.find(qn("w:rPr"))
        if r_pr is None:
            return None
        bold = r_pr.find(qn("w:b"))
        if bold is None:
            return None
        value = bold.get(qn("w:val"))
        return value not in {"0", "false", "False", "off"}
    def _check_line_spacing(self):
        """Check line spacing throughout document"""
        body_rules = self.rules.get("body", {})
        expected_spacing = body_rules.get("line_spacing", 1.5)
        
        spacing_issues = 0
        
        for cp in self.classifications:
            if cp.paragraph_type in [ParagraphType.BODY, ParagraphType.ABSTRACT_CONTENT]:
                para = self.document.paragraphs[cp.index]
                current_spacing = get_line_spacing(para)
                
                if current_spacing and abs(current_spacing - expected_spacing) > 0.1:
                    spacing_issues += 1
        
        if spacing_issues > 0:
            self._add_issue(
                category="line_spacing",
                location="Line Spacing",
                para_index=-1,
                description=f"{spacing_issues} paragraphs have incorrect line spacing",
                current="Varies",
                expected=f"{expected_spacing} line spacing",
                severity="warning"
            )
    
    def _gather_statistics(self) -> Dict[str, Any]:
        """Gather document statistics"""
        # Count paragraphs by type
        type_counts = {}
        for cp in self.classifications:
            type_name = cp.paragraph_type.value
            type_counts[type_name] = type_counts.get(type_name, 0) + 1
        
        # Count words
        total_words = sum(
            len(cp.text.split()) for cp in self.classifications
            if cp.paragraph_type not in [ParagraphType.EMPTY]
        )
        
        # Count tables and figures
        table_count = len(self.document.tables)
        figure_count = sum(
            1 for cp in self.classifications
            if cp.paragraph_type == ParagraphType.CAPTION and
            re.match(r'^(figure|fig\.?)\s*\d+', cp.text.lower())
        )
        
        # Count references
        reference_count = sum(
            1 for cp in self.classifications
            if cp.paragraph_type == ParagraphType.REFERENCE
        )
        
        return {
            "total_paragraphs": len(self.document.paragraphs),
            "classified_paragraphs": len(self.classifications),
            "paragraph_types": type_counts,
            "total_words": total_words,
            "table_count": table_count,
            "figure_count": figure_count,
            "reference_count": reference_count,
            "pages": len(self.document.sections)
        }
    
    def get_comparison_data(self) -> List[Dict[str, Any]]:
        """Get data for Turnitin-style comparison view"""
        comparison_data = []
        
        for category, issues in self.issues.items():
            for issue in issues:
                comparison_data.append({
                    "category": category,
                    "location": issue.location,
                    "paragraph_index": issue.paragraph_index,
                    "text_preview": issue.text_preview,
                    "current": issue.current_value,
                    "expected": issue.expected_value,
                    "description": issue.description,
                    "severity": issue.severity
                })
        
        return comparison_data