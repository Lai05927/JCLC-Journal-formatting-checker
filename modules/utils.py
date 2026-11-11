"""
Utility functions for the Academic Manuscript Format Checker
"""

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from collections import Counter
from copy import deepcopy
from io import BytesIO
import re
from typing import Dict, List, Tuple, Any, Optional


# Font equivalence groups - fonts that are considered equivalent
FONT_EQUIVALENTS = {
    "times": ["times new roman", "times", "times-roman", "timesnewroman"],
    "arial": ["arial", "arial unicode ms", "arial narrow"],
    "calibri": ["calibri", "calibri light"],
    "cambria": ["cambria", "cambria math"],
    "palatino": ["palatino linotype", "palatino", "book antiqua"],
}


def contains_east_asian_text(text: str) -> bool:
    """Return True when text contains CJK characters."""
    return bool(re.search(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", text or ""))


def is_font_equivalent(font1: str, font2: str) -> bool:
    """
    Check if two fonts are equivalent (e.g., 'Times New Roman' and 'Times')
    
    Args:
        font1: First font name
        font2: Second font name
        
    Returns:
        True if fonts are equivalent
    """
    if font1 is None or font2 is None:
        return False
    
    f1 = font1.lower().strip()
    f2 = font2.lower().strip()
    
    # Exact match
    if f1 == f2:
        return True
    
    # Check equivalence groups
    for group in FONT_EQUIVALENTS.values():
        if f1 in group and f2 in group:
            return True
    
    return False


def load_document(file_path_or_bytes) -> Document:
    """Load a Word document from file path or bytes"""
    try:
        if isinstance(file_path_or_bytes, str):
            return Document(file_path_or_bytes)
        else:
            return Document(file_path_or_bytes)
    except Exception as e:
        raise ValueError(f"Failed to load document: {str(e)}")


def get_paragraph_text(paragraph) -> str:
    """Get the full text of a paragraph"""
    return paragraph.text.strip()


def get_run_font_info(run) -> Dict[str, Any]:
    """Extract font information from a run - ENHANCED VERSION"""
    font = run.font
    
    font_name = None
    font_size = None
    bold = None
    italic = None
    
    # Method 1: Direct font properties
    if font.name:
        font_name = font.name
    
    if font.size is not None:
        font_size = font.size.pt
    
    bold = font.bold
    italic = font.italic
    
    # Method 2: Check XML directly for more accurate extraction
    rPr = run._element.rPr
    if rPr is not None:
        # Font name from XML
        if font_name is None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                latin_font = (
                    rFonts.get(qn('w:ascii')) or
                    rFonts.get(qn('w:hAnsi')) or
                    rFonts.get(qn('w:cs'))
                )
                east_asia_font = rFonts.get(qn('w:eastAsia'))
                font_name = latin_font
                if font_name is None and contains_east_asian_text(run.text):
                    font_name = east_asia_font
        
        # Font size from XML (w:sz is in half-points)
        if font_size is None:
            sz = rPr.find(qn('w:sz'))
            if sz is not None:
                sz_val = sz.get(qn('w:val'))
                if sz_val:
                    try:
                        font_size = int(sz_val) / 2  # Convert half-points to points
                    except:
                        pass
            
            # Also check szCs for complex script size
            if font_size is None:
                szCs = rPr.find(qn('w:szCs'))
                if szCs is not None:
                    sz_val = szCs.get(qn('w:val'))
                    if sz_val:
                        try:
                            font_size = int(sz_val) / 2
                        except:
                            pass
        
        # Bold from XML
        if bold is None:
            b = rPr.find(qn('w:b'))
            if b is not None:
                b_val = b.get(qn('w:val'))
                bold = b_val != '0' if b_val else True
        
        # Italic from XML
        if italic is None:
            i = rPr.find(qn('w:i'))
            if i is not None:
                i_val = i.get(qn('w:val'))
                italic = i_val != '0' if i_val else True
    
    return {
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "underline": font.underline,
        "strike": font.strike,
        "subscript": font.subscript,
        "superscript": font.superscript
    }


def get_hyperlink_run_font_names(paragraph) -> List[Optional[str]]:
    """Return the explicit font name (or None) for each hyperlinked run with text.

    paragraph.runs (and everything built on it, like get_paragraph_font_info)
    only sees direct <w:r> children of a paragraph. A run properly wrapped in
    <w:hyperlink> - the correct structure for a DOI/URL link - is invisible
    to that, so a hyperlink's font is otherwise never inspected at all.

    A hyperlink run commonly carries no direct font override - it relies on
    Word's built-in "Hyperlink" character style, which frequently isn't even
    serialized into the document's own styles.xml (as opposed to a genuine
    document theme font), making the *effective* font unreliable to resolve
    without parsing the theme. Returns None for such a run rather than
    guessing - the caller should treat "no explicit font" on a hyperlink as
    needing a fix in its own right, since that's exactly what setting one
    explicitly resolves, regardless of what a theme happens to resolve to.
    """
    from docx.text.run import Run

    results: List[Optional[str]] = []
    for hyperlink in paragraph._p.findall(qn("w:hyperlink")):
        for run_element in hyperlink.findall(qn("w:r")):
            run = Run(run_element, paragraph)
            if not run.text.strip():
                continue
            info = get_run_font_info(run)
            results.append(info.get("font_name"))
    return results


def has_run_level_size_mismatch(paragraph, expected_size: Optional[float], tolerance: float = 0.5) -> bool:
    """Return True if any individual non-whitespace run's font size
    deviates from expected_size beyond tolerance.

    get_paragraph_font_info reports the *most common* size across a
    paragraph's runs, which is the right summary for display purposes
    but is blind to a minority-mismatched run: a paragraph with two
    correctly-sized runs and one run still at the wrong size reports the
    correct size as "most common" and shows no issue at all, even
    though a fragment of it still needs fixing. Checked independently of
    that summary so a check can flag the paragraph as having an issue
    even when the majority vote alone would miss it entirely.
    """
    if expected_size is None:
        return False
    for run in paragraph.runs:
        if not run.text.strip():
            continue
        info = get_run_font_info(run)
        size = info.get("font_size")
        if size is not None and abs(size - expected_size) > tolerance:
            return True
    return False


def get_paragraph_font_info(paragraph) -> Dict[str, Any]:
    """Get the most common font info from a paragraph's runs - ENHANCED VERSION"""
    font_names = []
    font_sizes = []
    bold_values = []
    italic_values = []
    
    # Method 1: Extract from runs
    if paragraph.runs:
        for run in paragraph.runs:
            if run.text.strip():
                info = get_run_font_info(run)
                if info["font_name"]:
                    font_names.append(info["font_name"])
                if info["font_size"]:
                    font_sizes.append(info["font_size"])
                if info["bold"] is not None:
                    bold_values.append(info["bold"])
                if info["italic"] is not None:
                    italic_values.append(info["italic"])
    
    # Method 2: Check paragraph style if we didn't get values from runs
    if not font_names or not font_sizes:
        style = paragraph.style
        if style:
            # Check direct style font
            if style.font:
                if not font_names and style.font.name:
                    font_names.append(style.font.name)
                if not font_sizes and style.font.size:
                    font_sizes.append(style.font.size.pt)
                if not bold_values and style.font.bold is not None:
                    bold_values.append(style.font.bold)
            
            # Check base style
            base_style = style.base_style
            if base_style and base_style.font:
                if not font_names and base_style.font.name:
                    font_names.append(base_style.font.name)
                if not font_sizes and base_style.font.size:
                    font_sizes.append(base_style.font.size.pt)
    
    # Method 3: Check paragraph XML directly
    if not font_sizes:
        pPr = paragraph._element.pPr
        if pPr is not None:
            rPr = pPr.find(qn('w:rPr'))
            if rPr is not None:
                sz = rPr.find(qn('w:sz'))
                if sz is not None:
                    sz_val = sz.get(qn('w:val'))
                    if sz_val:
                        try:
                            font_sizes.append(int(sz_val) / 2)
                        except:
                            pass
    
    # Return most common values
    return {
        "font_name": Counter(font_names).most_common(1)[0][0] if font_names else None,
        "font_size": Counter(font_sizes).most_common(1)[0][0] if font_sizes else None,
        "bold": Counter(bold_values).most_common(1)[0][0] if bold_values else None,
        "italic": Counter(italic_values).most_common(1)[0][0] if italic_values else None
    }


def get_paragraph_alignment(paragraph) -> str:
    """Get paragraph alignment as string"""
    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
        WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
        WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY"
    }

    alignment = paragraph.alignment
    if alignment is None and paragraph.style is not None:
        alignment = paragraph.style.paragraph_format.alignment

    base_style = paragraph.style.base_style if paragraph.style is not None else None
    if alignment is None and base_style is not None:
        alignment = base_style.paragraph_format.alignment

    return alignment_map.get(alignment, "LEFT")


def get_sdt_reference_paragraphs(document) -> List[Paragraph]:
    """Return reference paragraphs stored inside Word content controls."""
    references = []
    in_references = False
    body = document.element.body

    for child in body.iterchildren():
        is_content_control = child.tag == qn("w:sdt")
        if child.tag == qn("w:p"):
            paragraphs = [Paragraph(child, document._body)]
        elif is_content_control:
            paragraphs = [
                Paragraph(node, document._body)
                for node in child.iter(qn("w:p"))
            ]
        else:
            paragraphs = []

        for paragraph in paragraphs:
            text = get_paragraph_text(paragraph).strip()
            normalized = re.sub(r"\s+", " ", text.lower())
            if normalized.startswith(("references", "bibliography", "works cited")):
                in_references = True
                continue
            if in_references and normalized.startswith(
                ("biographies of authors", "author biography", "appendix")
            ):
                return references
            if in_references and is_content_control and text:
                references.append(paragraph)

    return references


def get_line_spacing(paragraph) -> Optional[float]:
    """Get effective line spacing from direct, style, or document defaults."""
    line_spacing = paragraph.paragraph_format.line_spacing
    if line_spacing is None and paragraph.style is not None:
        line_spacing = paragraph.style.paragraph_format.line_spacing

    base_style = paragraph.style.base_style if paragraph.style is not None else None
    if line_spacing is None and base_style is not None:
        line_spacing = base_style.paragraph_format.line_spacing

    if line_spacing is None:
        document = getattr(paragraph.part, "document", None)
        styles = document.styles.element if document is not None else None
        doc_defaults = styles.find(qn("w:docDefaults")) if styles is not None else None
        spacing = (
            doc_defaults.find(
                f"{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:spacing')}"
            )
            if doc_defaults is not None
            else None
        )
        if spacing is not None:
            line_value = spacing.get(qn("w:line"))
            line_rule = spacing.get(qn("w:lineRule"), "auto")
            if line_value and line_rule == "auto":
                line_spacing = int(line_value) / 240

    return line_spacing


def set_space_after_pt(paragraph, points: float) -> None:
    """Set a paragraph's space-after in points, clearing any leftover
    afterLines attribute on the same <w:spacing> element.

    OOXML supports space-after in two independent units on the same
    element - w:after (points, twentieths-of-a-point) and w:afterLines
    (line multiples x100) - and Word uses w:afterLines whenever it's
    present and non-zero, ignoring w:after entirely. Setting
    paragraph_format.space_after only ever touches w:after, so a
    paragraph that had its spacing set in "lines" at some point (e.g. a
    user testing the checker by typing "8" into Word's "line spacing
    after" field) keeps rendering with that leftover line-based value
    even after this fix correctly writes the point-based one - Word
    still prefers the lines attribute that's still sitting there.
    """
    paragraph.paragraph_format.space_after = Pt(points)
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is not None and spacing.get(qn("w:afterLines")) is not None:
        del spacing.attrib[qn("w:afterLines")]


def set_space_before_pt(paragraph, points: float) -> None:
    """Set a paragraph's space-before in points, clearing any leftover
    beforeLines attribute - see set_space_after_pt for why this matters.
    """
    paragraph.paragraph_format.space_before = Pt(points)
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is not None and spacing.get(qn("w:beforeLines")) is not None:
        del spacing.attrib[qn("w:beforeLines")]


def has_lines_based_spacing(paragraph) -> bool:
    """Return True if the paragraph has a leftover afterLines/
    beforeLines attribute (spacing set via Word's "Lines" option rather
    than "Points"). Word renders using this value whenever it's present
    and non-zero, ignoring the points-based w:after/w:before entirely -
    so a paragraph can have a perfectly correct w:after value and still
    render with the wrong spacing if this is also set. Checked as its
    own, independent trigger: comparing only the points-based value
    against the target (as get_space_after_pt does) can come back
    "already matches" purely by coincidence with whatever w:after was
    left at before afterLines got added, which would skip calling
    set_space_after_pt/set_space_before_pt (the only place that clears
    this attribute) even though the paragraph still renders wrong.
    """
    pPr = paragraph._p.find(qn("w:pPr"))
    spacing = pPr.find(qn("w:spacing")) if pPr is not None else None
    if spacing is None:
        return False
    for attr in ("w:afterLines", "w:beforeLines"):
        value = spacing.get(qn(attr))
        if value is not None and value != "0":
            return True
    return False


def get_space_after_pt(paragraph) -> Optional[float]:
    """Get effective paragraph spacing after in points."""
    space_after = paragraph.paragraph_format.space_after
    if space_after is None and paragraph.style is not None:
        space_after = paragraph.style.paragraph_format.space_after

    base_style = paragraph.style.base_style if paragraph.style is not None else None
    if space_after is None and base_style is not None:
        space_after = base_style.paragraph_format.space_after

    if space_after is None:
        document = getattr(paragraph.part, "document", None)
        styles = document.styles.element if document is not None else None
        doc_defaults = styles.find(qn("w:docDefaults")) if styles is not None else None
        spacing = (
            doc_defaults.find(
                f"{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:spacing')}"
            )
            if doc_defaults is not None
            else None
        )
        if spacing is not None:
            after_value = spacing.get(qn("w:after"))
            if after_value is not None:
                try:
                    return int(after_value) / 20
                except ValueError:
                    return None

    return space_after.pt if space_after is not None else None


def get_space_before_pt(paragraph) -> Optional[float]:
    """Get effective paragraph spacing before in points."""
    space_before = paragraph.paragraph_format.space_before
    if space_before is None and paragraph.style is not None:
        space_before = paragraph.style.paragraph_format.space_before

    base_style = paragraph.style.base_style if paragraph.style is not None else None
    if space_before is None and base_style is not None:
        space_before = base_style.paragraph_format.space_before

    if space_before is None:
        document = getattr(paragraph.part, "document", None)
        styles = document.styles.element if document is not None else None
        doc_defaults = styles.find(qn("w:docDefaults")) if styles is not None else None
        spacing = (
            doc_defaults.find(
                f"{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:spacing')}"
            )
            if doc_defaults is not None
            else None
        )
        if spacing is not None:
            before_value = spacing.get(qn("w:before"))
            if before_value is not None:
                try:
                    return int(before_value) / 20
                except ValueError:
                    return None

    return space_before.pt if space_before is not None else None


def get_hanging_indent_inches(paragraph) -> Optional[float]:
    """Get effective hanging indent distance in inches."""
    first_line_indent = paragraph.paragraph_format.first_line_indent
    if first_line_indent is None and paragraph.style is not None:
        first_line_indent = paragraph.style.paragraph_format.first_line_indent

    base_style = paragraph.style.base_style if paragraph.style is not None else None
    if first_line_indent is None and base_style is not None:
        first_line_indent = base_style.paragraph_format.first_line_indent

    if first_line_indent is not None and first_line_indent.inches < 0:
        return abs(first_line_indent.inches)

    paragraph_properties = paragraph._p.pPr
    if paragraph_properties is not None:
        indentation = paragraph_properties.find(qn("w:ind"))
        if indentation is not None:
            hanging_value = indentation.get(qn("w:hanging"))
            if hanging_value is not None:
                try:
                    return int(hanging_value) / 1440
                except ValueError:
                    return None

    return None


def get_margins(document) -> Dict[str, float]:
    """Extract page margins from document in inches"""
    try:
        section = document.sections[0]
        return {
            "left": section.left_margin.inches if section.left_margin else 1.0,
            "right": section.right_margin.inches if section.right_margin else 1.0,
            "top": section.top_margin.inches if section.top_margin else 1.0,
            "bottom": section.bottom_margin.inches if section.bottom_margin else 1.0
        }
    except Exception:
        return {"left": 1.0, "right": 1.0, "top": 1.0, "bottom": 1.0}


def set_margins(document, margins: Dict[str, float]):
    """Set page margins for all sections"""
    for section in document.sections:
        section.left_margin = Inches(margins.get("left", 1.0))
        section.right_margin = Inches(margins.get("right", 1.0))
        section.top_margin = Inches(margins.get("top", 1.0))
        section.bottom_margin = Inches(margins.get("bottom", 1.0))


def format_change_description(before: Dict, after: Dict) -> str:
    """Create a description of formatting changes"""
    changes = []
    
    if before.get("font_name") != after.get("font_name"):
        changes.append(f"Font: {before.get('font_name', 'Unknown')} -> {after.get('font_name')}")
    
    if before.get("font_size") != after.get("font_size"):
        changes.append(f"Size: {before.get('font_size', 'Unknown')}pt -> {after.get('font_size')}pt")
    
    if before.get("bold") != after.get("bold"):
        changes.append(f"Bold: {before.get('bold')} -> {after.get('bold')}")
    
    if before.get("alignment") != after.get("alignment"):
        changes.append(f"Alignment: {before.get('alignment', 'Unknown')} -> {after.get('alignment')}")
    
    return "; ".join(changes) if changes else "No changes"


def truncate_text(text: str, max_length: int = 50) -> str:
    """Truncate text with ellipsis"""
    if len(text) <= max_length:
        return text
    return text[:max_length-3] + "..."


CAPTION_LOWERCASE_WORDS = {
    "a", "an", "the",
    "and", "as", "but", "for", "nor", "or", "so", "yet",
    "at", "by", "from", "in", "into", "of", "off", "on", "onto",
    "out", "over", "per", "to", "up", "via", "with",
}


def to_journal_caption_title_case(text: str) -> str:
    """Return JIWE-style caption title case while preserving the caption label."""
    match = re.match(r"^(\s*)((?:fig\.?|figure|table)\s*\d+\s*[\.:]\s*)(.*)$", text or "", re.IGNORECASE)
    if not match:
        return text

    leading, label, caption_body = match.groups()
    label = re.sub(r"^(fig\.?|figure|table)", lambda item: item.group(1).capitalize(), label, count=1, flags=re.IGNORECASE)
    words_seen = 0

    def convert_word(match_obj):
        nonlocal words_seen
        token = match_obj.group(0)
        words_seen += 1
        force_capital = words_seen == 1
        return _caption_title_case_token(token, force_capital)

    converted_body = re.sub(r"[A-Za-z][A-Za-z'’]*(?:[-–][A-Za-z][A-Za-z'’]*)*", convert_word, caption_body)
    return f"{leading}{label}{converted_body}"


def _caption_title_case_token(token: str, force_capital: bool = False) -> str:
    """Title-case one caption token, preserving acronyms and small joining words."""
    if re.fullmatch(r"[A-Z0-9]{2,}(?:[-–][A-Z0-9]{2,})*", token):
        return token

    parts = re.split(r"([-–])", token)
    converted = []
    for index, part in enumerate(parts):
        if part in {"-", "–"}:
            converted.append(part)
            continue
        lower = part.lower()
        should_lower = (
            not force_capital
            and index == 0
            and lower in CAPTION_LOWERCASE_WORDS
        )
        if should_lower:
            converted.append(lower)
        else:
            converted.append(lower[:1].upper() + lower[1:])
    return "".join(converted)


def is_empty_paragraph(paragraph) -> bool:
    """Check if paragraph is empty or contains only whitespace"""
    return not paragraph.text.strip()


def count_columns(document) -> int:
    """Detect number of columns in document"""
    try:
        section = document.sections[0]
        cols = section._sectPr.xpath('./w:cols/@w:num')
        if cols:
            return int(cols[0])
        return 1
    except Exception:
        return 1


def detect_reference_style(references: List[str]) -> str:
    """Detect reference citation style (IEEE, APA, etc.)"""
    ieee_pattern = r'^\[\d+\]'
    apa_pattern = r'\(\d{4}\)'
    
    ieee_count = sum(1 for ref in references if re.match(ieee_pattern, ref.strip()))
    apa_count = sum(1 for ref in references if re.search(apa_pattern, ref))
    
    if ieee_count > apa_count:
        return "IEEE"
    elif apa_count > ieee_count:
        return "APA"
    else:
        return "Unknown"


def clean_text_for_comparison(text: str) -> str:
    """Clean text for comparison purposes"""
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def classify_author_info_role(text: str, rules: Optional[Dict[str, Any]] = None) -> str:
    """Classify a front-matter author-information paragraph by its role."""
    normalized = re.sub(r"\s+", " ", text.lower()).strip()
    if (
        "corresponding author" in normalized
        or "orcid" in normalized
        or re.search(r"[\w.+-]+@[\w.-]+\.[a-z]{2,}", normalized)
    ):
        return "corresponding_author"
    if re.search(
        r"\b(university|universiti|faculty|department|school|college|institute|"
        r"centre|center|laboratory|address|jalan|street|campus|"
        r"sdn\.?\s*bhd\.?|bhd\.?|ltd\.?|limited|inc\.?|corp\.?|corporation|"
        r"company|systems|technologies)\b",
        normalized,
    ):
        return "affiliation"
    role_terms = rules.get("author_role_terms", {}) if isinstance(rules, dict) else {}
    affiliation_terms = role_terms.get("affiliation", []) if isinstance(role_terms, dict) else []
    for term in affiliation_terms:
        normalized_term = re.sub(r"\s+", " ", str(term or "").lower()).strip()
        if normalized_term and normalized_term in normalized:
            return "affiliation"
    return "author"


def calculate_compliance_score(issues: Dict[str, List]) -> float:
    """Calculate overall compliance score based on issues found"""
    weights = {
        "margins": 10,
        "layout": 10,
        "journal_header": 5,
        "title": 15,
        "author_info": 10,
        "body_text": 20,
        "headings": 10,
        "structure": 15,
        "tables": 5,
        "figures": 5,
        "references": 10,
        "line_spacing": 5,
        "other": 5
    }
    
    total_weight = sum(weights.values())
    deductions = 0
    
    for category, issue_list in issues.items():
        category_weight = weights.get(category, 5)
        if issue_list:
            # More issues = more deduction (up to full category weight)
            deduction = min(len(issue_list) * 2, category_weight)
            deductions += deduction
    
    score = max(0, 100 - (deductions / total_weight * 100))
    return round(score, 1)


def _remap_style_references(element, source_doc, target_doc) -> None:
    """Repoint every paragraph/character style reference inside a
    copied XML fragment from the source document's style ids to the
    target document's own ids for the *same-named* style.

    Style ids (w:pStyle / w:rStyle values like "16") are only meaningful
    within the document that defines them - the same id routinely names
    a completely different style in another file (observed: id "16" is
    "footer" in the template but "caption" - blue, accent1 - in the
    manuscript). Copying header/footer XML across documents as-is
    therefore silently applies whatever style happens to own that id in
    the target, e.g. turning the first-page page number blue. A
    reference whose style name doesn't exist in the target at all is
    dropped, falling back to the target's default style, rather than
    left pointing at an unrelated one.
    """
    def _name_map(doc, key_by_id):
        result = {}
        for style in doc.styles.element.findall(qn("w:style")):
            style_id = style.get(qn("w:styleId"))
            name_el = style.find(qn("w:name"))
            if style_id is None or name_el is None:
                continue
            name = (name_el.get(qn("w:val")) or "").lower()
            if key_by_id:
                result[style_id] = name
            else:
                result.setdefault(name, style_id)
        return result

    source_id_to_name = _name_map(source_doc, key_by_id=True)
    target_name_to_id = _name_map(target_doc, key_by_id=False)

    for tag in ("w:pStyle", "w:rStyle"):
        for ref in list(element.iter(qn(tag))):
            name = source_id_to_name.get(ref.get(qn("w:val")))
            target_id = target_name_to_id.get(name) if name else None
            if target_id is not None:
                ref.set(qn("w:val"), target_id)
            else:
                parent = ref.getparent()
                if parent is not None:
                    parent.remove(ref)


def _find_volume_issue_table(document):
    """Find the small table holding just the "Vol X, Issue Y, Month
    Year" / "eISSN: ..." running-header text sitting loose in the body,
    if present. Returns (table, vol_text) or (None, None).
    """
    for table in document.tables:
        if not table.rows:
            continue
        row_text = " ".join(cell.text for cell in table.rows[0].cells)
        if "issue" in row_text.lower() and "eissn" in row_text.lower():
            vol_text = table.rows[0].cells[0].text.strip()
            return table, vol_text
    return None, None


def _find_masthead_table(document):
    """Find the body table holding the journal name / DOI / copyright /
    "Published by MMU Press" masthead block, if present."""
    for table in document.tables:
        table_text = " ".join(cell.text for row in table.rows for cell in row.cells).lower()
        if "doi:" in table_text and ("mmu" in table_text or "published by" in table_text):
            return table
    return None


def _element_has_page_field(element) -> bool:
    """Return True if the XML element contains a PAGE field (either a
    complex field's instruction text or a simple field), i.e. it is or
    holds a page number."""
    for instr in element.iter(qn("w:instrText")):
        text = (instr.text or "").upper()
        if "PAGE" in text and "NUMPAGES" not in text:
            return True
    for simple in element.iter(qn("w:fldSimple")):
        instr = (simple.get(qn("w:instr")) or "").upper()
        if "PAGE" in instr and "NUMPAGES" not in instr:
            return True
    return False


def apply_template_first_page_footer(document: Document, template_path_or_bytes, warnings: Optional[List[str]] = None) -> None:
    """Replicate the template's own first-page footer exactly: the MMU
    logo and the DOI/copyright/license text sit in a single drawing
    anchored to the footer paragraph with a negative vertical offset,
    which is why it visually overlaps the page rather than sitting
    neatly inside the footer band the way an ordinary footer paragraph
    would - not something a plain "put this text in the footer"
    approach reproduces, so the whole drawing structure (and its exact
    anchor offsets) is copied from the template rather than
    reconstructed.

    The manuscript's own DOI (found in whatever masthead table it
    currently sits in, e.g. loose in the body) is substituted for the
    template's own placeholder DOI, and that source table is removed
    from the body once its content has been carried into the footer.

    Page numbering follows the manuscript, not the template: the
    template's own page-number field is not copied. If the manuscript's
    page 1 already had a page number it is kept (font set to Times New
    Roman 10pt, position/colour/spacing untouched); if it had none, none
    is added.
    """
    if warnings is None:
        warnings = []

    template_doc = template_path_or_bytes if hasattr(template_path_or_bytes, "sections") else Document(template_path_or_bytes)
    template_footer = template_doc.sections[0].first_page_footer
    template_part = template_footer.part
    image_rel = next(
        (rel for rel in template_part.rels.values() if "image" in rel.reltype and not rel.is_external),
        None,
    )
    if not template_footer.paragraphs or image_rel is None:
        return

    masthead_table = _find_masthead_table(document)
    doi_text = None
    if masthead_table is not None:
        for row in masthead_table.rows:
            for cell in row.cells:
                if "doi:" in cell.text.lower():
                    doi_text = cell.text.strip()
                    break

    section = document.sections[0]
    # Page 1 currently shows the first-page footer if the section already
    # had "different first page" on, otherwise the default footer. Any
    # page number it already carries is kept (copied before that footer's
    # content is replaced below); if it carries none, none is added -
    # the manuscript's own page numbering decides, not the template's.
    original_page_one_footer = (
        section.first_page_footer
        if section.different_first_page_header_footer
        else section.footer
    )
    preserved_page_number_children = [
        deepcopy(child)
        for child in original_page_one_footer._element
        if _element_has_page_field(child)
    ]
    section.different_first_page_header_footer = True
    manuscript_footer = section.first_page_footer
    ftr_element = manuscript_footer._element

    # Add the template's own logo image into the manuscript package and
    # get back a relationship id valid in *this* document, then copy the
    # footer paragraphs and repoint the drawing's image reference (the
    # r:embed attribute on its blip element) from the template's id to
    # this new one - copying the id as-is would dangle, since it isn't
    # registered in the manuscript's own relationships.
    manuscript_part = manuscript_footer.part
    new_rId, _ = manuscript_part.get_or_add_image(BytesIO(image_rel.target_part.blob))

    for child in list(ftr_element):
        ftr_element.remove(child)
    # Copy every top-level child, including the page-number field
    # wrapped in a <w:sdt> content control alongside the drawing/logo/
    # DOI paragraph - position and colour are left as whatever the
    # template/manuscript already has, only the font is set explicitly
    # below.
    for child in template_footer._element:
        # The template's own page-number field is never copied - whether
        # page 1 gets a page number is decided by what the manuscript
        # already had (see preserved_page_number_children above), not by
        # the template.
        if _element_has_page_field(child):
            continue
        new_child = deepcopy(child)
        for blip in new_child.iter(qn("a:blip")):
            if blip.get(qn("r:embed")) == image_rel.rId:
                blip.set(qn("r:embed"), new_rId)
        _remap_style_references(new_child, template_doc, document)
        ftr_element.append(new_child)

    # Put the manuscript's own page number (if it had one) back, ahead of
    # the drawing paragraph - the same order the template itself uses.
    for offset, child in enumerate(preserved_page_number_children):
        ftr_element.insert(offset, child)

    # Page-number font set explicitly to Times New Roman 10pt - nothing
    # else about it (position, colour, spacing) is touched.
    for preserved in preserved_page_number_children:
        for r_el in preserved.iter(qn("w:r")):
            r_pr = r_el.find(qn("w:rPr"))
            if r_pr is None:
                r_pr = OxmlElement("w:rPr")
                r_el.insert(0, r_pr)
            rfonts = r_pr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                r_pr.append(rfonts)
            rfonts.set(qn("w:ascii"), "Times New Roman")
            rfonts.set(qn("w:hAnsi"), "Times New Roman")
            sz = r_pr.find(qn("w:sz"))
            if sz is None:
                sz = OxmlElement("w:sz")
                r_pr.append(sz)
            sz.set(qn("w:val"), "20")

    if doi_text:
        # The DOI/copyright text lives inside the drawing's own text box
        # (once in the live drawing content, once more in its VML
        # fallback for older Word versions) rather than as an ordinary
        # footer paragraph run, so it has to be found and replaced by
        # walking every <w:t> text node in the footer's XML directly -
        # paragraph.runs never sees text nested this deep inside a
        # drawing.
        for t_el in ftr_element.iter(qn("w:t")):
            if t_el.text and "doi:" in t_el.text.lower():
                t_el.text = doi_text

    if masthead_table is not None:
        masthead_table._element.getparent().remove(masthead_table._element)

    warnings.append(
        "Replicated the template's first-page footer (MMU logo, DOI, "
        "copyright/license text, in its own overlapping position) and "
        "removed the corresponding duplicate content from the body."
    )


def _find_journal_title_paragraph_index(document) -> Optional[int]:
    """Return the index of the body paragraph holding the big
    "JOURNAL OF ..." masthead title, if present."""
    for i, p in enumerate(document.paragraphs):
        if p.text.strip().lower().startswith("journal of"):
            return i
    return None


def apply_template_first_page_header(document: Document, template_path_or_bytes, warnings: Optional[List[str]] = None) -> None:
    """Replicate the template's own first-page header exactly: the big
    journal title, the "Vol X, Issue Y" / "eISSN" line, and the two
    horizontal rule lines around the title - all as one first-page
    header, copied structurally from the template rather than
    reconstructed property-by-property, since the rule lines are drawn
    shapes anchored to the header paragraph, not something a formatting
    rule can express.

    The manuscript's own volume/issue line and masthead title, if they
    exist loose in the body (as ordinary pdf2docx-era leftovers or an
    author's own copy-paste), are removed from the body once their
    content has been carried into the header, so nothing is duplicated.
    """
    if warnings is None:
        warnings = []

    template_doc = template_path_or_bytes if hasattr(template_path_or_bytes, "sections") else Document(template_path_or_bytes)
    template_header = template_doc.sections[0].first_page_header
    if not template_header.paragraphs or not any(p.text.strip() for p in template_header.paragraphs):
        return

    vol_table, vol_text = _find_volume_issue_table(document)

    section = document.sections[0]
    section.different_first_page_header_footer = True
    manuscript_header = section.first_page_header
    hdr_element = manuscript_header._element

    for child in list(hdr_element):
        hdr_element.remove(child)
    # Copy every top-level child, not just <w:p> - the template's title
    # paragraph carries a bookmarkStart/bookmarkEnd pair as siblings of
    # the <w:p> elements themselves, and a w:p-only copy would leave a
    # dangling bookmarkStart with no matching end.
    for child in template_header._element:
        new_child = deepcopy(child)
        _remap_style_references(new_child, template_doc, document)
        hdr_element.append(new_child)

    if vol_text:
        for p in manuscript_header.paragraphs:
            # The template's own "Vol X, Issue Y, Month Year" text is
            # split across multiple runs (observed: "...July 202" then
            # a separate "6" run for the trailing digit) - replacing
            # only the one run matching "vol"/"issue" leaves that
            # trailing fragment appended after the substituted text
            # ("...2026" + "6" -> "20266"). Every run up to the first
            # tab (the tab separates this text from "eISSN: ...") is
            # this same volume/issue text, so the whole group is
            # collapsed into the first run and the rest cleared - except
            # the group can also include a leading, otherwise-empty run
            # that actually carries the rule-line drawing (observed at
            # index 0, before the text starts). Run.text's setter clears
            # a run's whole content to replace it with plain text, so
            # targeting that run for the replacement would silently
            # destroy the drawing along with it. Only text-bearing runs
            # within the group are eligible as the replacement target.
            runs = p.runs
            group_end = 0
            while group_end < len(runs) and "\t" not in runs[group_end].text:
                group_end += 1
            group = [r for r in runs[:group_end] if len(list(r._r.iter(qn("w:drawing")))) == 0]
            if any("vol" in r.text.lower() for r in group) and any("issue" in r.text.lower() for r in group):
                for idx, run in enumerate(group):
                    run.text = vol_text if idx == 0 else ""

    if vol_table is not None:
        vol_table._element.getparent().remove(vol_table._element)

    # The template highlights its own "Vol X, Issue Y" placeholder text
    # yellow to draw the author's eye to it while they fill it in - not
    # something that should carry over into a finished manuscript.
    for p in manuscript_header.paragraphs:
        for run in p.runs:
            run.font.highlight_color = None

    # Font/spacing set explicitly here rather than left as whatever the
    # copied paragraph happened to inherit - the template's own runs
    # don't all carry an explicit font, and inheriting the manuscript's
    # own default style (which can differ from the template's) would
    # make the Vol/eISSN line and the masthead title's spacing drift
    # from one manuscript to the next.
    header_paragraphs = manuscript_header.paragraphs
    if len(header_paragraphs) > 0:
        for run in header_paragraphs[0].runs:
            if run.text.strip():
                run.font.name = "Times New Roman"
                run.font.size = Pt(10)
    if len(header_paragraphs) > 1:
        title_pf = header_paragraphs[1].paragraph_format
        title_pf.space_before = Pt(0)
        title_pf.space_after = Pt(0)
        title_pf.line_spacing = 1.0

    title_idx = _find_journal_title_paragraph_index(document)
    if title_idx is not None:
        title_p = document.paragraphs[title_idx]
        title_p._element.getparent().remove(title_p._element)

    # The default (page 2+) running header repeats the same "Vol X,
    # Issue Y" text in its own convention ("Vol. 6, Issue 1, 2026") -
    # kept in sync with whatever the manuscript's own volume/issue text
    # says, rather than left at the template's own placeholder issue
    # number.
    if vol_text:
        match = re.search(r"issue\s*(\d+)", vol_text, re.IGNORECASE)
        vol_match = re.search(r"vol\.?\s*(\d+)", vol_text, re.IGNORECASE)
        year_match = re.search(r"(20\d{2})", vol_text)
        if match and vol_match and year_match:
            for p in section.header.paragraphs:
                for run in p.runs:
                    if "vol" in run.text.lower() and "issue" in run.text.lower():
                        run.text = f"Vol. {vol_match.group(1)}, Issue {match.group(1)}, {year_match.group(1)} "

    warnings.append(
        "Replicated the template's first-page header (journal title, "
        "volume/eISSN line, rule lines) and removed the corresponding "
        "duplicate content from the body."
    )


def get_document_sections(document) -> Dict[str, Tuple[int, int]]:
    """Identify sections in the document and their paragraph ranges"""
    sections = {}
    current_section = "preamble"
    section_start = 0
    
    section_keywords = {
        "abstract": ["abstract"],
        "keywords": ["keywords", "key words"],
        "introduction": ["introduction", "1. introduction", "1 introduction"],
        "literature_review": ["literature review", "related work", "2. literature"],
        "methodology": ["methodology", "method", "materials and methods", "3. methodology"],
        "results": ["results", "findings", "4. results"],
        "discussion": ["discussion", "5. discussion"],
        "conclusion": ["conclusion", "conclusions", "6. conclusion"],
        "references": ["references", "bibliography", "works cited"]
    }
    
    for i, para in enumerate(document.paragraphs):
        text = para.text.strip().lower()
        
        for section_name, keywords in section_keywords.items():
            if any(text == kw or text.startswith(kw + ":") for kw in keywords):
                if current_section != "preamble":
                    sections[current_section] = (section_start, i - 1)
                current_section = section_name
                section_start = i
                break
    
    # Add the last section
    if current_section:
        sections[current_section] = (section_start, len(document.paragraphs) - 1)
    
    return sections


def rgba_to_hex(r: int, g: int, b: int, a: float = 1.0) -> str:
    """Convert RGBA to hex color"""
    return f"#{r:02x}{g:02x}{b:02x}"


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color to RGB tuple"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))