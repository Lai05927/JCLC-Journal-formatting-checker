"""Deterministic review guidance payloads and fallback explanations."""

from collections import OrderedDict
import hashlib
import json
import re
from typing import Any, Dict, Iterable, List, Optional


EMAIL_PATTERN = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
ORCID_PATTERN = re.compile(r"\b\d{4}-\d{4}-\d{4}-\d{3}[\dXx]\b")


class ReviewGuidanceBuilder:
    """Build privacy-limited guidance without using an LLM."""

    MAX_GROUPS = 20
    MAX_EXAMPLES = 2

    _MANUAL_PATTERNS = (
        "not found",
        "missing",
        "numbering is not continuous",
        "citation has no matching",
        "not cited",
        "publication source may need italic",
        "should appear above",
        "should appear below",
        "may be missing a number",
        "required section",
        "required declaration",
        "word count",
        "invalid email",
        "invalid orcid",
        "multiple abstract paragraphs",
        "contains citation",
        "contains equation",
        "contains table or figure",
    )

    def build_pre_fix_payload(
        self,
        result: Any,
        rules: Dict[str, Any],
    ) -> Dict[str, Any]:
        """Build a compact payload from deterministic checker results."""
        profile = rules.get("_profile", {}) if isinstance(rules, dict) else {}
        groups = self._group_issues(
            getattr(result, "issues_by_category", {}) or {}
        )
        return {
            "profile": profile.get("name", "Generic"),
            "total_issues": int(getattr(result, "total_issues", 0) or 0),
            "compliance_index": float(
                getattr(result, "compliance_score", 0.0) or 0.0
            ),
            "groups": groups[:self.MAX_GROUPS],
        }

    def build_post_fix_payload(
        self,
        before_result: Any,
        after_result: Any,
        changes: Iterable[Any],
        validation: Any,
        rules: Dict[str, Any],
    ) -> Dict[str, Any]:
        """Build a compact payload from deterministic post-fix results."""
        profile = rules.get("_profile", {}) if isinstance(rules, dict) else {}
        change_counts: Dict[tuple, int] = {}
        for change in changes or []:
            key = (
                str(getattr(change, "change_type", "formatting") or "formatting"),
                str(getattr(change, "property_name", "formatting") or "formatting"),
            )
            change_counts[key] = change_counts.get(key, 0) + 1

        change_groups = [
            {"type": key[0], "property": key[1], "count": count}
            for key, count in sorted(
                change_counts.items(),
                key=lambda item: (-item[1], item[0][0], item[0][1]),
            )
        ]
        remaining_groups = self._group_issues(
            getattr(after_result, "issues_by_category", {}) or {}
        )
        increased = getattr(
            validation,
            "new_or_increased_categories",
            {},
        ) or {}
        return {
            "profile": profile.get("name", "Generic"),
            "issues_before": int(
                getattr(
                    validation,
                    "before_issues",
                    getattr(before_result, "total_issues", 0),
                ) or 0
            ),
            "issues_after": int(
                getattr(
                    validation,
                    "after_issues",
                    getattr(after_result, "total_issues", 0),
                ) or 0
            ),
            "safe": bool(getattr(validation, "is_safe", False)),
            "change_groups": change_groups,
            "remaining_groups": remaining_groups[:self.MAX_GROUPS],
            "increased_categories": {
                str(category): int(count)
                for category, count in sorted(increased.items())
            },
        }

    def _group_issues(
        self,
        issues_by_category: Dict[str, Iterable[Any]],
    ) -> List[Dict[str, Any]]:
        grouped: OrderedDict = OrderedDict()
        for category, issues in issues_by_category.items():
            for issue in issues:
                issue_category = (
                    getattr(issue, "category", "") or category or "other"
                )
                description = (
                    getattr(issue, "description", "") or "Formatting issue"
                )
                key = (
                    self._normalize(issue_category),
                    self._normalize(description),
                )
                if key not in grouped:
                    property_name = self._infer_safe_property(
                        issue_category,
                        description,
                        getattr(issue, "location", ""),
                    )
                    grouped[key] = {
                        "category": issue_category,
                        "description": description,
                        "severity": (
                            getattr(issue, "severity", "warning") or "warning"
                        ),
                        "count": 0,
                        "examples": [],
                        "auto_fix_supported": property_name is not None,
                        "property_name": property_name,
                        "review_reason": (
                            ""
                            if property_name
                            else self._manual_review_reason(
                                issue_category,
                                description,
                            )
                        ),
                        "priority": self._priority(
                            issue_category,
                            description,
                            getattr(issue, "severity", "warning"),
                        ),
                    }

                group = grouped[key]
                group["count"] += 1
                if len(group["examples"]) < self.MAX_EXAMPLES:
                    group["examples"].append(self._build_example(issue))

        return sorted(
            grouped.values(),
            key=lambda group: (
                group["priority"],
                group["category"].lower(),
                group["description"].lower(),
            ),
        )

    def build_issue_evidence(
        self,
        issue: Any,
        category: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Return user-facing evidence for one deterministic checker issue."""
        issue_category = (
            getattr(issue, "category", "")
            or category
            or "other"
        )
        description = (
            getattr(issue, "description", "")
            or "Formatting issue"
        )
        location = self.redact_text(
            getattr(issue, "location", "Document"),
            80,
        ) or "Document"
        raw_paragraph_index = getattr(issue, "paragraph_index", -1)
        try:
            paragraph_index = (
                int(raw_paragraph_index)
                if raw_paragraph_index is not None
                else -1
            )
        except (TypeError, ValueError):
            paragraph_index = -1
        paragraph = paragraph_index + 1 if paragraph_index >= 0 else "Document"
        property_name = self._infer_safe_property(
            issue_category,
            description,
            location,
        )
        auto_fix_supported = property_name is not None
        review_reason = (
            ""
            if auto_fix_supported
            else self._manual_review_reason(issue_category, description)
        )
        current_value = self.redact_text(
            getattr(issue, "current_value", ""),
            120,
        ) or "Not available"
        expected_value = self.redact_text(
            getattr(issue, "expected_value", ""),
            120,
        ) or "Not available"
        text_preview = self.redact_text(
            getattr(issue, "text_preview", ""),
            160,
        )

        if auto_fix_supported:
            action_label = "Auto-fix supported"
            action_detail = (
                f"Can safely adjust {self._friendly_property_name(property_name)} "
                "because this issue maps to a deterministic formatting property."
            )
        else:
            action_label = "Manual review required"
            action_detail = review_reason

        return {
            "category": issue_category,
            "category_label": issue_category.replace("_", " ").title(),
            "location": location,
            "paragraph": paragraph,
            "description": description,
            "current_value": current_value,
            "expected_value": expected_value,
            "severity": getattr(issue, "severity", "warning") or "warning",
            "text_preview": text_preview,
            "auto_fix_supported": auto_fix_supported,
            "property_name": property_name,
            "action_label": action_label,
            "action_detail": action_detail,
            "review_reason": review_reason,
        }

    def build_auto_fix_preview(
        self,
        issues_by_category: Dict[str, Iterable[Any]],
    ) -> Dict[str, Any]:
        """Summarize which detected issues are auto-fixable before fixing."""
        groups = self._group_issues(issues_by_category or {})
        supported_groups = [
            group for group in groups if group.get("auto_fix_supported")
        ]
        manual_groups = [
            group for group in groups if not group.get("auto_fix_supported")
        ]
        supported_count = sum(int(group.get("count", 0) or 0) for group in supported_groups)
        manual_count = sum(int(group.get("count", 0) or 0) for group in manual_groups)
        total_issues = supported_count + manual_count

        if total_issues == 0:
            summary = "No detected issues need auto-fix."
        elif supported_count and manual_count:
            summary = (
                f"{supported_count} detected issue"
                f"{'s' if supported_count != 1 else ''} can be auto-fixed; "
                f"{manual_count} issue"
                f"{'s' if manual_count != 1 else ''} need manual review."
            )
        elif supported_count:
            summary = (
                f"All {supported_count} detected issue"
                f"{'s' if supported_count != 1 else ''} are auto-fix candidates."
            )
        else:
            summary = (
                f"All {manual_count} detected issue"
                f"{'s' if manual_count != 1 else ''} need manual review."
            )

        return {
            "total_issues": total_issues,
            "supported_count": supported_count,
            "manual_count": manual_count,
            "can_run_auto_fix": supported_count > 0,
            "summary": summary,
            "supported_groups": supported_groups[:8],
            "manual_groups": manual_groups[:8],
        }

    def build_pre_fix_fallback(self, payload: Dict[str, Any]) -> str:
        """Return useful pre-fix guidance without an API."""
        groups = payload.get("groups", [])
        priority_groups = groups[:5]
        quick_fixes = [
            group for group in groups if group.get("auto_fix_supported")
        ]
        manual_groups = [
            group for group in groups if not group.get("auto_fix_supported")
        ]

        lines = ["Most important issues:"]
        lines.extend(self._format_groups(priority_groups))
        lines.append("")
        lines.append("Safe auto-fix candidates:")
        lines.extend(self._format_groups(quick_fixes))
        lines.append("")
        lines.append("Needs manual checking:")
        lines.extend(self._format_groups(manual_groups, include_reason=True))
        lines.append("")
        lines.append("Recommended review order:")
        if groups:
            lines.append(
                "1. Resolve errors and missing structural requirements. "
                "2. Apply supported formatting fixes. "
                "3. Review citations, equations, figures, tables, and remaining warnings."
            )
        else:
            lines.append("No detected issues require correction.")
        lines.append("")
        lines.append("What this guidance cannot decide:")
        lines.append(
            "This guidance organizes deterministic checker results only. "
            "It does not assess research quality, language quality, or manuscript acceptance."
        )
        return "\n".join(lines)

    def build_post_fix_fallback(self, payload: Dict[str, Any]) -> str:
        """Return useful post-fix guidance without an API."""
        change_groups = payload.get("change_groups", [])
        remaining_groups = payload.get("remaining_groups", [])

        lines = ["Auto-fixed items:"]
        if change_groups:
            for group in change_groups:
                lines.append(f"- {self._friendly_change_summary(group)}")
        else:
            lines.append("- No automatic changes were recorded.")

        lines.append("")
        lines.append("Issues still needing review:")
        lines.extend(self._format_groups(remaining_groups))
        lines.append("")
        lines.append("Why these were not auto-fixed:")
        lines.extend(
            self._format_groups(remaining_groups, include_reason=True)
        )
        lines.append("")
        lines.append("What to check next:")
        if remaining_groups:
            lines.append(
                "- Review unresolved errors first, then manually inspect "
                "citations, equations, figures, tables, and missing content."
            )
        else:
            lines.append(
                "- Perform a final visual review before submission."
            )
        lines.append("")
        lines.append("Auto-fix safety check:")
        if payload.get("safe"):
            lines.append(
                "- The auto-fix did not create additional detected issues."
            )
        else:
            lines.append(
                "- The auto-fix may have introduced a new issue. "
                "Review the corrected document before use."
            )
        return "\n".join(lines)

    @staticmethod
    def cache_key(payload: Dict[str, Any], guidance_type: str) -> str:
        """Return a stable cache key for a structured payload."""
        canonical = json.dumps(
            {"type": guidance_type, "payload": payload},
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=True,
        )
        return hashlib.sha256(canonical.encode("utf-8")).hexdigest()

    @staticmethod
    def redact_text(value: Any, limit: int = 80) -> str:
        """Redact personal identifiers and truncate evidence."""
        text = EMAIL_PATTERN.sub("[redacted email]", str(value or ""))
        text = ORCID_PATTERN.sub("[redacted ORCID]", text)
        text = re.sub(r"\s+", " ", text).strip()
        if len(text) > limit:
            return text[:limit - 3].rstrip() + "..."
        return text

    def _build_example(self, issue: Any) -> str:
        location = self.redact_text(
            getattr(issue, "location", "Document"),
            55,
        )
        current = self.redact_text(
            getattr(issue, "current_value", ""),
            80,
        )
        expected = self.redact_text(
            getattr(issue, "expected_value", ""),
            55,
        )
        value_text = f"{current} -> {expected}" if current or expected else ""
        return self.redact_text(
            f"{location}: {value_text}".strip(": "),
            150,
        )

    def _infer_safe_property(
        self,
        category: str,
        description: str,
        location: str,
    ) -> Optional[str]:
        lower = description.lower()
        lower_location = str(location or "").lower()
        if any(pattern in lower for pattern in self._MANUAL_PATTERNS):
            return None

        if "number" in lower or "heading number" in lower_location:
            if "bold" in lower:
                return "number_bold"
            if "font size" in lower or "size" in lower:
                return "number_font_size"
            if "font" in lower:
                return "number_font_name"
            return None

        if "line spacing" in lower:
            return "line_spacing"
        if "spacing after" in lower or "after spacing" in lower:
            return "space_after"
        if "hanging indent" in lower:
            return "hanging_indent"
        if "manual tab" in lower:
            return "manual_tabs"
        if "capitalization" in lower:
            return "capitalization"
        if "alignment" in lower:
            return "alignment"
        if "font size" in lower or " size " in f" {lower} ":
            return "font_size"
        if "font" in lower:
            return "font_name"
        if "bold" in lower:
            return "bold"
        if "italic" in lower:
            return "italic"
        if category == "margins" and "margin" in lower:
            return "margins"
        if category == "layout":
            if "page size" in lower:
                return "page_size"
            if "orientation" in lower:
                return "orientation"
            if "column" in lower:
                return "columns"
        return None

    @staticmethod
    def _manual_review_reason(category: str, description: str) -> str:
        lower = description.lower()
        if "publication source may need italic" in lower:
            return "The exact publication source segment must be selected before applying italic formatting."
        if "citation" in lower or "reference numbering" in lower:
            return "Citation changes may alter academic meaning."
        if "equation" in lower or "numbering" in lower:
            return "Renumbering may break cross-references."
        if "should appear" in lower or category in {"figures", "tables"}:
            return "Moving document objects can damage Word layout."
        if "not found" in lower or "missing" in lower or category == "structure":
            return "Missing content must be supplied and verified by an author."
        if category == "author_info":
            return "Author identity and contact details require human verification."
        return "No deterministic safe auto-fix is available for this issue."

    @staticmethod
    def _priority(
        category: str,
        description: str,
        severity: str,
    ) -> int:
        lower = description.lower()
        if str(severity).lower() == "error":
            return 0
        if (
            category == "structure"
            or "required section" in lower
            or "required declaration" in lower
        ):
            return 1
        if category in {"references", "figures", "tables"} or "equation" in lower:
            return 2
        if category in {
            "title",
            "author_info",
            "headings",
            "layout",
            "margins",
            "journal_header",
        }:
            return 3
        if category in {"body_text", "line_spacing"}:
            return 4
        return 5

    @staticmethod
    def _format_groups(
        groups: List[Dict[str, Any]],
        include_reason: bool = False,
    ) -> List[str]:
        if not groups:
            return ["- None."]
        lines = []
        for group in groups:
            line = (
                f"- {group['description']} "
                f"({group['count']} issue{'s' if group['count'] != 1 else ''})"
            )
            if include_reason and group.get("review_reason"):
                line += f" Reason: {group['review_reason']}"
            lines.append(line)
        return lines

    @staticmethod
    def _friendly_change_summary(group: Dict[str, Any]) -> str:
        change_type = str(group.get("type", "formatting") or "formatting")
        property_name = str(group.get("property", "formatting") or "formatting")
        count = int(group.get("count", 0) or 0)
        suffix = f" ({count} change{'s' if count != 1 else ''})"

        labels = {
            ("body", "alignment"): "Body text alignment was corrected to match the template.",
            ("body_text", "alignment"): "Body text alignment was corrected to match the template.",
            ("body", "space_after"): "Body paragraph spacing after was corrected.",
            ("body_text", "space_after"): "Body paragraph spacing after was corrected.",
            ("page_header", "manual_tabs"): "Page header spacing was adjusted.",
            ("journal_header", "manual_tabs"): "Page header spacing was adjusted.",
            ("author_info", "bold"): "Author information bold formatting was corrected.",
            ("author_info", "font_size"): "Author information font size was corrected.",
            ("author_info", "font_name"): "Author information font was corrected.",
            ("title", "font_size"): "Paper title font size was corrected.",
            ("title", "font_name"): "Paper title font was corrected.",
            ("heading", "bold"): "Section heading bold formatting was corrected.",
            ("headings", "bold"): "Section heading bold formatting was corrected.",
            ("caption", "capitalization"): "Caption capitalization was corrected to journal title case.",
            ("caption", "space_after"): "Caption paragraph spacing after was corrected.",
            ("reference", "font_size"): "Reference font size was corrected.",
            ("references", "font_size"): "Reference font size was corrected.",
            ("reference", "line_spacing"): "Reference line spacing was corrected.",
            ("references", "line_spacing"): "Reference line spacing was corrected.",
            ("reference", "space_after"): "Reference paragraph spacing after was corrected.",
            ("references", "space_after"): "Reference paragraph spacing after was corrected.",
            ("reference", "hanging_indent"): "Reference hanging indent was corrected.",
            ("references", "hanging_indent"): "Reference hanging indent was corrected.",
            ("layout", "page_size"): "Page size was corrected.",
            ("layout", "orientation"): "Page orientation was corrected.",
            ("margins", "margins"): "Page margins were corrected.",
        }
        message = labels.get((change_type, property_name))
        if not message:
            readable_type = change_type.replace("_", " ")
            readable_property = property_name.replace("_", " ")
            message = f"{readable_type.title()} {readable_property} was corrected."
        return f"{message}{suffix}"

    @staticmethod
    def _friendly_property_name(property_name: str) -> str:
        labels = {
            "font_name": "font name",
            "font_size": "font size",
            "line_spacing": "line spacing",
            "space_after": "paragraph spacing after",
            "hanging_indent": "hanging indent",
            "capitalization": "capitalization",
            "manual_tabs": "header spacing",
            "number_bold": "heading number bold formatting",
            "number_font_size": "heading number font size",
            "number_font_name": "heading number font",
            "page_size": "page size",
            "orientation": "page orientation",
            "columns": "page column layout",
            "margins": "page margins",
        }
        return labels.get(
            str(property_name or "formatting"),
            str(property_name or "formatting").replace("_", " "),
        )

    @staticmethod
    def _normalize(value: Any) -> str:
        return re.sub(r"\s+", " ", str(value or "").strip().lower())
