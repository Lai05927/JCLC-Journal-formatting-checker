"""
Template profile loading utilities.

Load and normalize the JCLC template profile.
"""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict

from config import DEFAULT_RULES, REQUIRED_SECTIONS

class ProfileLoader:
    """Load and normalize the JCLC template profile."""

    def __init__(self, profile_dir: Path | None = None):
        base_dir = Path(__file__).resolve().parent.parent
        self.profile_dir = profile_dir or base_dir / "template_profiles"

    def load(self, profile_name: str = "jclc") -> Dict[str, Any]:
        """Load the JCLC profile."""

        profile_path = self.profile_dir / "jclc.json"

        with open(profile_path, "r", encoding="utf-8") as profile_file:
            profile = json.load(profile_file)

        return self._normalize(profile, "jclc")

    def detect_from_document(self, document, template_name: str = "") -> Dict[str, Any]:
        """
        Always use the JCLC profile.
        """
        return self.load("jclc")

    def default_rules(self, profile: Dict[str, Any] | None = None) -> Dict[str, Any]:
        """Return default JCLC rules."""

        selected_profile = profile or self.load("jclc")
        return deepcopy(selected_profile.get("rules", DEFAULT_RULES))

    def apply_rule_defaults(
        self,
        rules: Dict[str, Any],
        profile: Dict[str, Any] | None = None,
    ) -> Dict[str, Any]:
        """Fill missing rule values using the JCLC defaults."""

        merged = self.default_rules(profile)

        for category, values in rules.items():

            if category.startswith("_") or not isinstance(values, dict):
                merged[category] = values
                continue

            category_defaults = merged.get(category, {})

            if not isinstance(category_defaults, dict):
                merged[category] = values
                continue

            updated = deepcopy(category_defaults)

            for field, value in values.items():
                if value is not None:
                    updated[field] = value

            merged[category] = updated

        return merged

    def _normalize(self, profile: Dict[str, Any], profile_name: str) -> Dict[str, Any]:
        """Normalize the JCLC profile."""

        normalized = deepcopy(profile)

        normalized.setdefault("name", "JCLC")
        normalized.setdefault(
            "description",
            "Journal of Communication, Language and Culture profile",
        )
        normalized.setdefault("required_sections", REQUIRED_SECTIONS)
        normalized.setdefault("required_declarations", [])
        normalized.setdefault("heading_patterns", [])
        normalized.setdefault("classification_patterns", {})
        normalized.setdefault("rule_weights", {})
        normalized.setdefault("fallback_defaults", {})

        rules = deepcopy(DEFAULT_RULES)

        profile_rules = normalized.get("rules", {})

        if isinstance(profile_rules, dict):

            for category, values in profile_rules.items():

                if (
                    isinstance(values, dict)
                    and isinstance(rules.get(category), dict)
                ):
                    category_rules = deepcopy(rules[category])
                    category_rules.update(values)
                    rules[category] = category_rules
                else:
                    rules[category] = deepcopy(values)

        fallback_defaults = normalized.get("fallback_defaults", {})

        if isinstance(fallback_defaults, dict):

            font_name = fallback_defaults.get("font_name")

            if font_name:

                for category in [
                    "title",
                    "author",
                    "affiliation",
                    "body",
                    "heading",
                    "subheading",
                    "abstract",
                    "keywords",
                    "caption",
                    "reference",
                ]:

                    if category in rules and isinstance(rules[category], dict):
                        rules[category].setdefault("font_name", font_name)

            fallback_map = {
                "body_font_size": ("body", "font_size"),
                "abstract_font_size": ("abstract", "font_size"),
                "caption_font_size": ("caption", "font_size"),
                "reference_font_size": ("reference", "font_size"),
                "page_size": ("layout", "page_size"),
            }

            for fallback_key, (category, field) in fallback_map.items():

                if fallback_key in fallback_defaults and category in rules:
                    rules[category][field] = fallback_defaults[fallback_key]

        normalized["rules"] = rules

        return normalized