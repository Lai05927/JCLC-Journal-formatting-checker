"""
LLM Integration Module
Provides optional explanation and report-assistance features through NVIDIA API.
"""

import os
import re
import json
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from pathlib import Path

# Try to load .env file
def load_env_file():
    """Load environment variables from .env file"""
    env_paths = [
        Path(__file__).parent.parent / '.env',
        Path.cwd() / '.env'
    ]
    
    for env_path in env_paths:
        if env_path.exists():
            try:
                with open(env_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#') and '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip().strip('"').strip("'")
                            if key and value and key not in os.environ:
                                os.environ[key] = value
                return True
            except Exception:
                pass
    return False

# Load .env file on import
load_env_file()

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

from config import LLM_CONFIG
from .review_guidance import ReviewGuidanceBuilder


@dataclass
class LLMResponse:
    """Response from LLM"""
    success: bool
    content: str
    error: Optional[str] = None


class LLMIntegration:
    """
    Optional LLM integration using NVIDIA API.
    
    Features:
    - Intelligent error explanations
    - Abstract quality analysis
    - Report assistance
    - Writing suggestions
    """
    
    def __init__(self, api_key: str = None):
        """
        Initialize LLM integration with NVIDIA API
        
        Args:
            api_key: NVIDIA API key (optional, can use env var)
        """
        self.api_key = api_key or os.environ.get("NVIDIA_API_KEY", LLM_CONFIG.get("nvidia_api_key", ""))
        self.model = LLM_CONFIG.get("nvidia_model", "deepseek-ai/deepseek-r1")
        self.max_tokens = LLM_CONFIG.get("max_tokens", 1024)
        self.temperature = LLM_CONFIG.get("temperature", 0.3)
        self.timeout_seconds = LLM_CONFIG.get("timeout_seconds", 20)
        self.last_error = ""
        
        self._client = None
        self._available = False
        self._init_client()
    
    def _init_client(self):
        """Initialize the NVIDIA API client"""
        if OPENAI_AVAILABLE and self.api_key:
            try:
                self._client = OpenAI(
                    base_url="https://integrate.api.nvidia.com/v1",
                    api_key=self.api_key,
                    timeout=self.timeout_seconds,
                    max_retries=0
                )
                self._available = True
                self.last_error = ""
            except Exception as exc:
                self._client = None
                self._available = False
                self.last_error = f"{type(exc).__name__}: {exc}"
        else:
            self._available = False
            self.last_error = "OpenAI client package or NVIDIA API key is missing"
    
    def is_available(self) -> bool:
        """Check if LLM is available"""
        return self._available
    
    def generate(self, prompt: str, system_prompt: str = None) -> str:
        """
        Generate a response from the LLM
        
        Args:
            prompt: User prompt
            system_prompt: Optional system prompt
            
        Returns:
            Generated text response
        """
        if not self._available:
            return ""
        
        try:
            messages = []
            
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            
            messages.append({"role": "user", "content": prompt})
            
            response = self._client.chat.completions.create(
                model=self.model,
                messages=messages,
                max_tokens=self.max_tokens,
                temperature=self.temperature,
                top_p=0.7
            )
            
            content = response.choices[0].message.content
            return content.strip() if content else ""
        except Exception as exc:
            self._available = False
            self.last_error = f"{type(exc).__name__}: {exc}"
            return ""
    
    def explain_error(self, issue: Dict[str, Any]) -> str:
        """
        Generate an intelligent explanation for a formatting error
        
        Args:
            issue: Dictionary with keys: description, current_value, expected_value, location
            
        Returns:
            Human-friendly explanation of the error
        """
        if not self._available:
            return self._fallback_explanation(issue)
        
        prompt = f"""Explain this rule-detected manuscript formatting issue.

Important boundaries:
- Do not decide whether the issue is correct.
- The rule-based checker already detected this issue.
- Do not mention AI, model uncertainty, or alternative formatting standards.
- Do not add new issues.
- Keep the explanation short and practical.

Issue details:
Category: {issue.get('category', 'Unknown')}
Location: {issue.get('location', 'Unknown')}
Issue: {issue.get('description', 'Formatting error')}
Current: {issue.get('current_value', 'Unknown')}
Expected: {issue.get('expected_value', 'Unknown')}
Severity: {issue.get('severity', 'warning')}
Text preview: {issue.get('text_preview', '')}

Return exactly this format:
Problem: one sentence describing the mismatch.
Why it matters: one sentence explaining journal/template impact.
How to fix: one sentence with the direct correction.
Rule used: current value -> expected value.
Confidence: explain that this is based on deterministic template rules, not an LLM decision."""

        system_prompt = (
            "You explain formatting issues found by a rule-based manuscript checker. "
            "You are concise, concrete, and do not re-judge the detection."
        )
        
        try:
            response = self.generate(prompt, system_prompt)
            return response if self._is_structured_explanation(response) else self._fallback_explanation(issue)
        except Exception:
            return self._fallback_explanation(issue)

    def _is_structured_explanation(self, explanation: str) -> bool:
        """Return True when the explanation follows the required section format."""
        if not explanation:
            return False
        required_headings = [
            "Problem:",
            "Why it matters:",
            "How to fix:",
            "Rule used:",
            "Confidence:",
        ]
        return all(heading in explanation for heading in required_headings)
    
    def _fallback_explanation(self, issue: Dict[str, Any]) -> str:
        """Fallback explanation when LLM is unavailable"""
        description = issue.get('description', 'Formatting issue')
        current = issue.get('current_value', 'current format')
        expected = issue.get('expected_value', 'expected format')
        location = issue.get('location', 'the selected location')
        severity = issue.get('severity', 'warning')
        
        return (
            f"Problem: {location} has this issue: {description}.\n"
            f"Why it matters: Journal templates require consistent formatting so editors can review submissions quickly.\n"
            f"How to fix: Change the current value from {current} to {expected}.\n"
            f"Rule used: {current} -> {expected}.\n"
            f"Confidence: This is a {severity} from deterministic template rules, not an LLM decision."
        )

    def generate_review_guidance(self, payload: Dict[str, Any]) -> str:
        """Explain grouped pre-fix results without changing their meaning."""
        builder = ReviewGuidanceBuilder()
        fallback = builder.build_pre_fix_fallback(payload)
        if not self._available:
            return fallback

        prompt = (
            "Organize the following rule-detected manuscript issues into practical "
            "review guidance.\n\n"
            "Important boundaries:\n"
            "- Do not add, remove, validate, or contradict issues.\n"
            "- Do not infer manuscript content that is not present in the payload.\n"
            "- Do not change priorities or auto-fix capability.\n"
            "- Keep the response concise and actionable.\n\n"
            "User-facing wording:\n"
            "- Do not show internal field names such as category, severity, priority, "
            "property_name, or auto_fix_supported.\n"
            "- Rewrite grouped issues as plain reviewer actions.\n\n"
            "Return exactly these headings:\n"
            "Most important issues:\n"
            "Safe auto-fix candidates:\n"
            "Needs manual checking:\n"
            "Recommended review order:\n"
            "What this guidance cannot decide:\n\n"
            f"Structured rule results:\n{json.dumps(payload, ensure_ascii=True)}"
        )
        system_prompt = (
            "You explain results produced by deterministic rules. You do not detect "
            "new issues or change compliance decisions. Do not recommend acceptance or rejection."
        )
        response = self.generate(prompt, system_prompt)
        required = [
            "Most important issues:",
            "Safe auto-fix candidates:",
            "Needs manual checking:",
            "Recommended review order:",
            "What this guidance cannot decide:",
        ]
        return response if self._has_required_headings(response, required) else fallback

    def generate_post_fix_summary(self, payload: Dict[str, Any]) -> str:
        """Explain applied fixes and remaining rule-detected issues."""
        builder = ReviewGuidanceBuilder()
        fallback = builder.build_post_fix_fallback(payload)
        if not self._available:
            return fallback

        prompt = (
            "Explain the following deterministic post-fix results.\n\n"
            "Important boundaries:\n"
            "- You must not reinterpret change records, checker results, or safety status.\n"
            "- Do not add new issues or claim that a manuscript is publication-ready.\n"
            "- Explain why manual-review items were not changed automatically.\n\n"
            "User-facing wording:\n"
            "- Do not show internal field names such as category, severity, priority, "
            "property_name, or auto_fix_supported.\n"
            "- Rewrite grouped issues as plain reviewer actions.\n\n"
            "Return exactly these headings:\n"
            "Auto-fixed items:\n"
            "Issues still needing review:\n"
            "Why these were not auto-fixed:\n"
            "What to check next:\n"
            "Auto-fix safety check:\n\n"
            f"Structured post-fix results:\n{json.dumps(payload, ensure_ascii=True)}"
        )
        system_prompt = (
            "You explain post-fix results produced by deterministic rules. "
            "You do not alter safety status or recommend acceptance or rejection."
        )
        response = self.generate(prompt, system_prompt)
        required = [
            "Auto-fixed items:",
            "Issues still needing review:",
            "Why these were not auto-fixed:",
            "What to check next:",
            "Auto-fix safety check:",
        ]
        return response if self._has_required_headings(response, required) else fallback

    @staticmethod
    def _has_required_headings(
        response: str,
        required_headings: List[str],
    ) -> bool:
        """Return True only for complete structured guidance."""
        forbidden_markers = [
            "category:",
            "severity:",
            "priority:",
            "property_name",
            "auto_fix_supported",
        ]
        lower_response = (response or "").lower()
        return bool(response) and not any(
            marker in lower_response for marker in forbidden_markers
        ) and all(
            heading in response
            for heading in required_headings
        )
    
    def analyze_template_rules(self, paragraphs_info: List[Dict]) -> Dict[str, Any]:
        """
        Use AI as an optional fallback to suggest missing template rules.
        Deterministic extraction remains the primary source of rules.
        
        Args:
            paragraphs_info: List of dicts with 'text', 'font', 'size', 'bold', 'italic' info
            
        Returns:
            Dictionary of extracted rules for each element type
        """
        if not self._available:
            return {}
        
        # Format paragraphs for AI analysis - include more context
        formatted = []
        for i, p in enumerate(paragraphs_info[:40]):  # Analyze first 40 paragraphs
            text = p.get('text', '')[:100]
            font = p.get('font', 'Unknown')
            size = p.get('size', '?')
            bold = p.get('bold', 'Unknown')
            italic = p.get('italic', False)
            formatted.append(f"{i+1}. [{font}, {size}pt, italic={italic}] \"{text}\"")
        
        prompt = f"""Analyze this academic template and extract formatting rules.

IMPORTANT: DISTINGUISH BETWEEN JOURNAL TITLE AND PAPER TITLE!

In academic templates, there are TWO different titles:
1. JOURNAL TITLE - The name of the journal (e.g., "Journal of Informatics and Web Engineering")
   - Usually has "(Journal Title)" label
   - Often uses Palatino Linotype font
   - IGNORE THIS - we don't format journal names
   
2. PAPER TITLE - The title of the research paper itself (e.g., "Preparation template for...")
   - Usually has "(Title)" label WITHOUT "Journal"
   - This is what we need to extract rules for!
   - Look for instructions like "(24-Font size, Times New Roman)" - NO BOLD mentioned = bold: false

EXAMPLES OF INSTRUCTION TEXT PATTERNS:

PATTERN 1: "(24-Font size, bold Palatino Linotype)" - used for JOURNAL title
- Font variant: "Palatino Linotype Bold"  
- Bold formatting: FALSE (bold is part of font name, not a format instruction)

PATTERN 2: "(24-Font size, Times New Roman)" - used for PAPER title  
- Font: Times New Roman
- Bold: FALSE (bold is NOT mentioned at all!)

PATTERN 3: "(10-Font size, bold, Times New Roman)" - note the comma before bold
- Bold: TRUE (", bold," with commas = format instruction)

KEY RULES:
- "bold [FontName]" = font variant name, NOT bold formatting
- "(X-Font size, FontName)" with NO "bold" mentioned = bold: false
- Only ", bold," or ", bold)" with commas means bold: true

PARAGRAPHS FROM TEMPLATE:
{chr(10).join(formatted)}

For the "title" field, extract rules for the PAPER TITLE (marked with "(Title)"), NOT the journal title!

Return ONLY valid JSON:
{{
    "title": {{"font": "FontName", "size": NUM, "bold": false, "italic": false}},
    "heading": {{"font": "FontName", "size": NUM, "bold": BOOL, "italic": BOOL}},
    "body": {{"font": "FontName", "size": NUM, "bold": false, "italic": false}},
    "abstract": {{"font": "FontName", "size": NUM, "bold": false, "italic": false}},
    "reference": {{"font": "FontName", "size": NUM, "bold": false, "italic": false}},
    "caption": {{"font": "FontName", "size": NUM, "bold": false, "italic": BOOL}}
}}

CRITICAL FOR JIWE TEMPLATE: 
- Paper Title instruction is "(24-Font size, Times New Roman)" 
- This means: font="Times New Roman", size=24, bold=FALSE (bold not mentioned!)"""

        system_prompt = """You are analyzing academic template formatting instructions.
CRITICAL: Look for "(Title)" to find PAPER title rules, ignore "(Journal Title)".
For "(24-Font size, Times New Roman)" - bold is FALSE because it's not mentioned.
Only return bold=true if there's ", bold," with commas."""
        
        try:
            response = self.generate(prompt, system_prompt)
            # Parse JSON from response
            import json
            import re
            
            response = response.strip()
            
            # Try to extract JSON from various formats
            # Method 1: Direct parse
            try:
                rules = json.loads(response)
                rules['_ai_extracted'] = True
                return rules
            except json.JSONDecodeError:
                pass
            
            # Method 2: Remove markdown code blocks
            if response.startswith('```'):
                lines = response.split('\n')
                # Remove first and last lines if they are markdown markers
                if lines[0].startswith('```'):
                    lines = lines[1:]
                if lines and lines[-1].strip() == '```':
                    lines = lines[:-1]
                response = '\n'.join(lines)
                try:
                    rules = json.loads(response)
                    rules['_ai_extracted'] = True
                    return rules
                except json.JSONDecodeError:
                    pass
            
            # Method 3: Find JSON object in response using regex
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                try:
                    rules = json.loads(json_match.group())
                    rules['_ai_extracted'] = True
                    return rules
                except json.JSONDecodeError:
                    pass
            
            return {}
        except Exception:
            return {}
    
def create_llm_integration(api_key: str = None) -> LLMIntegration:
    """
    Factory function to create LLM integration
    
    Args:
        api_key: NVIDIA API key
        
    Returns:
        LLMIntegration instance
    """
    return LLMIntegration(api_key=api_key)


def fallback_explain_issue(issue: Dict[str, Any]) -> str:
    """Return the deterministic explanation used when no LLM is available."""
    llm = LLMIntegration.__new__(LLMIntegration)
    llm._available = False
    return LLMIntegration._fallback_explanation(llm, issue)
