"""
Academic Manuscript Format Checker
A Streamlit web application for checking and fixing manuscript formatting

Main Application Entry Point
"""

import streamlit as st
import html
import inspect
import os
import subprocess
import traceback
from io import BytesIO
from docx import Document
from datetime import datetime
from pathlib import Path
# Import modules
from modules.template_extractor import TemplateExtractor
from modules.manuscript_checker import ManuscriptChecker
from modules.auto_fixer import (
    AutoFixer,
    summarize_remaining_issues,
    validate_fixed_document,
    validate_post_fix_result,
)
from modules.report_generator import ReportGenerator
from modules.llm_integration import create_llm_integration, fallback_explain_issue
from modules.profile_loader import ProfileLoader
from modules.review_guidance import ReviewGuidanceBuilder
from modules.utils import apply_template_first_page_header, apply_template_first_page_footer
from config import APP_TITLE, APP_VERSION

# Page configuration
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)


def get_build_commit() -> str:
    """Return the current Git commit short hash when available."""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=2,
            check=True,
        )
        return result.stdout.strip()
    except Exception:
        return "unknown"


def get_default_template_bytes():
    """Return the bundled JCLC template's own file bytes, used to
    replicate structural elements (the first-page header/footer's exact
    drawing layout) that a rules dict alone can't express."""
    template_path = Path(__file__).resolve().parent / "template_profiles" / "assets" / "jclc_template.docx"
    if template_path.exists():
        return template_path.read_bytes()
    return None


def get_default_template_rules():
    """Return the validated JCLC default rules used when no template is uploaded."""
    loader = ProfileLoader()
    profile = loader.load("jclc")
    rules = loader.default_rules(profile)
    # The classifier reads its heading/declaration term lists and other
    # classification data from rules["_profile"] (see ManuscriptChecker.
    # __init__: self.profile = rules.get("_profile", {})) - it needs the
    # *whole* loaded profile dict here (heading_patterns,
    # classification_patterns, required_sections, ...), not just a
    # name/source summary. Passing only a name/source stub here was the
    # actual bug: declaration-style headings like "Author Contribution
    # Statement" or "Data Access Statement" have no numbered prefix to
    # fall back on, so without heading_patterns present the classifier
    # had no way to recognize them as headings at all when no template
    # was uploaded, even though the exact same manuscript classified
    # correctly through the upload-a-template path (which does build a
    # complete _profile).
    rules["_profile"] = dict(profile)
    rules["_profile"]["source"] = "default_profile"
    return rules


def get_local_env_value(name: str) -> str:
    """Read a single key from a local .env file for local development."""
    env_paths = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent / ".env",
    ]
    seen = set()
    for env_path in env_paths:
        if env_path in seen or not env_path.exists():
            continue
        seen.add(env_path)
        try:
            for line in env_path.read_text(encoding="utf-8").splitlines():
                stripped = line.strip()
                if not stripped or stripped.startswith("#") or "=" not in stripped:
                    continue
                key, value = stripped.split("=", 1)
                if key.strip() == name:
                    return value.strip().strip('"').strip("'")
        except OSError:
            continue
    return ""


def get_server_nvidia_api_key() -> str:
    """Return a server-managed NVIDIA API key without exposing it in the UI."""
    key_names = [
        "NVIDIA_API_KEY",
        "NVIDIA_NIM_API_KEY",
        "NVIDIA_API_TOKEN",
        "nvidia_api_key",
        "nvidia_nim_api_key",
        "nvidia_api_token",
    ]
    secret_candidates = []
    try:
        secret_candidates.extend(st.secrets.get(name, "") for name in key_names)
        llm_secrets = st.secrets.get("llm", {})
        secret_candidates.extend(llm_secrets.get(name, "") for name in key_names)
    except Exception:
        pass

    secret_candidates.extend(os.environ.get(name, "") for name in key_names)
    secret_candidates.extend(get_local_env_value(name) for name in key_names)
    return next((key for key in secret_candidates if key), "")


def ensure_llm_connection() -> bool:
    """Connect to the server-managed LLM once when AI explanations are enabled."""
    if st.session_state.llm and st.session_state.llm.is_available():
        st.session_state.llm_connection_error = ""
        return True
    if st.session_state.get("llm_connection_attempted"):
        return False

    st.session_state.llm_connection_attempted = True
    api_key = get_server_nvidia_api_key()
    if not api_key:
        st.session_state.llm = None
        st.session_state.llm_connection_error = "missing_key"
        return False

    st.session_state.llm = create_llm_integration(api_key=api_key)
    if st.session_state.llm.is_available():
        st.session_state.llm_source = "server"
        st.session_state.llm_connection_error = ""
        return True

    st.session_state.llm = None
    st.session_state.llm_connection_error = "connection_failed"
    return False


def get_llm_status_notice(
    ai_enabled: bool,
    connected: bool,
    key_configured: bool,
    attempted: bool,
):
    """Return a user-facing AI status level and message."""
    if not ai_enabled:
        return (
            "info",
            "AI explanations are disabled. Core checking is rule-based.",
        )
    if connected:
        return (
            "success",
            "AI explanations are connected. AI only explains detected issues; format checking remains rule-based.",
        )
    if not key_configured:
        return (
            "warning",
            "AI explanation service is not configured on this deployment. You do not need to enter an API key; rule-based explanations will still be shown.",
        )
    if attempted:
        return (
            "warning",
            "AI explanation service could not connect with the configured server key. Rule-based explanations will still be shown. After updating Streamlit Secrets, use Retry AI connection.",
        )
    return (
        "info",
        "AI explanations will connect through the server key when needed. Core checking remains rule-based.",
    )


def run_llm_smoke_test(llm):
    """Run a tiny generation request to prove that the AI explanation service works."""
    if not llm or not llm.is_available():
        return (
            "warning",
            "AI explanation service is not connected yet. Check the server key and retry the AI connection.",
        )
    try:
        response = llm.generate(
            "Reply with exactly AI_READY.",
            "You are a service health checker. Return only the requested token.",
        )
    except Exception as exc:
        return (
            "warning",
            f"AI test request failed: {type(exc).__name__}. Rule-based explanations are still available.",
        )

    normalized_response = response.strip().strip(" .,!?:;\"'`")
    if normalized_response == "AI_READY":
        return (
            "success",
            "AI explanation service returned a valid test response. AI guidance buttons can use the server model.",
        )
    if not response.strip():
        last_error = getattr(llm, "last_error", "")
        if last_error:
            return (
                "warning",
                f"AI test request failed: {last_error}. Rule-based explanations are still available.",
            )
        return (
            "warning",
            "AI service connected but did not return a test response. Rule-based explanations will still be shown.",
        )
    return (
        "warning",
        "AI service returned an unexpected response. Rule-based explanations will still be shown.",
    )


def get_issue_explanation_button_label(ai_enabled: bool, connected: bool) -> str:
    """Return the issue explanation button label for the current AI state."""
    if ai_enabled and connected:
        return "Explain with AI"
    return "Explain with Rule-Based Guidance"


def get_review_guidance_mode_notice(
    ai_enabled: bool,
    connected: bool,
    key_configured: bool,
):
    """Return a clear notice explaining which guidance source will be used."""
    if connected:
        return (
            "success",
            "AI-enhanced guidance is available. AI explains grouped rule results; it does not detect new issues.",
        )
    if ai_enabled and key_configured:
        return (
            "warning",
            "Rule-based guidance will be used because AI is enabled but not connected.",
        )
    if ai_enabled:
        return (
            "warning",
            "Rule-based guidance will be used because no server AI key is configured.",
        )
    return (
        "info",
        "Rule-based guidance is available. Enable AI explanations only if you want AI-written wording.",
    )


def get_system_capability_sections():
    """Return user-facing capability and limitation text for the sidebar."""
    return [
        {
            "title": "Best Supported Use",
            "items": [
                "Primary journal profile: JCLC manuscript formatting in DOCX format.",
                "Checks are rule-based and compare detected manuscript properties against uploaded template rules or the default JCLC profile.",
                "Useful for pre-screening formatting issues before editor or supervisor review.",
            ],
        },
        {
            "title": "Can Detect",
            "items": [
                "Margins, page size, journal header, paper title, author information, body text, headings, abstract rules, keywords, captions, references, citations, and required sections.",
                "JCLC-specific checks include paragraph spacing after, caption title case, reference line spacing, and reference hanging indent.",
                "Some checks are exact formatting checks; others are warning-level structural or content-pattern checks that should be reviewed by a person.",
            ],
        },
        {
            "title": "Auto-Fix Can Change",
            "items": [
                "Only rule-detected formatting properties such as font, font size, bold, italic, alignment, line spacing, spacing after, hanging indent, margins, page size, capitalization, and stable header spacing.",
                "Auto-Fix re-checks the corrected document and reports remaining issues after supported changes are applied.",
            ],
        },
        {
            "title": "Needs Manual Review",
            "items": [
                "Missing sections, missing captions, figure/table movement, equation numbering, citation meaning, reference source selection, author identity, and research content.",
                "The system does not decide manuscript acceptance, research quality, novelty, or language quality.",
            ],
        },
        {
            "title": "AI Boundaries",
            "items": [
                "Manuscripts must be submitted as native Word (.docx) files.",
                "AI explanations are optional. AI explains rule-detected issues only; it does not perform core checking, auto-fix, or approval.",
            ],
        },
    ]


def format_section_label(section_name: str) -> str:
    """Return a readable label for a document section key."""
    return str(section_name).replace("_", " ").title()


def format_structure_evidence(status: str) -> str:
    """Return user-friendly structure evidence text."""
    labels = {
        "valid": "Clear heading",
        "weak": "Needs review",
        "not_checked": "Not checked",
    }
    return labels.get(str(status), format_section_label(status))


def build_structure_summary_items(structure):
    """Build user-friendly document structure summary items."""
    structure_details = structure if isinstance(structure.get("sections"), dict) else None
    sections = (
        structure_details.get("expected_order", [])
        if structure_details else ["abstract", "keywords", "introduction", "conclusion", "references"]
    )
    if structure_details:
        found_map = {
            name: details.get("found", False)
            for name, details in structure_details.get("sections", {}).items()
        }
    else:
        found_map = structure

    return [
        {
            "label": format_section_label(section),
            "status": "Found" if found_map.get(section, False) else "Missing",
        }
        for section in sections
    ]


def build_issue_display_record(issue, category):
    """Build one sanitized issue record for UI display and explanations."""
    return ReviewGuidanceBuilder().build_issue_evidence(issue, category)


def build_auto_fix_preview_record(issues_by_category):
    """Build a user-facing preview of supported and manual auto-fix work."""
    return ReviewGuidanceBuilder().build_auto_fix_preview(issues_by_category)


def build_pre_fix_guidance_payload(result, rules):
    """Build privacy-limited pre-fix guidance input."""
    return ReviewGuidanceBuilder().build_pre_fix_payload(result, rules)


def resolve_pre_fix_guidance(payload, llm, cache):
    """Return cached AI-enhanced or deterministic pre-fix guidance."""
    builder = ReviewGuidanceBuilder()
    cache_key = builder.cache_key(payload, "pre-fix")
    if cache_key in cache:
        cached = cache[cache_key]
        return cached["text"], cache_key, cached["source"]

    fallback = builder.build_pre_fix_fallback(payload)
    if llm and llm.is_available():
        guidance = llm.generate_review_guidance(payload)
        source = (
            "Rule-based fallback after AI response was unavailable or incomplete"
            if guidance == fallback
            else "AI-enhanced guidance"
        )
    else:
        guidance = fallback
        source = "Rule-based guidance"

    cache[cache_key] = {"text": guidance, "source": source}
    while len(cache) > 20:
        cache.pop(next(iter(cache)))
    return guidance, cache_key, source


def build_post_fix_guidance_payload(
    before_result,
    after_result,
    changes,
    validation,
    rules,
):
    """Build privacy-limited post-fix guidance input."""
    return ReviewGuidanceBuilder().build_post_fix_payload(
        before_result,
        after_result,
        changes,
        validation,
        rules,
    )


def resolve_post_fix_guidance(payload, llm, cache):
    """Return cached AI-enhanced or deterministic post-fix guidance."""
    builder = ReviewGuidanceBuilder()
    cache_key = builder.cache_key(payload, "post-fix")
    if cache_key in cache:
        cached = cache[cache_key]
        return cached["text"], cache_key, cached["source"]

    fallback = builder.build_post_fix_fallback(payload)
    if llm and llm.is_available():
        guidance = llm.generate_post_fix_summary(payload)
        source = (
            "Rule-based fallback after AI response was unavailable or incomplete"
            if guidance == fallback
            else "AI-enhanced guidance"
        )
    else:
        guidance = fallback
        source = "Rule-based guidance"

    cache[cache_key] = {"text": guidance, "source": source}
    while len(cache) > 20:
        cache.pop(next(iter(cache)))
    return guidance, cache_key, source


def display_auto_fix_preview(preview):
    """Display what auto-fix can and cannot safely handle."""
    st.info(preview["summary"])
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("Detected Issues", preview["total_issues"])
    metric_col2.metric("Auto-Fix Candidates", preview["supported_count"])
    metric_col3.metric("Manual Review", preview["manual_count"])

    supported_groups = preview.get("supported_groups", [])
    manual_groups = preview.get("manual_groups", [])
    if supported_groups:
        with st.expander("What Auto-Fix Can Change", expanded=False):
            for group in supported_groups:
                st.write(
                    "- "
                    f"{group['description']} "
                    f"({group['count']} issue"
                    f"{'s' if group['count'] != 1 else ''})"
                )
    if manual_groups:
        with st.expander("What Still Needs Manual Review", expanded=False):
            for group in manual_groups:
                reason = group.get("review_reason") or "Manual checking is required."
                st.write(
                    "- "
                    f"{group['description']} "
                    f"({group['count']} issue"
                    f"{'s' if group['count'] != 1 else ''})"
                    f" Reason: {reason}"
                )


# Load custom CSS
def load_css():
    css_path = os.path.join(os.path.dirname(__file__), "assets", "styles.css")
    if os.path.exists(css_path):
        with open(css_path) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    else:
        # Inline CSS if file not found
        st.markdown("""
        <style>
            .main-header {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                padding: 20px;
                border-radius: 10px;
                color: white;
                margin-bottom: 20px;
            }
            .score-card {
                padding: 20px;
                border-radius: 10px;
                text-align: center;
                margin: 10px 0;
            }
            .score-good { background: #d4edda; border: 2px solid #28a745; }
            .score-medium { background: #fff3cd; border: 2px solid #ffc107; }
            .score-poor { background: #f8d7da; border: 2px solid #dc3545; }
            .issue-card {
                padding: 15px;
                border-radius: 8px;
                margin: 10px 0;
                border-left: 4px solid;
            }
            .issue-error { background: #fff5f5; border-color: #dc3545; }
            .issue-warning { background: #fffbf0; border-color: #ffc107; }
            .rules-display {
                background: #f8f9fa;
                padding: 15px;
                border-radius: 8px;
                margin: 10px 0;
            }
            .stTabs [data-baseweb="tab-list"] {
                gap: 10px;
            }
            .stTabs [data-baseweb="tab"] {
                padding: 10px 20px;
                border-radius: 5px;
            }
        </style>
        """, unsafe_allow_html=True)

load_css()


def init_session_state():
    """Initialize session state variables"""
    if "template_rules" not in st.session_state:
        st.session_state.template_rules = None
    if "template_bytes" not in st.session_state:
        st.session_state.template_bytes = None
    if "template_uploaded" not in st.session_state:
        st.session_state.template_uploaded = False
    if "manuscript_uploaded" not in st.session_state:
        st.session_state.manuscript_uploaded = False
    if "check_result" not in st.session_state:
        st.session_state.check_result = None
    if "classifications" not in st.session_state:
        st.session_state.classifications = []
    if "changes" not in st.session_state:
        st.session_state.changes = []
    if "fixed_doc_bytes" not in st.session_state:
        st.session_state.fixed_doc_bytes = None
    if "report_bytes" not in st.session_state:
        st.session_state.report_bytes = None
    if "post_fix_result" not in st.session_state:
        st.session_state.post_fix_result = None
    if "post_fix_validation" not in st.session_state:
        st.session_state.post_fix_validation = None
    if "llm" not in st.session_state:
        st.session_state.llm = None
    if "llm_connection_attempted" not in st.session_state:
        st.session_state.llm_connection_attempted = False
    if "llm_source" not in st.session_state:
        st.session_state.llm_source = ""
    if "llm_connection_error" not in st.session_state:
        st.session_state.llm_connection_error = ""
    if "llm_smoke_test_notice" not in st.session_state:
        st.session_state.llm_smoke_test_notice = None
    if "ai_explanations_enabled" not in st.session_state:
        st.session_state.ai_explanations_enabled = False
    if "manuscript_bytes" not in st.session_state:
        st.session_state.manuscript_bytes = None
    if "manuscript_filename" not in st.session_state:
        st.session_state.manuscript_filename = "manuscript.docx"
    if "highlighted_doc_bytes" not in st.session_state:
        st.session_state.highlighted_doc_bytes = None
    if "output_timestamp" not in st.session_state:
        st.session_state.output_timestamp = ""
    if "review_guidance_cache" not in st.session_state:
        st.session_state.review_guidance_cache = {}
    if "pre_fix_guidance" not in st.session_state:
        st.session_state.pre_fix_guidance = ""
    if "pre_fix_guidance_key" not in st.session_state:
        st.session_state.pre_fix_guidance_key = ""
    if "pre_fix_guidance_source" not in st.session_state:
        st.session_state.pre_fix_guidance_source = ""
    if "post_fix_guidance" not in st.session_state:
        st.session_state.post_fix_guidance = ""
    if "post_fix_guidance_key" not in st.session_state:
        st.session_state.post_fix_guidance_key = ""
    if "post_fix_guidance_source" not in st.session_state:
        st.session_state.post_fix_guidance_source = ""


def display_header():
    """Display application header"""
    st.markdown("""
    <div class="main-header">
        <h1 style="margin: 0;">Automated Manuscript Template Compliance Checker</h1>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">
            Automatically check and fix manuscript formatting against journal templates
        </p>
    </div>
    """, unsafe_allow_html=True)


def display_exception(summary: str, error: Exception):
    """Display a user-safe error with optional technical details."""
    st.error(f"{summary}: {str(error)}")
    with st.expander("Technical details", expanded=False):
        st.code(traceback.format_exc())


def display_sidebar():
    """Display sidebar with template rules and settings"""
    with st.sidebar:
        st.header("Template Rules")

        # LLM Configuration
        with st.expander("AI Explanation Settings", expanded=False):
            st.caption(
                "AI is optional and is used only for issue explanations. "
                "Template extraction, compliance checking, auto-fix, and accuracy evaluation remain rule-based."
            )
            ai_enabled = st.checkbox(
                "Enable AI explanations",
                value=st.session_state.ai_explanations_enabled,
                help="Turn this on only when you want AI-generated explanations for detected issues."
            )
            st.session_state.ai_explanations_enabled = ai_enabled

            if not ai_enabled:
                st.session_state.llm = None
                st.session_state.llm_connection_attempted = False
                st.session_state.llm_source = ""
                st.session_state.llm_connection_error = ""
                st.session_state.llm_smoke_test_notice = None

            # Show current status
            key_configured = bool(get_server_nvidia_api_key()) if ai_enabled else False
            llm_connected = bool(ai_enabled and ensure_llm_connection())
            status_level, status_message = get_llm_status_notice(
                ai_enabled=ai_enabled,
                connected=llm_connected,
                key_configured=key_configured,
                attempted=bool(st.session_state.get("llm_connection_attempted")),
            )
            getattr(st, status_level)(status_message)

            if llm_connected:
                if st.button("Test AI explanation service"):
                    st.session_state.llm_smoke_test_notice = run_llm_smoke_test(
                        st.session_state.llm
                    )
                    st.rerun()
                if st.session_state.llm_smoke_test_notice:
                    test_level, test_message = st.session_state.llm_smoke_test_notice
                    getattr(st, test_level)(test_message)
                if st.button("Disconnect LLM"):
                    st.session_state.llm = None
                    st.session_state.llm_connection_attempted = False
                    st.session_state.llm_source = ""
                    st.session_state.llm_connection_error = ""
                    st.session_state.llm_smoke_test_notice = None
                    st.session_state.ai_explanations_enabled = False
                    st.rerun()
            elif ai_enabled:
                if st.button("Retry AI connection"):
                    st.session_state.llm = None
                    st.session_state.llm_connection_attempted = False
                    st.session_state.llm_connection_error = ""
                    st.session_state.llm_smoke_test_notice = None
                    st.rerun()
                if os.environ.get("SHOW_LLM_KEY_OVERRIDE") == "1":
                    with st.expander("Developer key override", expanded=False):
                        st.caption("Optional for local testing only. Normal users do not need to enter an API key.")
                        api_key = st.text_input(
                            "NVIDIA API Key",
                            type="password",
                            placeholder="Paste a local testing key",
                            help="Use only when the server-managed key is not configured."
                        )
                        if st.button("Connect with override key", type="primary"):
                            if api_key:
                                st.session_state.llm = create_llm_integration(api_key=api_key)
                                st.session_state.llm_connection_attempted = True
                                if st.session_state.llm.is_available():
                                    st.session_state.llm_source = "override"
                                    st.session_state.llm_connection_error = ""
                                    st.success("Connected successfully.")
                                    st.rerun()
                                else:
                                    st.session_state.llm = None
                                    st.session_state.llm_connection_error = "connection_failed"
                                    st.error("Connection failed. Check the override key.")
                            else:
                                st.error("Please enter an API key")

        with st.expander("System Capability / Limitation", expanded=False):
            st.caption(
                "Use this section to understand what the checker can verify, "
                "what Auto-Fix can safely change, and what still needs human review."
            )
            for section in get_system_capability_sections():
                st.markdown(f"**{section['title']}**")
                for item in section["items"]:
                    st.write(f"- {item}")

        st.divider()

        # Display extracted rules if available
        if st.session_state.template_rules:
            rules = st.session_state.template_rules
            extraction_summary = rules.get("_extraction_summary", {})
            profile = rules.get("_profile", {})
            if extraction_summary:
                st.info(
                    f"Profile: {profile.get('name', 'Generic')} | "
                    f"{extraction_summary.get('extracted', 0)} extracted, "
                    f"{extraction_summary.get('inferred', 0)} inferred, "
                    f"{extraction_summary.get('default', 0)} defaulted"
                )

            # Show AI enhancement status
            if rules.get("_ai_enhanced"):
                st.success("AI-enhanced extraction")

            st.subheader("Extracted Text Styles")

            # Title style
            st.markdown("**Title Style:**")
            title = rules.get("title", {})
            bold_display = "Yes" if title.get('bold') else ("No" if title.get('bold') is False else "Auto")
            st.caption(f"Font: {title.get('font_name', 'Times New Roman')} | Size: {title.get('font_size', 24)}pt | Bold: {bold_display}")

            # Body style
            st.markdown("**Body Text Style:**")
            body = rules.get("body", {})
            body_spacing = body.get("space_after")
            spacing_text = f" | After: {body_spacing}pt" if body_spacing is not None else ""
            st.caption(
                f"Font: {body.get('font_name', 'Times New Roman')} | "
                f"Size: {body.get('font_size', 12)}pt{spacing_text}"
            )

            # Heading style
            st.markdown("**Heading Style:**")
            heading = rules.get("heading", {})
            bold_display = "Yes" if heading.get('bold') else ("No" if heading.get('bold') is False else "Auto")
            st.caption(f"Font: {heading.get('font_name', 'Times New Roman')} | Size: {heading.get('font_size', 14)}pt | Bold: {bold_display}")

            # Show margins in collapsed expander (optional viewing)
            with st.expander("Page Margins (Advanced)", expanded=False):
                margins = rules.get("margins", {})
                st.caption(f"L: {margins.get('left', 1.0):.2f}in | R: {margins.get('right', 1.0):.2f}in | T: {margins.get('top', 1.0):.2f}in | B: {margins.get('bottom', 1.0):.2f}in")
        else:
            st.info("The official JCLC template is built in.")

        st.divider()

        # About section
        with st.expander("About"):
            st.markdown(f"""
            **{APP_TITLE}**
            Version: {APP_VERSION}
            Build: {get_build_commit()}

            This tool helps you:
                        - Check manuscripts against templates
            - Auto-fix formatting issues
            - Generate comparison reports
            """)


def handle_template_upload(uploaded_file):
    """Handle template file upload and rule extraction"""
    if uploaded_file:
        try:
            file_bytes = uploaded_file.read()
            original_filename = uploaded_file.name

            with st.spinner("Extracting formatting rules from template..."):
                extractor = TemplateExtractor(llm_integration=None)
                extractor.load(BytesIO(file_bytes), template_name=original_filename)
                rules = extractor.extract_all_rules()

                st.session_state.template_rules = rules
                st.session_state.template_bytes = file_bytes
                st.session_state.template_uploaded = True

                summary = rules.get("_extraction_summary", {})
                profile = rules.get("_profile", {})
                st.success(
                    "Template rules processed: "
                    f"{summary.get('extracted', 0)} extracted, "
                    f"{summary.get('inferred', 0)} inferred, "
                    f"{summary.get('default', 0)} defaulted."
                )
                if str(profile.get("name", "")).upper() == "JCLC":
                    st.info(
                        "Validated profile: JCLC. Auto-Fix remains conservative and only changes deterministic formatting issues."
                    )
                else:
                    st.warning(
                        "This template was not matched to the validated JCLC profile. "
                        "Checking can still run, but auto-fix results should be manually validated."
                    )

                # Display summary - user-friendly view
                with st.expander("View Extracted Rules Summary", expanded=True):
                    st.text(extractor.get_rules_summary())

        except Exception as e:
            display_exception("Error extracting template rules", e)


def handle_manuscript_check(uploaded_file):
    """Handle manuscript upload and format checking"""
    if uploaded_file:
        try:
            # Store manuscript bytes and filename for later use
            manuscript_bytes = uploaded_file.read()
            original_filename = uploaded_file.name

            # Store original filename (without extension) for download
            base_name = original_filename.rsplit('.', 1)[0] if '.' in original_filename else original_filename
            st.session_state.manuscript_filename = base_name

            # PDF upload/conversion support has been removed - author
            # submissions are always native Word documents, never a PDF
            # run through pdf2docx, so the whole pdf-conversion-artifact
            # repair pass (manual list markers needing real numbering,
            # a masthead table needing to move into the first-page
            # footer, stray page-number columns, section-per-page
            # structure, ...) no longer applies and is skipped entirely -
            # a real .docx from Word never has these artifacts in the
            # first place, so there's nothing here for that pass to fix,
            # only processing time and edge-case risk to spend on it.
            st.session_state.manuscript_bytes = manuscript_bytes

            with st.spinner("Analyzing manuscript formatting..."):
                # Use template rules or defaults
                rules = st.session_state.template_rules or get_default_template_rules()

                # Warn user if using default rules

                # Core checking is rule-based. LLM is not used for compliance decisions.
                checker = ManuscriptChecker(rules, None)
                checker.load_manuscript(BytesIO(manuscript_bytes))

                # Run all checks
                result = checker.check_all()

                st.session_state.check_result = result
                st.session_state.classifications = result.classifications
                st.session_state.manuscript_uploaded = True
                st.session_state.changes = []
                st.session_state.fixed_doc_bytes = None
                st.session_state.highlighted_doc_bytes = None
                st.session_state.report_bytes = None
                st.session_state.post_fix_result = None
                st.session_state.post_fix_validation = None
                st.session_state.pre_fix_guidance = ""
                st.session_state.pre_fix_guidance_key = ""
                st.session_state.pre_fix_guidance_source = ""
                st.session_state.post_fix_guidance = ""
                st.session_state.post_fix_guidance_key = ""
                st.session_state.post_fix_guidance_source = ""

                return result

        except Exception as e:
            display_exception("Error checking manuscript", e)

    return None


def display_check_results(result):
    """Display format check results"""
    st.header("Format Check Results")

    # Compliance score card
    score = result.compliance_score
    if score >= 80:
        score_class = "score-good"
        score_emoji = ""
    elif score >= 60:
        score_class = "score-medium"
        score_emoji = ""
    else:
        score_class = "score-poor"
        score_emoji = ""
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(f"""
        <div class="score-card {score_class}">
            <h2 style="margin: 0;">{score}%</h2>
            <p style="margin: 5px 0 0 0;">Compliance Index</p>
        </div>
        """, unsafe_allow_html=True)
        st.caption("UI index only. FYP accuracy is evaluated with Precision, Recall, and F1.")

    with col2:
        st.markdown(f"""
        <div class="score-card" style="background: #e3f2fd; border: 2px solid #2196f3;">
            <h2 style="margin: 0;">{result.total_issues}</h2>
            <p style="margin: 5px 0 0 0;">Issues Found</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        stats = result.statistics
        st.markdown(f"""
        <div class="score-card" style="background: #f3e5f5; border: 2px solid #9c27b0;">
            <h2 style="margin: 0;">{stats.get('total_words', 0)}</h2>
            <p style="margin: 5px 0 0 0;">Total Words</p>
        </div>
        """, unsafe_allow_html=True)

    # Document structure
    st.subheader("Document Structure")
    structure = result.document_structure
    structure_details = structure if isinstance(structure.get("sections"), dict) else None
    if structure_details:
        structure = {
            name: details.get("found", False)
            for name, details in structure_details.get("sections", {}).items()
        }

    summary_items = build_structure_summary_items(structure)
    sections = (
        structure_details.get("expected_order", [])
        if structure_details else [item["label"].lower() for item in summary_items]
    )
    structure_cols = st.columns(max(1, len(summary_items)))

    for i, item in enumerate(summary_items):
        with structure_cols[i]:
            st.markdown(f"**{item['label']}**")
            st.write(item["status"])

    if structure_details:
        format_rows = []
        for section in sections:
            details = structure_details.get("sections", {}).get(section, {})
            format_rows.append({
                "Section": section.title(),
                "Found": "Yes" if details.get("found") else "No",
                "Paragraph": str(details.get("index")) if details.get("index") is not None else "-",
                "Format Evidence": format_structure_evidence(
                    details.get("format_status", "not_checked")
                ),
            })
        st.dataframe(format_rows, use_container_width=True, hide_index=True)
        if structure_details.get("order_correct", True):
            st.caption("Section order check: passed")
        else:
            st.warning("Section order check: required sections may be out of order.")

    # Issues by category
    st.subheader("Issues by Category")

    tabs = st.tabs([
        "All Issues", "Margins", "Layout", "Journal Header", "Title", "Author Info", "Body Text",
        "Headings", "Structure", "References"
    ])

    with tabs[0]:
        # All issues view
        if result.total_issues == 0:
            st.success("No formatting issues found.")
        else:
            for category, issues in result.issues_by_category.items():
                if issues:
                    with st.expander(f"**{category.replace('_', ' ').title()}** ({len(issues)} issues)", expanded=False):
                        for issue in issues:
                            record = build_issue_display_record(issue, category)
                            severity_class = "issue-error" if issue.severity == "error" else "issue-warning"
                            action_color = "#0f766e" if record["auto_fix_supported"] else "#856404"
                            st.markdown(f"""
                            <div class="issue-card {severity_class}">
                                <strong>{html.escape(record["severity"].title())}: {html.escape(record["location"])}</strong><br>
                                {html.escape(record["description"])}<br>
                                <span style="color: #dc3545;">Current: {html.escape(record["current_value"])}</span> |
                                <span style="color: #28a745;">Expected: {html.escape(record["expected_value"])}</span><br>
                                <span style="color: {action_color};"><strong>Review action:</strong> {html.escape(record["action_label"])}</span><br>
                                <span style="color: #6c757d; font-size: 0.92rem;">{html.escape(record["action_detail"])}</span>
                            </div>
                            """, unsafe_allow_html=True)

    # Category-specific tabs
    category_map = {
        1: "margins",
        2: "layout",
        3: "journal_header",
        4: "title",
        5: "author_info",
        6: "body_text",
        7: "headings",
        8: "structure",
        9: "references"
    }

    for tab_idx, category in category_map.items():
        with tabs[tab_idx]:
            issues = result.issues_by_category.get(category, [])
            if not issues:
                st.success(f"No {category.replace('_', ' ')} issues found")
            else:
                ai_ready = bool(
                    st.session_state.ai_explanations_enabled
                    and st.session_state.llm
                    and st.session_state.llm.is_available()
                )
                explanation_label = get_issue_explanation_button_label(
                    st.session_state.ai_explanations_enabled,
                    ai_ready,
                )
                for issue_idx, issue in enumerate(issues):
                    record = build_issue_display_record(issue, category)
                    severity_label = "Error" if record["severity"] == "error" else "Warning"
                    st.markdown(f"**{severity_label}: {record['location']}**")
                    st.write(f"Problem: {record['description']}")
                    st.write(
                        f"Current: {record['current_value']} | "
                        f"Expected: {record['expected_value']}"
                    )
                    st.caption(
                        f"Review action: {record['action_label']} - "
                        f"{record['action_detail']}"
                    )

                    if st.button(
                        explanation_label,
                        key=f"explain_{category}_{tab_idx}_{issue_idx}_{issue.paragraph_index}",
                    ):
                        issue_payload = {
                            "category": record["category"],
                            "location": record["location"],
                            "description": record["description"],
                            "current_value": record["current_value"],
                            "expected_value": record["expected_value"],
                            "severity": record["severity"],
                            "text_preview": record["text_preview"],
                            "auto_fix_supported": record["auto_fix_supported"],
                            "property_name": record["property_name"],
                            "review_reason": record["review_reason"],
                        }
                        used_ai = (
                            st.session_state.ai_explanations_enabled
                            and ensure_llm_connection()
                        )
                        if used_ai:
                            explanation = st.session_state.llm.explain_error(issue_payload)
                            st.caption("Source: AI explanation based on rule-detected issue metadata")
                        else:
                            explanation = fallback_explain_issue(issue_payload)
                            st.caption("Source: Rule-based explanation")
                        st.info(explanation)

    st.subheader("Review Guidance")
    st.caption(
        "AI organizes and explains rule-detected results. It does not detect "
        "issues or approve manuscripts. When enabled, only grouped and redacted "
        "issue metadata is sent to the configured API."
    )
    guidance_ai_ready = bool(
        st.session_state.ai_explanations_enabled
        and st.session_state.llm
        and st.session_state.llm.is_available()
    )
    guidance_level, guidance_message = get_review_guidance_mode_notice(
        ai_enabled=st.session_state.ai_explanations_enabled,
        connected=guidance_ai_ready,
        key_configured=bool(get_server_nvidia_api_key()),
    )
    getattr(st, guidance_level)(guidance_message)
    if st.button("Generate Review Guidance", key="generate_review_guidance"):
        payload = build_pre_fix_guidance_payload(
            result,
            st.session_state.template_rules or get_default_template_rules(),
        )
        llm = None
        if (
            st.session_state.ai_explanations_enabled
            and ensure_llm_connection()
        ):
            llm = st.session_state.llm
        with st.spinner("Preparing review guidance..."):
            guidance, cache_key, source = resolve_pre_fix_guidance(
                payload,
                llm,
                st.session_state.review_guidance_cache,
            )
        st.session_state.pre_fix_guidance = guidance
        st.session_state.pre_fix_guidance_key = cache_key
        st.session_state.pre_fix_guidance_source = source

    if st.session_state.pre_fix_guidance:
        st.caption(f"Source: {st.session_state.pre_fix_guidance_source}")
        with st.container(border=True):
            st.markdown(
                st.session_state.pre_fix_guidance.replace("\n", "  \n")
            )


def build_report_generator(
    rules,
    changes,
    check_result,
    post_fix_validation=None,
    post_fix_result=None,
    generator_cls=ReportGenerator,
):
    """Create a report generator while tolerating older constructor signatures."""
    optional_kwargs = {
        "post_fix_validation": post_fix_validation,
        "post_fix_result": post_fix_result,
    }
    signature = inspect.signature(generator_cls.__init__)
    supports_kwargs = any(
        parameter.kind == inspect.Parameter.VAR_KEYWORD
        for parameter in signature.parameters.values()
    )
    supported_kwargs = (
        optional_kwargs
        if supports_kwargs
        else {
            key: value
            for key, value in optional_kwargs.items()
            if key in signature.parameters
        }
    )
    return generator_cls(rules, changes, check_result, **supported_kwargs)


def handle_auto_fix():
    """Handle auto-fix process"""
    if not st.session_state.manuscript_bytes:
        st.warning("Please upload a manuscript first")
        return

    st.session_state.post_fix_guidance = ""
    st.session_state.post_fix_guidance_key = ""
    st.session_state.post_fix_guidance_source = ""

    try:
        with st.spinner("Applying formatting fixes..."):
            rules = st.session_state.template_rules or get_default_template_rules()
            classifications = st.session_state.classifications

            # Create auto-fixer
            issues_by_category = {}
            if st.session_state.check_result:
                issues_by_category = st.session_state.check_result.issues_by_category
            fixer = AutoFixer(rules, classifications, issues_by_category=issues_by_category)
            fixer.load_manuscript(BytesIO(st.session_state.manuscript_bytes))

            # Apply fixes
            fixed_doc, changes = fixer.fix_all()

            # Replicate the template's own first-page header/footer
            # structure (journal title, volume/eISSN line and rule
            # lines; the MMU logo and DOI/copyright text in its
            # intentionally-overlapping position) - a rules dict alone
            # can't express these, since they're drawn shapes anchored
            # with specific offsets, not paragraph-level formatting.
            template_source = st.session_state.template_bytes or get_default_template_bytes()
            if template_source:
                masthead_warnings = []
                try:
                    apply_template_first_page_header(fixed_doc, BytesIO(template_source), warnings=masthead_warnings)
                    apply_template_first_page_footer(fixed_doc, BytesIO(template_source), warnings=masthead_warnings)
                except Exception as e:
                    masthead_warnings.append(f"Could not replicate template header/footer: {str(e)}")
                for message in masthead_warnings:
                    st.info(message)

            st.session_state.changes = changes
            st.session_state.fixed_doc_bytes = fixer.get_fixed_document_bytes()
            st.session_state.highlighted_doc_bytes = fixer.get_highlighted_document_bytes()
            st.session_state.output_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.session_state.post_fix_result = validate_fixed_document(
                rules,
                st.session_state.fixed_doc_bytes,
            )
            st.session_state.post_fix_validation = validate_post_fix_result(
                st.session_state.check_result,
                st.session_state.post_fix_result,
            )

            # Generate report
            report_gen = build_report_generator(
                rules,
                changes,
                st.session_state.check_result,
                post_fix_validation=st.session_state.post_fix_validation,
                post_fix_result=st.session_state.post_fix_result,
            )
            report_gen.generate_comparison_report()
            st.session_state.report_bytes = report_gen.get_report_bytes()

            validation = st.session_state.post_fix_validation
            if validation and validation.is_safe:
                st.success(f"{len(changes)} formatting fixes applied. {validation.message}")
            elif validation:
                st.warning(f"{len(changes)} formatting fixes applied. {validation.message}")
            else:
                st.success(f"{len(changes)} formatting fixes applied.")

            return changes

    except Exception as e:
        display_exception("Error during auto-fix", e)

    return None


def format_change_display_value(value):
    """Return user-facing text for a recorded format-change value."""
    if value == "(inherited)":
        return "Inherited from Word style"
    return value


def display_comparison_view(changes):
    """Display structured change records without parsing free-text strings."""
    st.header("Format Changes Applied")

    if not changes:
        st.info("No changes were made to the document")
        return

    changes_by_type = {}
    for change in changes:
        changes_by_type[change.change_type] = changes_by_type.get(change.change_type, 0) + 1
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Properties Changed", len(changes))
    with col2:
        most_common = max(changes_by_type, key=changes_by_type.get)
        st.metric("Most Common", most_common.replace("_", " ").title())
    with col3:
        st.metric("Change Types", len(changes_by_type))

    with st.expander("Changes Summary by Type", expanded=False):
        for change_type, count in sorted(changes_by_type.items(), key=lambda item: -item[1]):
            st.write(f"- **{change_type.replace('_', ' ').title()}**: {count} properties")

    table_data = []
    for index, change in enumerate(changes, 1):
        text_preview = change.text_preview[:60] + "..." if len(change.text_preview) > 60 else change.text_preview
        table_data.append({
            "#": index,
            "Paragraph": str(change.paragraph_index + 1) if change.paragraph_index >= 0 else "Document",
            "Type": change.change_type,
            "Property": change.property_name,
            "Current Value": format_change_display_value(change.current_value or change.before),
            "Target Value": format_change_display_value(change.target_value or change.after),
            "Text": text_preview,
            "Evidence": change.evidence,
        })

    import pandas as pd
    df = pd.DataFrame(table_data)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.success(f"Recorded {len(changes)} formatting property changes")


def display_post_fix_validation():
    """Display checker results after auto-fix."""
    validation = st.session_state.get("post_fix_validation")
    if not validation:
        return

    st.subheader("Post-Fix Validation")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Issues Before", validation.before_issues)
    with col2:
        st.metric("Issues After", validation.after_issues, delta=validation.issue_delta)
    with col3:
        st.metric("Compliance After", f"{validation.after_score}%", delta=validation.score_delta)

    if validation.is_safe:
        st.success(validation.message)
    else:
        st.warning(validation.message)
        if validation.new_or_increased_categories:
            category_text = ", ".join(
                f"{category.replace('_', ' ').title()}: +{count}"
                for category, count in sorted(validation.new_or_increased_categories.items())
            )
            st.caption(f"Increased categories: {category_text}")

    remaining_result = st.session_state.get("post_fix_result")
    remaining_rows = summarize_remaining_issues(remaining_result) if remaining_result else []
    if remaining_rows:
        with st.expander("Remaining Issues After Auto-Fix", expanded=not validation.is_safe):
            st.dataframe(remaining_rows, use_container_width=True, hide_index=True)
            st.caption("These issues still need manual review or a more specific journal rule.")
    else:
        st.success("No remaining issues were detected after auto-fix.")

    st.caption(
        "The explanation uses recorded changes and the post-fix rule check. "
        "It does not inspect or approve manuscript content."
    )
    if st.button("Explain Fix Results", key="explain_fix_results"):
        payload = build_post_fix_guidance_payload(
            st.session_state.get("check_result"),
            remaining_result,
            st.session_state.get("changes", []),
            validation,
            st.session_state.template_rules or get_default_template_rules(),
        )
        llm = None
        if (
            st.session_state.ai_explanations_enabled
            and ensure_llm_connection()
        ):
            llm = st.session_state.llm
        with st.spinner("Preparing post-fix guidance..."):
            guidance, cache_key, source = resolve_post_fix_guidance(
                payload,
                llm,
                st.session_state.review_guidance_cache,
            )
        st.session_state.post_fix_guidance = guidance
        st.session_state.post_fix_guidance_key = cache_key
        st.session_state.post_fix_guidance_source = source

    if st.session_state.post_fix_guidance:
        st.caption(st.session_state.post_fix_guidance_source)
        with st.container(border=True):
            st.markdown(
                st.session_state.post_fix_guidance.replace("\n", "  \n")
            )


def display_download_section():
    """Display download buttons for output files"""
    st.header("Download Results")
    st.caption(
        "Corrected Manuscript is the fixed file. Marked Original is not fixed; it is an audit copy of the submitted "
        "manuscript with changed locations highlighted in yellow. Yellow means the location was changed or reviewed, "
        "not necessarily that the issue remains. Fix Summary Report lists applied changes and remaining manual-review items."
    )

    output_timestamp = st.session_state.output_timestamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    col1, col2, col3 = st.columns(3)

    with col1:
        if st.session_state.fixed_doc_bytes:
            base_filename = st.session_state.manuscript_filename
            st.download_button(
                label="Download Corrected Manuscript",
                data=st.session_state.fixed_doc_bytes,
                file_name=f"corrected_{base_filename}_{output_timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                help="This is the editable manuscript after supported auto-fixes were applied."
            )
        else:
            st.button("Download Corrected Manuscript", disabled=True, use_container_width=True)

    with col2:
        if st.session_state.highlighted_doc_bytes:
            base_filename = st.session_state.manuscript_filename
            st.download_button(
                label="Download Marked Original for Review",
                data=st.session_state.highlighted_doc_bytes,
                file_name=f"highlighted_{base_filename}_{output_timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                help="This copy keeps the original text/layout and marks changed or review-needed locations in yellow."
            )
        else:
            st.button("Download Marked Original for Review", disabled=True, use_container_width=True)

    with col3:
        if st.session_state.report_bytes:
            base_filename = st.session_state.manuscript_filename
            st.download_button(
                label="Download Fix Summary Report",
                data=st.session_state.report_bytes,
                file_name=f"{base_filename}_report_{output_timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                help="This report summarizes detected issues, applied fixes, target rules, and remaining checks."
            )
        else:
            st.button("Download Fix Summary Report", disabled=True, use_container_width=True)


def main():
    """Main application function"""
    init_session_state()
    display_header()
    display_sidebar()

    # Main content area
    st.header("Upload Documents")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Template")
        st.write("The official JCLC template is already built in - nothing to upload.")
        if not st.session_state.template_uploaded:
            st.session_state.template_rules = get_default_template_rules()
            st.session_state.template_bytes = get_default_template_bytes()
            st.session_state.template_uploaded = True

    with col2:
        st.subheader("2. Upload Manuscript")
        st.write("Upload your manuscript (.docx) to check formatting")

        manuscript_file = st.file_uploader(
            "Choose manuscript file",
            type=["docx"],
            key="manuscript_uploader",
            help="Upload your manuscript document to check (DOCX)"
        )

        if manuscript_file and st.button("Check Format", type="primary"):
            result = handle_manuscript_check(manuscript_file)
            if result:
                st.rerun()

    # Display results if check has been performed
    if st.session_state.check_result:
        st.divider()
        display_check_results(st.session_state.check_result)

        # Auto-fix section
        st.divider()
        st.header("Auto-Fix Formatting")

        col1, col2 = st.columns([3, 1])
        auto_fix_preview = build_auto_fix_preview_record(
            st.session_state.check_result.issues_by_category
        )

        with col1:
            st.write("""
            Auto-Fix only changes deterministic formatting properties that were detected by rules.
            It does not rewrite content, move figures/tables, invent missing sections, or approve the manuscript.
            """)
            display_auto_fix_preview(auto_fix_preview)

        with col2:
            if st.button(
                "Auto-Fix Supported Issues",
                type="primary",
                use_container_width=True,
                disabled=not auto_fix_preview["can_run_auto_fix"],
            ):
                changes = handle_auto_fix()
                if changes:
                    st.rerun()
            if not auto_fix_preview["can_run_auto_fix"]:
                st.caption("No deterministic auto-fix candidates were detected.")

        # Display comparison view if fixes have been applied
        if st.session_state.changes:
            st.divider()
            display_comparison_view(st.session_state.changes)
            display_post_fix_validation()

            # Download section
            st.divider()
            display_download_section()

    # Footer
    st.divider()
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>Automated Manuscript Template Compliance Checker | Built with Streamlit</p>
        <p>Final Year Project 2025</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()