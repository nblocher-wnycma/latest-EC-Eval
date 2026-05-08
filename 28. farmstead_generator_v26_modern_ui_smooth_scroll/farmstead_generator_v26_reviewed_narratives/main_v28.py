"""Farmstead Evaluation Generator v28 runner.

This version keeps the working v27 code intact and adds a stronger final
proofreading/editing layer before the Word document is created.

Important:
- Uses LanguageTool when the optional package language-tool-python is installed.
- Falls back to the built-in deterministic editor if LanguageTool is unavailable.
- The final editor runs inside generate_docx before the report is saved, so the
  generated Word document is edited before the user opens it.
"""

import os
import re
import runpy

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V27 = os.path.join(BASE_DIR, "main_v27.py")

base = runpy.run_path(BASE_V27, run_name="farmstead_generator_v27")
App = base["App"]
SECTIONS = base["SECTIONS"]
QUICK_BUILD_SECTIONS = base["QUICK_BUILD_SECTIONS"]

# Keep v27 methods so v28 can build on them rather than replace the working code.
_V27_PROOFREAD_TEXT = App.proofread_text
_V27_QUALITY_CHECK = App.quality_check_narrative
_V27_PROOFREAD_ALL = App.proofread_all_report_text

_LANGUAGE_TOOL = None
_LANGUAGE_TOOL_TRIED = False


def get_language_tool():
    """Load LanguageTool if available.

    language-tool-python provides a real grammar/spell/style checker. If it is not
    installed or cannot start on the computer, the program continues using the
    built-in rule-based editor.
    """
    global _LANGUAGE_TOOL, _LANGUAGE_TOOL_TRIED
    if _LANGUAGE_TOOL_TRIED:
        return _LANGUAGE_TOOL
    _LANGUAGE_TOOL_TRIED = True
    try:
        import language_tool_python  # type: ignore
        _LANGUAGE_TOOL = language_tool_python.LanguageTool("en-US")
    except Exception:
        _LANGUAGE_TOOL = None
    return _LANGUAGE_TOOL


def protect_ag_terms(text):
    """Protect common ag/NRCS terms from over-editing."""
    replacements = {
        "O&M": "OPERATION_AND_MAINTENANCE_TOKEN",
        "NRCS": "NRCS_TOKEN",
        "CNMP": "CNMP_TOKEN",
        "BMP": "BMP_TOKEN",
        "BMPs": "BMPS_TOKEN",
        "WNYCMA": "WNYCMA_TOKEN",
        "AWM": "AWM_TOKEN",
        "HUA": "HUA_TOKEN",
    }
    protected = text
    for key, value in replacements.items():
        protected = protected.replace(key, value)
    return protected, replacements


def restore_ag_terms(text, replacements):
    restored = text
    for key, value in replacements.items():
        restored = restored.replace(value, key)
    return restored


def rule_based_final_edit(text):
    """Final deterministic edit pass for grammar, flow, and redundancy.

    This is intentionally conservative. It fixes recurring problems from the
    generated EC narratives without inventing new farm facts.
    """
    text = str(text or "").strip()
    if not text:
        return ""

    # Keep bullet lists as lists.
    if "\n" in text or "•" in text:
        lines = []
        for raw in text.splitlines():
            line = raw.strip()
            if not line:
                lines.append("")
                continue
            bullet = ""
            if line.startswith("•"):
                bullet = "• "
                line = line[1:].strip()
            line = rule_based_final_edit(line)
            if bullet:
                line = line.rstrip(".")
            lines.append(bullet + line if line else "")
        return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()

    # Common phrase cleanup.
    replacements = {
        "ag wastes": "agricultural wastes",
        "Ag wastes": "agricultural wastes",
        "look for": "evaluate",
        "regarding any movement of": "and determine whether",
        "inbetween": "between",
        "clean rain water": "clean rainwater",
        "storm water": "stormwater",
        "manure-laden run off": "manure-laden runoff",
        "run off": "runoff",
        "farmstead's landbase": "farmstead's land base",
        "landbase": "land base",
        "BMP's": "BMPs",
        "bmp's": "BMPs",
        "OandM": "O&M",
        "oandm": "O&M",
        "w/": "with",
        "Sq. Ft.": "square feet",
        "Sq. Ft": "square feet",
        "Sq Ft": "square feet",
        "sq. ft.": "square feet",
        "sq ft": "square feet",
        "Lf.": "linear feet",
        " LF": " linear feet",
        "lf": "linear feet",
    }
    for old, new in replacements.items():
        text = re.sub(r"\b" + re.escape(old) + r"\b", new, text, flags=re.IGNORECASE)

    # Remove control/placeholder phrases that should not appear in narrative prose.
    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\[item #\]|\[units\]", "", text, flags=re.I)

    # Fix awkward generated transitions.
    text = re.sub(r"\s*;\s*", ". ", text)
    text = re.sub(r"\.\s+And\s+", ". ", text)
    text = re.sub(r"\.\s+But\s+", ". However, ", text)
    text = re.sub(r"\bThis creates a resource concern because this creates a resource concern because\b", "This creates a resource concern because", text, flags=re.I)
    text = re.sub(r"\bThe planned work is intended to The planned work is intended to\b", "The planned work is intended to", text, flags=re.I)
    text = re.sub(r"\bcontinued continued\b", "continued", text, flags=re.I)

    # Remove duplicate location/size wording that sometimes appears from the BMP table and size box.
    text = re.sub(r"Approximate planned dimensions or units include\s*\.\s*", "", text, flags=re.I)
    text = re.sub(r"Approximate planned dimensions or units include ([^.]+)\.\s+Based on the planned BMP table,", r"Based on the planned BMP table,", text, flags=re.I)

    # Spacing and punctuation.
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = text.replace("County County", "County").replace("Basin Basin", "Basin").replace("Watershed Watershed", "Watershed")
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{4,})\s+\1\b", r"\1", text, flags=re.I)

    # Sentence de-duplication.
    sentences = re.findall(r"[^.!?]+[.!?]|[^.!?]+$", text.strip())
    output = []
    seen_keys = set()
    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue
        key = re.sub(r"[^a-z0-9\s]", " ", sentence.lower())
        key = re.sub(r"\b(the|a|an|and|or|to|of|for|in|on|at|as|by|with|from|this|that|these|those|it|is|are|was|were|be|been|being|should|continue|continued)\b", " ", key)
        key = re.sub(r"\s+", " ", key).strip()
        if key in seen_keys:
            continue
        if "no significant resource concern" in sentence.lower() and any("no significant resource concern" in s.lower() for s in output):
            continue
        if "continued operation and maintenance" in sentence.lower() and any("continued operation and maintenance" in s.lower() for s in output):
            continue
        output.append(sentence)
        seen_keys.add(key)

    text = " ".join(output).strip()

    # Capitalize sentence starts.
    def cap(match):
        return match.group(1) + match.group(2).upper()
    text = re.sub(r"(^|[.!?]\s+)([a-z])", cap, text)

    if text and text[-1] not in ".!?:":
        text += "."
    return text.strip()


def language_tool_edit(text):
    """Run LanguageTool correction when available."""
    tool = get_language_tool()
    if tool is None:
        return text
    text = str(text or "")
    if not text.strip():
        return ""

    # Do not send bullet lists as one paragraph because grammar tools often remove
    # list formatting. Edit bullet lines individually.
    if "\n" in text or "•" in text:
        lines = []
        for raw in text.splitlines():
            line = raw.strip()
            if not line:
                lines.append("")
                continue
            bullet = ""
            if line.startswith("•"):
                bullet = "• "
                line = line[1:].strip()
            protected, repl = protect_ag_terms(line)
            try:
                corrected = tool.correct(protected)
            except Exception:
                corrected = protected
            corrected = restore_ag_terms(corrected, repl)
            corrected = rule_based_final_edit(corrected)
            if bullet:
                corrected = corrected.rstrip(".")
            lines.append(bullet + corrected if corrected else "")
        return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()

    protected, repl = protect_ag_terms(text)
    try:
        corrected = tool.correct(protected)
    except Exception:
        corrected = protected
    corrected = restore_ag_terms(corrected, repl)
    return rule_based_final_edit(corrected)


def final_editor(text):
    """Complete final edit: v27 cleanup -> rule-based edit -> LanguageTool -> final rule edit."""
    text = _V27_QUALITY_CHECK(None, text) if callable(_V27_QUALITY_CHECK) else str(text or "")
    text = rule_based_final_edit(text)
    text = language_tool_edit(text)
    text = rule_based_final_edit(text)
    return text


def v28_proofread_text(self, text):
    return final_editor(text)


def v28_quality_check_narrative(self, text):
    return final_editor(text)


def v28_clean_paragraph(self, text):
    return final_editor(text)


def v28_proofread_all_report_text(self):
    """Edit all report narratives immediately before Word export."""
    # First allow v27 to do its cleanup.
    _V27_PROOFREAD_ALL(self)

    front_fields = [
        "Introduction Narrative",
        "Overview Narrative",
        "Animal Housing and Feed Management Narrative",
        "Manure Management and Application Narrative",
    ]
    for field in front_fields:
        value = self.get_widget_value("Basic Info", field)
        if value:
            self.set_widget_value("Basic Info", field, final_editor(value))

    for sec in SECTIONS[1:]:
        for field in ["Existing Conditions narrative", "Planned Actions narrative"]:
            value = self.get_widget_value(sec, field)
            if value:
                self.set_widget_value(sec, field, final_editor(value))


def v28_add_paragraphs(self, container, text):
    """Write final-edited paragraphs while preserving bullet lists."""
    text = (text or "").strip()
    if not text:
        container.add_paragraph("[No narrative entered.]")
        return
    for para in text.split("\n"):
        para = para.strip()
        if not para:
            continue
        if para.startswith("•"):
            p = container.add_paragraph(style=None)
            p.paragraph_format.left_indent = base["base"]["Inches"](0.25) if "base" in base else runpy.run_path(BASE_V27)["base"]["Inches"](0.25)
            p.add_run(final_editor(para[1:].strip()).rstrip("."))
        else:
            container.add_paragraph(final_editor(para))


# Patch v27 App methods.
App.proofread_text = v28_proofread_text
App.quality_check_narrative = v28_quality_check_narrative
App.clean_paragraph = v28_clean_paragraph
App.proofread_all_report_text = v28_proofread_all_report_text
# Keep v27 add_paragraphs unless LanguageTool is installed; generate_docx already calls clean_paragraph on paragraphs.


if __name__ == "__main__":
    App().mainloop()
