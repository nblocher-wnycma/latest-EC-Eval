"""Farmstead Existing Conditions Evaluation Generator v34.

Safe test version built from the current working base, main_v30.py.

This file keeps the working Quick Build Wizard, BMP practice search/table,
planned BMP dimension handling, front-page inventory bullets, and Word export
logic intact. It adds a stronger internal proofreading pipeline that edits the
actual outgoing narratives before the Word document is saved.
"""

from __future__ import annotations

import os
import re
import runpy

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V30 = os.path.join(BASE_DIR, "main_v30.py")

base = runpy.run_path(BASE_V30, run_name="farmstead_generator_v30")
App = base["App"]
SECTIONS = base["SECTIONS"]
QUICK_BUILD_SECTIONS = base["QUICK_BUILD_SECTIONS"]

_V30_GENERATE_DOCX = App.generate_docx
_V30_MAKE_PLANNED_NARRATIVE = App.make_planned_narrative

BULLET = "\u2022"


COMMON_TYPO_FIXES = {
    "cattlee": "cattle",
    "cattleee": "cattle",
    "cattleeeee": "cattle",
    "horss": "horses",
    "witner": "winter",
    "inclimate": "inclement",
    "suplimentaly": "supplementally",
    "supplementaly": "supplementally",
    "suplimental": "supplemental",
    "mnnure": "manure",
    "approprite": "appropriate",
    "acces": "access",
    "draned": "drained",
    "tnaks": "tanks",
    "mortaliy": "mortality",
    "inbetween": "between",
    "landbase": "land base",
    "drive ways": "driveways",
}

UNIT_REPLACEMENTS = [
    (r"\bsq\.?\s*ft\.?\b", "square feet"),
    (r"\bsqft\b", "square feet"),
    (r"\bsqf\b", "square feet"),
    (r"\bSF\b", "square feet"),
    (r"\bln\s*ft\.?\b", "linear feet"),
    (r"\blin\s*ft\.?\b", "linear feet"),
    (r"\bl\.?\s*f\.?\b", "linear feet"),
    (r"\bLF\b", "linear feet"),
]

TERM_REPLACEMENTS = [
    (r"\bstorm\s+water\b", "stormwater"),
    (r"\brun\s+off\b", "runoff"),
    (r"\bclean\s+rain\s+water\b", "clean rainwater"),
    (r"\bBMP's\b", "BMPs"),
]

LONG_LIST_REPLACEMENTS = {
    "ditches, wetlands, streams, wells, poorly drained soils, or other sensitive areas": "nearby sensitive areas",
    "ditches, wetlands, streams, wells, poorly drained soils and other sensitive areas": "nearby sensitive areas",
    "manure, bedding, feed waste, livestock traffic areas, and other potential contamination sources": "agricultural waste sources",
    "manure handling and livestock use areas so rainfall and snowmelt remain separated from agricultural waste": "manure handling and livestock use areas",
    "manure, sediment, nutrients, leachate, or other contaminants": "contaminants",
    "odor, dust, ammonia, methane, or nuisance conditions": "air quality concerns",
}

PLACEHOLDER_PATTERNS = [
    r"\[Farm Name\]",
    r"\[address\]",
    r"\[town\]",
    r"\[county\]",
    r"\[watershed\]",
    r"\[item #\]",
    r"\[units\]",
    r"\[No narrative entered\.\]",
]

REDUNDANT_BOILERPLATE_PATTERNS = [
    r"The planned work is intended to improve day-to-day management, keep clean water separated from agricultural waste where applicable, and reduce the potential for manure, sediment, nutrients, leachate, or other contaminants to move from the farmstead\.",
    r"Final layout, grades, outlets, and construction details should be confirmed during design and installation so the practice functions as intended and meets applicable NRCS standards\.",
    r"The concern should be reviewed with the planned action for this section and with the applicable NRCS practice standards listed in the evaluation\.\s*The concern should be reviewed with the planned action for this section and with the applicable NRCS practice standards listed in the evaluation\.",
]

SENTENCE_STOP_WORDS = {
    "the", "a", "an", "and", "or", "to", "of", "for", "in", "on", "at",
    "as", "by", "with", "from", "this", "that", "these", "those", "it",
    "is", "are", "was", "were", "be", "been", "being", "should",
    "continue", "continued", "planned", "work", "intended", "practice",
    "practices", "area", "farm", "farmstead",
}


def normalize_units(text: str) -> str:
    """Normalize common unit shorthand while preserving numeric values."""
    for pattern, replacement in UNIT_REPLACEMENTS:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def normalize_terminology(text: str) -> str:
    """Normalize common CNMP/EC terminology variants."""
    for pattern, replacement in TERM_REPLACEMENTS:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def correct_common_typos(text: str) -> str:
    """Correct common field-note typos and obvious repeated-letter accidents."""
    for typo, correction in COMMON_TYPO_FIXES.items():
        text = re.sub(r"\b" + re.escape(typo) + r"\b", correction, text, flags=re.IGNORECASE)

    def shrink_repeated_letters(match: re.Match[str]) -> str:
        word = match.group(0)
        return re.sub(r"([A-Za-z])\1{2,}", r"\1", word)

    return re.sub(r"\b[A-Za-z]*([A-Za-z])\1{3,}[A-Za-z]*\b", shrink_repeated_letters, text)


def remove_placeholders(text: str) -> str:
    """Remove leftover placeholders and dropdown/control text from narrative prose."""
    for pattern in PLACEHOLDER_PATTERNS:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.IGNORECASE)
    return text


def fix_punctuation(text: str) -> str:
    """Repair spacing, duplicate punctuation, and awkward sentence starts."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = text.replace(". .", ".")
    text = re.sub(r"\.\s+And\s+", ". ", text)
    text = re.sub(r"\.\s+But\s+", ". However, ", text)
    text = re.sub(r"\s+", " ", text).strip()

    def cap_sentence(match: re.Match[str]) -> str:
        return match.group(1) + match.group(2).upper()

    return re.sub(r"(^|[.!?]\s+)([a-z])", cap_sentence, text)


def remove_repeated_words(text: str) -> str:
    """Remove accidental repeated words such as 'the the'."""
    return re.sub(r"\b([A-Za-z]{2,})\s+\1\b", r"\1", text, flags=re.IGNORECASE)


def simplify_long_lists(text: str) -> str:
    """Shorten known long comma-list phrases without removing technical meaning."""
    for long_phrase, replacement in LONG_LIST_REPLACEMENTS.items():
        text = re.sub(re.escape(long_phrase), replacement, text, flags=re.IGNORECASE)
    return text


def _split_sentences(text: str) -> list[str]:
    return [s.strip() for s in re.findall(r"[^.!?]+[.!?]|[^.!?]+$", text.strip()) if s.strip()]


def _sentence_key(sentence: str) -> set[str]:
    key = re.sub(r"\([^)]*\)", "", sentence.lower())
    key = re.sub(r"[^a-z0-9\s]", " ", key)
    return {word for word in key.split() if word not in SENTENCE_STOP_WORDS}


def remove_redundant_sentences(text: str) -> str:
    """Remove exact and near-repeated sentences while preserving unique details."""
    for pattern in REDUNDANT_BOILERPLATE_PATTERNS:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE)

    output: list[str] = []
    seen: list[set[str]] = []
    for sentence in _split_sentences(text):
        cleaned = sentence.strip()
        key = _sentence_key(cleaned)
        if not key:
            continue
        duplicate = False
        for old_key in seen:
            overlap = len(key & old_key)
            if key == old_key or (overlap >= 7 and overlap / max(len(key), 1) > 0.82):
                duplicate = True
                break
        if not duplicate:
            output.append(cleaned)
            seen.append(key)
    return " ".join(output).strip()


def clean_bmp_dimension_language(text: str) -> str:
    """Clean planned BMP dimension wording without dropping dimensions.

    The v30 BMP dimension builder is useful, but descriptions can produce nested
    parentheses such as 'Dry Stack Pad (Waste Storage Facility (313))'. In the
    narrative, the description is enough because the NRCS standard remains in
    the standards table. This keeps output natural while preserving dimensions.
    """
    text = re.sub(r"\bsizing includes\b", "includes", text, flags=re.IGNORECASE)
    text = re.sub(
        r"([A-Za-z0-9][A-Za-z0-9 /&.,'\u2019#-]*?)\s+\(([A-Za-z][A-Za-z /&.,'\u2019#-]+ \(\d{3}\))\)\s+(covering|measuring|totaling|with planned|with a planned)",
        r"\1 \3",
        text,
    )
    text = re.sub(r"\s+\)", ")", text)
    text = re.sub(r"\(\s+", "(", text)
    return text


def proofread_paragraph(text: str) -> str:
    """Run the safe rule-based proofreading pipeline on one narrative paragraph."""
    text = str(text or "").strip()
    if not text:
        return ""

    # Preserve animal inventory and any other bullet-style list formatting.
    if "\n" in text or text.lstrip().startswith((BULLET, "-", "*")):
        cleaned_lines = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            bullet = line.startswith(BULLET) or re.match(r"^[-*]\s+", line)
            line = re.sub(r"^(?:\u2022|[-*])\s*", "", line).strip()
            line = proofread_paragraph(line)
            if bullet:
                line = line.rstrip(".")
                cleaned_lines.append(BULLET + " " + line)
            elif line:
                cleaned_lines.append(line)
        return "\n".join(cleaned_lines)

    text = correct_common_typos(text)
    text = normalize_units(text)
    text = normalize_terminology(text)
    text = remove_placeholders(text)
    text = simplify_long_lists(text)
    text = clean_bmp_dimension_language(text)
    text = remove_repeated_words(text)
    text = fix_punctuation(text)
    text = remove_redundant_sentences(text)
    text = fix_punctuation(text)

    if text and text[-1] not in ".!?:":
        text += "."
    return text


def proofread_report_text(app) -> None:
    """Edit every outgoing narrative field before Word export.

    The base v30 export already generates front-page narratives, rebuilds Quick
    Build sections, and then calls self.proofread_all_report_text() before it
    writes the Word document. v34 replaces that hook with this function so the
    actual text written to the document is cleaned automatically.
    """
    front_fields = [
        "Introduction Narrative",
        "Overview Narrative",
        "Animal Housing and Feed Management Narrative",
        "Manure Management and Application Narrative",
    ]
    for field in front_fields:
        value = app.get_widget_value("Basic Info", field)
        if value:
            app.set_widget_value("Basic Info", field, proofread_paragraph(value))

    for section in SECTIONS[1:]:
        for field in ["Existing Conditions narrative", "Planned Actions narrative"]:
            value = app.get_widget_value(section, field)
            if value:
                app.set_widget_value(section, field, proofread_paragraph(value))


def make_planned_narrative_with_clean_dimensions(self, section, subject, practices, concern_yes, loc, size, action_language, om_language):
    """Keep v30 planned BMP dimension logic, then clean its final sentence flow."""
    narrative = _V30_MAKE_PLANNED_NARRATIVE(self, section, subject, practices, concern_yes, loc, size, action_language, om_language)
    return proofread_paragraph(narrative)


def generate_docx_with_internal_proofreader(self):
    """Preserve v30 export and run v34 cleanup through the existing pre-save hook.

    Export flow:
    1. v30/base export generates front page narratives.
    2. v30/base export rebuilds section narratives from the Quick Build Wizard.
    3. v30 planned action logic generates BMP dimension language.
    4. v34 proofread_report_text() edits the outgoing narrative widgets.
    5. v30/base export writes the cleaned text into the Word document.
    """
    return _V30_GENERATE_DOCX(self)


App.proofread_text = proofread_paragraph
App.quality_check_narrative = proofread_paragraph
App.clean_paragraph = proofread_paragraph
App.proofread_all_report_text = proofread_report_text
App.make_planned_narrative = make_planned_narrative_with_clean_dimensions
App.generate_docx = generate_docx_with_internal_proofreader


if __name__ == "__main__":
    App().mainloop()
