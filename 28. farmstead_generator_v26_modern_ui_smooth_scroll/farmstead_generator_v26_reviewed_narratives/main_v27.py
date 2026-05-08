"""Farmstead Evaluation Generator v27 runner.

This file keeps the working v26 code intact and applies targeted runtime patches for:
- stronger grammar/redundancy cleanup before Word export
- preserving animal inventory as a list on the front page
- using planned BMP table dimensions more naturally in planned-action narratives

Run this file instead of main.py, or use Run_Farmstead_Eval_v27.bat.
"""

import os
import re
import runpy

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_MAIN = os.path.join(BASE_DIR, "main.py")

base = runpy.run_path(BASE_MAIN, run_name="farmstead_generator_base")
App = base["App"]
SECTIONS = base["SECTIONS"]
QUICK_BUILD_SECTIONS = base["QUICK_BUILD_SECTIONS"]

# Keep original methods so the v27 patch can wrap them without rewriting the
# working v26 program.
_ORIGINAL_MAKE_PLANNED_NARRATIVE = App.make_planned_narrative


def _clean_common_text(text: str) -> str:
    """Rules-based cleanup for field-note style input."""
    text = str(text or "").strip()
    if not text:
        return ""

    replacements = {
        "&": "and",
        "witner": "winter",
        "inclimate": "inclement",
        "suplimentaly": "supplementally",
        "supplementaly": "supplementally",
        "suplimental": "supplemental",
        "Mnnure": "Manure",
        "mnnure": "manure",
        " inbetween ": " between ",
        "inbetween": "between",
        " ino ": " into ",
        "tnaks": "tanks",
        "mortaliy": "mortality",
        "mortal. facility": "mortality facility",
        "approprite": "appropriate",
        "acces": "access",
        "draned": "drained",
        "cattl": "cattle",
        "landbase": "land base",
        "drive ways": "driveways",
        "BMP's": "BMPs",
        "bmp's": "BMPs",
        "clean rain water": "clean rainwater",
        "storm water": "stormwater",
        "OandM": "O&M",
        "oandm": "O&M",
        "w/": "with",
        "UO": "underground outlet",
    }
    for bad, good in replacements.items():
        if bad.startswith(" ") or bad.endswith(" "):
            text = text.replace(bad, good)
        else:
            text = re.sub(r"\b" + re.escape(bad) + r"\b", good, text, flags=re.IGNORECASE)

    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\[item #\]|\[units\]", "", text, flags=re.I)

    text = text.replace("County County", "County")
    text = text.replace("Basin Basin", "Basin")
    text = text.replace("Watershed Watershed", "Watershed")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = text.replace(". .", ".")
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]+)\s+\1\b", r"\1", text, flags=re.I)

    text = re.sub(r"\.\s+And\s+", ". ", text)
    text = re.sub(r"\.\s+But\s+", ". However, ", text)
    text = re.sub(r"\bin barn\b", "in the barn", text, flags=re.I)
    text = re.sub(r"\bis all exported\b", "is exported", text, flags=re.I)

    def cap(match):
        return match.group(1) + match.group(2).upper()

    text = re.sub(r"(^|[.!?]\s+)([a-z])", cap, text.strip())
    return text.strip()


def _split_sentences(text: str):
    text = str(text or "").strip()
    if not text:
        return []
    sentences = re.findall(r"[^.!?]+[.!?]|[^.!?]+$", text)
    return [s.strip() for s in sentences if s.strip()]


def _sentence_key(sentence: str) -> str:
    key = sentence.lower()
    key = re.sub(r"\([^)]*\)", "", key)
    key = re.sub(r"[^a-z0-9\s]", " ", key)
    key = re.sub(r"\b(the|a|an|and|or|to|of|for|in|on|at|as|by|with|from|this|that|these|those|it|is|are|was|were|be|been|being|should|continue|continued)\b", " ", key)
    key = re.sub(r"\s+", " ", key).strip()
    return key


def _too_similar(new_key: str, previous_keys):
    if not new_key:
        return True
    if new_key in previous_keys:
        return True
    new_words = set(new_key.split())
    if len(new_words) < 5:
        return False
    for old in previous_keys:
        old_words = set(old.split())
        if not old_words:
            continue
        overlap = len(new_words & old_words) / max(len(new_words), 1)
        if overlap >= 0.82 and len(new_words & old_words) >= 7:
            return True
    return False


def _dedupe_sentences(text: str) -> str:
    sentences = _split_sentences(text)
    output = []
    keys = []
    for sentence in sentences:
        cleaned = _clean_common_text(sentence).strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if "no significant resource concern" in lowered:
            if any("no significant resource concern" in s.lower() for s in output):
                continue
        if "continued operation and maintenance" in lowered:
            if any("continued operation and maintenance" in s.lower() for s in output):
                continue
        key = _sentence_key(cleaned)
        if _too_similar(key, keys):
            continue
        output.append(cleaned)
        keys.append(key)

    final = " ".join(output)
    final = re.sub(r"\s+", " ", final).strip()
    if final and final[-1] not in ".!?:":
        final += "."
    return final


def _parse_practice_table(raw: str):
    """Parse Standard|Item #|Description|Units rows from the Quick Build table."""
    rows = []
    for line in str(raw or "").splitlines():
        line = line.strip()
        if not line or line.lower().startswith("standard"):
            continue
        if "|" not in line:
            continue
        parts = [p.strip() for p in line.split("|")]
        parts += [""] * (4 - len(parts))
        practice, item, desc, units = parts[:4]
        if not practice:
            continue
        if not units or units.lower() in {"[units]", "units", "n/a", "na", "none"}:
            continue
        rows.append({"practice": practice, "item": item, "desc": desc or practice, "units": units})
    return rows


def _dimension_connector(units: str) -> str:
    lower = units.lower().strip()
    if " x " in lower or "'x" in lower or "’x" in lower or " by " in lower:
        return "measuring"
    if any(term in lower for term in ["sq", "square", "ac", "acre"]):
        return "covering"
    if any(term in lower for term in ["lf", "linear", " ft", "feet", "foot"]):
        return "totaling approximately"
    if re.fullmatch(r"\d+(\.\d+)?", lower):
        return "with a planned quantity of"
    return "with planned units of"


def _clean_dimension_text(text: str) -> str:
    text = _clean_common_text(text).rstrip(".")
    text = text.replace("Sq. Ft", "square feet")
    text = text.replace("Sq Ft", "square feet")
    text = text.replace("sq. ft", "square feet")
    text = text.replace("sq ft", "square feet")
    text = re.sub(r"\bLf\.?\b", "linear feet", text, flags=re.I)
    return text


def _practice_dimension_phrase(row):
    desc = _clean_dimension_text(row.get("desc") or row.get("practice") or "planned practice")
    practice = _clean_dimension_text(row.get("practice") or "")
    units = _clean_dimension_text(row.get("units") or "")
    if not units:
        return ""

    # If the description is just the practice name, do not repeat it.
    if practice and desc.lower() == practice.lower():
        label = practice
    else:
        label = desc
        if practice and practice.lower() not in desc.lower():
            label = f"{desc} ({practice})"

    connector = _dimension_connector(units)
    return f"{label} {connector} {units}"


def _list_phrase(items):
    items = [i for i in items if i]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def planned_bmp_sizing_sentence(self, section):
    """Create a natural sizing sentence from the planned BMP table."""
    selected = self.get_widget_value("Quick Build Wizard", f"{section} - Selected Practices")
    if not selected:
        selected = self.get_widget_value(section, "NRCS Standards and Units Used")
    rows = _parse_practice_table(selected)
    phrases = [_practice_dimension_phrase(row) for row in rows]
    phrases = [p for p in phrases if p]
    if not phrases:
        return ""
    return "Based on the planned BMP table, the proposed work includes " + _list_phrase(phrases) + "."


def _insert_after_first_sentence(text: str, insert_sentence: str) -> str:
    text = str(text or "").strip()
    insert_sentence = str(insert_sentence or "").strip()
    if not text or not insert_sentence:
        return text
    if insert_sentence.lower() in text.lower():
        return text
    sentences = _split_sentences(text)
    if not sentences:
        return text + " " + insert_sentence
    return " ".join([sentences[0], insert_sentence] + sentences[1:])


def improved_make_planned_narrative(self, section, subject, practices, concern_yes, loc, size, action_language, om_language):
    """Use planned BMP table dimensions in a clearer way.

    The base narrative used one generic sentence like "Approximate planned dimensions
    or units include...". This patch reads each BMP table row and writes a more useful
    sentence tying each practice/description to its own units.
    """
    sizing_sentence = planned_bmp_sizing_sentence(self, section)

    # If we have a real sizing sentence from the BMP table, do not also pass the
    # generic size field into the base method; that avoids repeated sizing language.
    size_for_base = "" if sizing_sentence else size
    narrative = _ORIGINAL_MAKE_PLANNED_NARRATIVE(
        self, section, subject, practices, concern_yes, loc, size_for_base, action_language, om_language
    )
    if sizing_sentence:
        narrative = _insert_after_first_sentence(narrative, sizing_sentence)
    return improved_quality_check_narrative(self, narrative)


def improved_proofread_text(self, text):
    """Improved deterministic grammar/flow cleanup."""
    text = str(text or "").strip()
    if not text:
        return ""

    if "\n" in text or "•" in text:
        cleaned_lines = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                cleaned_lines.append("")
                continue
            bullet = ""
            if line.startswith("•"):
                bullet = "• "
                line = line[1:].strip()
            elif re.match(r"^[-*]\s+", line):
                bullet = "• "
                line = re.sub(r"^[-*]\s+", "", line).strip()
            line = _clean_common_text(line)
            if line and bullet:
                line = line.rstrip(".")
            elif line and line[-1] not in ".!?:":
                line += "."
            cleaned_lines.append(bullet + line if line else "")
        out = "\n".join(cleaned_lines)
        out = re.sub(r"\n{3,}", "\n\n", out).strip()
        return out

    text = _clean_common_text(text)
    text = re.sub(r"\s*;\s*", ". ", text)
    text = _clean_common_text(text)
    text = _dedupe_sentences(text)

    def cap(match):
        return match.group(1) + match.group(2).upper()

    text = re.sub(r"(^|[.!?]\s+)([a-z])", cap, text.strip())
    if text and text[-1] not in ".!?:":
        text += "."
    return text


def improved_quality_check_narrative(self, text):
    """Final narrative cleanup before saving/export."""
    text = improved_proofread_text(self, text)
    if "\n" in text:
        return text
    text = _dedupe_sentences(text)
    text = re.sub(
        r"Based on the conditions observed during the evaluation, this area does not appear to be creating a significant resource concern at this time\.\s+Based on the conditions observed during the evaluation,",
        "Based on the conditions observed during the evaluation,",
        text,
        flags=re.I,
    )
    text = re.sub(r"\s+", " ", text).strip()
    return text


def improved_clean_paragraph(self, text):
    return improved_quality_check_narrative(self, text)


def improved_proofread_all_report_text(self):
    """Run grammar/flow cleanup while preserving the animal inventory list."""
    front_fields = [
        "Introduction Narrative",
        "Animal Housing and Feed Management Narrative",
        "Manure Management and Application Narrative",
    ]
    for field in front_fields:
        value = self.get_widget_value("Basic Info", field)
        if value:
            self.set_widget_value("Basic Info", field, improved_quality_check_narrative(self, value))

    overview = self.get_widget_value("Basic Info", "Overview Narrative")
    if overview:
        self.set_widget_value("Basic Info", "Overview Narrative", improved_proofread_text(self, overview))

    for sec in SECTIONS[1:]:
        for field in ["Existing Conditions narrative", "Planned Actions narrative"]:
            value = self.get_widget_value(sec, field)
            if value:
                self.set_widget_value(sec, field, improved_quality_check_narrative(self, value))


def improved_bullet_inventory(self, raw):
    """Keep animal inventory as a true bullet list instead of a run-on paragraph."""
    lines = []
    for line in (raw or "").splitlines():
        line = _clean_common_text(line).strip().strip("•-").strip()
        if not line or line.lower().startswith("example"):
            continue
        line = line.rstrip(".")
        lines.append(f"• {line}")
    return "\n".join(lines)


def improved_add_paragraphs(self, container, text):
    """Write paragraphs while preserving bullet lists."""
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
            p.paragraph_format.left_indent = base["Inches"](0.25)
            p.add_run(improved_proofread_text(self, para[1:].strip()).rstrip("."))
        else:
            container.add_paragraph(improved_clean_paragraph(self, para))


# Apply patches to the existing app class.
App.make_planned_narrative = improved_make_planned_narrative
App.proofread_text = improved_proofread_text
App.quality_check_narrative = improved_quality_check_narrative
App.clean_paragraph = improved_clean_paragraph
App.proofread_all_report_text = improved_proofread_all_report_text
App.bullet_inventory = improved_bullet_inventory
App.add_paragraphs = improved_add_paragraphs


if __name__ == "__main__":
    App().mainloop()
