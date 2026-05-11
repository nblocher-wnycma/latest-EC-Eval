"""Farmstead Evaluation Generator v30 runner.

This version avoids the broken local v28 recursion issue by loading v27 directly,
then applying the v29 planned-action and redundancy improvements.

Run with Run_Farmstead_Eval_v30.bat.
"""

import os
import re
import runpy
from tkinter import messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V27 = os.path.join(BASE_DIR, "main_v27.py")

base = runpy.run_path(BASE_V27, run_name="farmstead_generator_v27")
App = base["App"]
SECTIONS = base["SECTIONS"]
QUICK_BUILD_SECTIONS = base["QUICK_BUILD_SECTIONS"]

_PREV_FINAL_EDITOR = App.quality_check_narrative
_PREV_GENERATE_DOCX = App.generate_docx


def _clean(text):
    text = str(text or "").strip()
    if not text:
        return ""
    replacements = {
        "Sq. Ft.": "square feet",
        "Sq. Ft": "square feet",
        "Sq Ft": "square feet",
        "sq. ft.": "square feet",
        "sq. ft": "square feet",
        "sq ft": "square feet",
        "Sq Ft.": "square feet",
        "Lf.": "linear feet",
        "LF": "linear feet",
        "Lf": "linear feet",
        " w/ ": " with ",
        "UO": "underground outlet",
        "BMP's": "BMPs",
        "bmp's": "BMPs",
        "OandM": "O&M",
        "oandm": "O&M",
        "clean rain water": "clean rainwater",
        "storm water": "stormwater",
        "run off": "runoff",
        "manure-laden run off": "manure-laden runoff",
        "inbetween": "between",
        "tnaks": "tanks",
        "acces": "access",
        "draned": "drained",
        "cattl": "cattle",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\[item #\]|\[units\]", "", text, flags=re.I)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{4,})\s+\1\b", r"\1", text, flags=re.I)
    return text.strip(" .")


def _sent(text):
    text = _clean(text)
    if not text:
        return ""
    return text if text[-1] in ".!?:" else text + "."


def _list_phrase(items):
    items = [_clean(i) for i in items if _clean(i)]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def _parse_bmp_rows(raw):
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
        practice = _clean(practice)
        desc = _clean(desc) or practice
        units = _clean(units)
        item = _clean(item)
        if not practice or units.lower() in {"", "[units]", "units", "none", "n/a", "na"}:
            continue
        rows.append({"practice": practice, "item": item, "desc": desc, "units": units})
    return rows


def _practice_label(row):
    practice = row.get("practice", "")
    desc = row.get("desc", "")
    if not desc or desc.lower() == practice.lower():
        return practice
    if practice.lower() not in desc.lower():
        return f"{desc} ({practice})"
    return desc


def _unit_phrase(units):
    units = _clean(units)
    lower = units.lower()
    if not units:
        return ""
    if " x " in lower or "'x" in lower or "’x" in lower or " by " in lower:
        return f"measuring approximately {units}"
    if "square" in lower or "acre" in lower or "acres" in lower:
        return f"covering approximately {units}"
    if "linear feet" in lower or "feet" in lower or "ft" in lower:
        return f"totaling approximately {units}"
    if re.fullmatch(r"\d+(\.\d+)?", lower):
        return f"with a planned quantity of {units}"
    return f"with planned units of {units}"


def _bmp_dimension_sentence(section, rows):
    if not rows:
        return ""
    phrases = []
    for row in rows:
        label = _practice_label(row)
        unit = _unit_phrase(row.get("units", ""))
        if label and unit:
            phrases.append(f"{label} {unit}")
    if not phrases:
        return ""
    prefix_by_section = {
        "Barnyard": "The planned barnyard BMP sizing includes",
        "Clean Stormwater": "The planned clean-water system sizing includes",
        "Livestock Manure Storage": "The planned manure-storage system sizing includes",
        "Pasture": "The planned pasture-management system sizing includes",
        "Animal Mortalities": "The planned mortality-management BMP sizing includes",
        "Fuel / Petroleum": "The planned petroleum-storage BMP sizing includes",
        "Driveways": "The planned access BMP sizing includes",
        "Feed Storage": "The planned feed-storage BMP sizing includes",
        "Air Quality": "The planned air-quality management items include",
    }
    return _sent(prefix_by_section.get(section, "The planned BMP sizing includes") + " " + _list_phrase(phrases))


def _practice_names(rows, fallback):
    if rows:
        return _list_phrase([r["practice"] for r in rows if r.get("practice")])
    return _clean(fallback)


def _section_purpose(section):
    purposes = {
        "Barnyard": "This work is intended to provide a stable livestock use area, reduce hoof pressure on abused areas, improve manure collection, and limit the movement of sediment- and nutrient-laden runoff from the production area.",
        "Clean Stormwater": "This work is intended to keep clean roof runoff and upslope stormwater separated from manure, bedding, feed waste, livestock traffic areas, and other potential contamination sources.",
        "Livestock Manure Storage": "This work is intended to eliminate or reduce uncontrolled temporary manure piles, contain manure and bedding in a defined storage area, and reduce the potential for leachate or manure-laden runoff to leave the farmstead.",
        "Pasture": "This work is intended to improve livestock distribution, reduce prolonged pressure in high-use areas, maintain vegetative cover, and reduce sediment and nutrient transport from pasture areas.",
        "Animal Mortalities": "This work is intended to provide a defined mortality-management method that limits leachate, runoff, scavenger access, and impacts to soil and water resources.",
        "Fuel / Petroleum": "This work is intended to reduce the risk of fuel or petroleum products reaching soil, drainageways, wells, wetlands, streams, or other water resources.",
        "Air Quality": "This work is intended to reduce odor, dust, ammonia, methane, or nuisance conditions associated with manure handling, livestock areas, traffic, or field application.",
        "Driveways": "This work is intended to stabilize farm access, improve trafficability, and reduce the potential for manure, sediment, or nutrients to be transported from traffic areas.",
        "Feed Storage": "This work is intended to control feed runoff, spoiled feed, and silage leachate so nutrients and organic material do not leave the storage area in runoff.",
    }
    return purposes.get(section, "This work is intended to address the observed resource concern and improve long-term operation and maintenance of the farmstead.")


def _design_sentence(section, rows):
    practices = " ".join(r.get("practice", "") + " " + r.get("desc", "") for r in rows).lower()
    if section == "Livestock Manure Storage":
        return "Final storage layout, grades, access, push-off areas, roof runoff handling, and outlet locations should be confirmed during design so the system can be operated safely and maintained over the planned storage period."
    if section == "Clean Stormwater" or "roof runoff" in practices or "underground outlet" in practices:
        return "Outlet locations should be stable and protected so clean water can discharge without causing erosion or entering livestock or manure-handling areas."
    if section == "Barnyard":
        return "Final grades, curbs, roof runoff handling, access points, and manure scraping routes should be laid out so daily operation is practical and runoff remains controlled."
    if section == "Pasture":
        return "Fence, water, and laneway locations should be arranged to support rotation, protect sensitive areas, and maintain forage recovery time."
    return "Final layout and installation details should be checked during design so the planned practice functions as intended and can be maintained by the operator."


def _remove_empty_words(text):
    weak_patterns = [
        r"The planned work is intended to improve day-to-day management, keep clean water separated from agricultural waste where applicable, and reduce the potential for manure, sediment, nutrients, leachate, or other contaminants to move from the farmstead\.",
        r"Final layout, grades, outlets, and construction details should be confirmed during design and installation so the practice functions as intended and meets applicable NRCS standards\.",
    ]
    text = str(text or "")
    for pat in weak_patterns:
        text = re.sub(pat, "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def _dedupe_meaningful(text):
    sentences = re.findall(r"[^.!?]+[.!?]|[^.!?]+$", str(text or "").strip())
    output = []
    seen = []
    stop = {"the", "a", "an", "and", "or", "to", "of", "for", "in", "on", "at", "as", "by", "with", "from", "this", "that", "these", "those", "it", "is", "are", "was", "were", "be", "been", "being", "should", "continue", "continued", "planned", "work", "intended"}
    for s in sentences:
        s = _sent(s)
        words = [w for w in re.sub(r"[^a-z0-9\s]", " ", s.lower()).split() if w not in stop]
        key = set(words)
        if not key:
            continue
        duplicate = False
        for old in seen:
            if len(key & old) / max(len(key), 1) > 0.80 and len(key & old) >= 7:
                duplicate = True
                break
        if not duplicate:
            output.append(s)
            seen.append(key)
    return " ".join(output).strip()


def improved_planned_narrative(self, section, subject, practices, concern_yes, loc, size, action_language, om_language):
    selected = self.get_widget_value("Quick Build Wizard", f"{section} - Selected Practices")
    standards = self.get_widget_value(section, "NRCS Standards and Units Used")
    rows = _parse_bmp_rows(selected) or _parse_bmp_rows(standards)
    practice_sentence = _practice_names(rows, practices)
    dimension_sentence = _bmp_dimension_sentence(section, rows)
    loc = _clean(loc)
    size = _clean(size)
    no_structural = not practice_sentence or practice_sentence.lower() in {"no structural bmp - continue o&m", "continue o&m", "continued operation and maintenance"}
    parts = []
    if no_structural:
        parts.append(f"No additional structural Best Management Practice is recommended for {subject.lower()} at this time because no significant resource concern was identified during the evaluation.")
        parts.append(om_language or f"The farm should continue routine operation and maintenance for {subject.lower()}.")
        parts.append("If site conditions change or a resource concern develops, the area should be reevaluated and an appropriate BMP should be planned using applicable NRCS standards.")
    else:
        lead = "To reduce the identified resource concerns" if concern_yes else "To maintain current conditions and prevent future resource concerns"
        first = f"{lead}, it is recommended that {practice_sentence} be implemented"
        if loc:
            first += f" at {loc}"
        parts.append(first + ".")
        if dimension_sentence:
            parts.append(dimension_sentence)
        elif size:
            parts.append(f"The planned work should be sized to approximately {size}, subject to final design and field verification.")
        parts.append(_section_purpose(section))
        action_language = _remove_empty_words(_sent(action_language))
        if action_language:
            parts.append(action_language)
        parts.append(_design_sentence(section, rows))
        if om_language:
            parts.append(_sent(om_language))
    narrative = " ".join(p for p in parts if p).strip()
    narrative = _remove_empty_words(narrative)
    narrative = _dedupe_meaningful(narrative)
    return self.quality_check_narrative(narrative)


def improved_quality_check(self, text):
    text = str(text or "").strip()
    if not text:
        return ""
    try:
        text = _PREV_FINAL_EDITOR(self, text)
    except Exception:
        pass
    text = _remove_empty_words(text)
    text = _dedupe_meaningful(text)
    text = re.sub(r"\s+", " ", text).strip()
    if text and text[-1] not in ".!?:":
        text += "."
    return text


def safe_generate_docx(self):
    try:
        return _PREV_GENERATE_DOCX(self)
    except PermissionError as exc:
        messagebox.showerror(
            "Could not save Word document",
            "The Word document could not be saved because the file is open, locked, or you do not have permission to overwrite it.\n\n"
            "Close the existing Word document or choose a different file name, then generate the document again.\n\n"
            f"Details: {exc}"
        )
        return None


App.make_planned_narrative = improved_planned_narrative
App.quality_check_narrative = improved_quality_check
App.clean_paragraph = improved_quality_check
App.proofread_text = improved_quality_check
App.generate_docx = safe_generate_docx


if __name__ == "__main__":
    App().mainloop()
