"""Farmstead Evaluation Generator v31 runner.

This version builds on v30 and focuses on output quality:
- stronger final spelling cleanup for field-note mistakes such as cattleeeee
- cleaner BMP sizing language without nested parentheses
- shorter, more meaningful planned-action purpose statements
- less repetitive comma-list boilerplate

Run with Run_Farmstead_Eval_v31.bat.
"""

import os
import re
import runpy
from tkinter import messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V30 = os.path.join(BASE_DIR, "main_v30.py")

base = runpy.run_path(BASE_V30, run_name="farmstead_generator_v30")
App = base["App"]
SECTIONS = base["SECTIONS"]

_PREV_GENERATE_DOCX = App.generate_docx


def _collapse_bad_repeats(text):
    """Fix obvious accidental repeated letters without changing normal double letters."""
    text = str(text or "")
    # cattleeeee -> cattle, runnnnnnoff -> runoff, etc.
    text = re.sub(r"([A-Za-z])\1{2,}", r"\1", text)
    return text


def _strip_code(value):
    """Remove NRCS code parentheses for narrative prose only.

    The standards table still keeps the codes. Narrative sentences read better as
    'Dry Stack Pad covering...' instead of 'Dry Stack Pad (Waste Storage Facility (313))'.
    """
    value = str(value or "").strip()
    value = re.sub(r"\s*\(\s*\d{3}\s*\)", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" .")


def _clean(text):
    text = str(text or "").strip()
    if not text:
        return ""
    text = _collapse_bad_repeats(text)
    replacements = {
        "cattlee": "cattle",
        "cattleee": "cattle",
        "cattl": "cattle",
        "horss": "horses",
        "witner": "winter",
        "inclimate": "inclement",
        "suplimentaly": "supplementally",
        "supplementaly": "supplementally",
        "Mnnure": "Manure",
        "mnnure": "manure",
        "approprite": "appropriate",
        "acces": "access",
        "draned": "drained",
        "tnaks": "tanks",
        "mortaliy": "mortality",
        "inbetween": "between",
        "landbase": "land base",
        "BMP's": "BMPs",
        "bmp's": "BMPs",
        "OandM": "O&M",
        "oandm": "O&M",
        "clean rain water": "clean rainwater",
        "storm water": "stormwater",
        "run off": "runoff",
        "manure-laden run off": "manure-laden runoff",
        " w/ ": " with ",
        "UO": "underground outlet",
    }
    for old, new in replacements.items():
        text = re.sub(r"\b" + re.escape(old) + r"\b", new, text, flags=re.IGNORECASE)

    unit_replacements = {
        "Sq. Ft.": "square feet",
        "Sq. Ft": "square feet",
        "Sq Ft": "square feet",
        "sq. ft.": "square feet",
        "sq. ft": "square feet",
        "sq ft": "square feet",
        "LF": "linear feet",
        "Lf.": "linear feet",
        "Lf": "linear feet",
        "lf.": "linear feet",
        "lf": "linear feet",
    }
    for old, new in unit_replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\[item #\]|\[units\]", "", text, flags=re.I)
    text = text.replace("County County", "County").replace("Basin Basin", "Basin").replace("Watershed Watershed", "Watershed")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = re.sub(r"\.\s+And\s+", ". ", text)
    text = re.sub(r"\.\s+But\s+", ". However, ", text)
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{4,})\s+\1\b", r"\1", text, flags=re.I)

    def cap(match):
        return match.group(1) + match.group(2).upper()
    text = re.sub(r"(^|[.!?]\s+)([a-z])", cap, text.strip())
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
        if not line or line.lower().startswith("standard") or "|" not in line:
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
    desc = _strip_code(row.get("desc", ""))
    practice = _strip_code(row.get("practice", ""))
    if not desc:
        return practice
    # Prefer the user's description. It is usually the field-friendly BMP name:
    # Dry Stack Pad, Gutter, Gravel Pad with Geotextile, etc.
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
        "Barnyard": "Planned BMP sizing includes",
        "Clean Stormwater": "Planned clean-water system sizing includes",
        "Livestock Manure Storage": "Planned manure-storage sizing includes",
        "Pasture": "Planned pasture-system sizing includes",
        "Animal Mortalities": "Planned mortality-management sizing includes",
        "Fuel / Petroleum": "Planned petroleum-storage sizing includes",
        "Driveways": "Planned access sizing includes",
        "Feed Storage": "Planned feed-storage sizing includes",
        "Air Quality": "Planned air-quality management items include",
    }
    return _sent(prefix_by_section.get(section, "Planned BMP sizing includes") + " " + _list_phrase(phrases))


def _practice_names(rows, fallback):
    if rows:
        return _list_phrase([_strip_code(r["practice"]) for r in rows if r.get("practice")])
    return _strip_code(_clean(fallback))


def _section_purpose(section):
    # Shorter, less repetitive purpose language. Avoids long comma lists.
    purposes = {
        "Barnyard": "These practices will stabilize livestock traffic areas, improve manure collection, and reduce contaminated runoff from the production area.",
        "Clean Stormwater": "These practices will keep clean stormwater separated from manure and other agricultural waste sources.",
        "Livestock Manure Storage": "These practices will provide controlled manure storage and reduce leachate and runoff risk.",
        "Pasture": "These practices will improve livestock distribution, maintain forage cover, and reduce soil and nutrient loss.",
        "Animal Mortalities": "These practices will provide a defined mortality-management area and reduce leachate, runoff, and scavenger concerns.",
        "Fuel / Petroleum": "These practices will reduce the risk of fuel spills reaching soil or water resources.",
        "Air Quality": "These actions will limit odor, dust, and nuisance impacts around the farmstead.",
        "Driveways": "These practices will improve access while reducing sediment, manure tracking, and runoff concerns.",
        "Feed Storage": "These practices will control feed waste, leachate, and runoff from the storage area.",
    }
    return purposes.get(section, "These practices will address the observed resource concern and improve long-term operation and maintenance.")


def _design_sentence(section, rows):
    practices = " ".join(r.get("practice", "") + " " + r.get("desc", "") for r in rows).lower()
    if section == "Livestock Manure Storage":
        return "Final design should confirm storage layout, grades, access, roof runoff handling, outlet locations, and operation needs."
    if section == "Clean Stormwater" or "roof runoff" in practices or "underground outlet" in practices:
        return "Outlet locations should be stable and protected from erosion."
    if section == "Barnyard":
        return "Final grades, curbs, access points, and scraping routes should support daily operation and runoff control."
    if section == "Pasture":
        return "Fence, water, and laneway locations should support rotation and forage recovery."
    return "Final layout and installation details should be confirmed during design."


def _trim_redundant_boilerplate(text):
    text = str(text or "")
    replacements = {
        r"This creates a resource concern because manure stored in temporary or uncontrolled locations can generate leachate and contaminated runoff, especially when exposed to precipitation or located near ditches, wetlands, streams, wells, poorly drained soils, or other sensitive areas\.":
            "This creates a resource concern because uncontrolled manure storage can generate leachate and contaminated runoff.",
        r"This work is intended to eliminate or reduce uncontrolled temporary manure piles, contain manure and bedding in a defined storage area, and reduce the potential for leachate or manure-laden runoff to leave the farmstead\.":
            "This work will provide controlled manure storage and reduce leachate and runoff risk.",
        r"This work is intended to keep clean roof runoff and upslope stormwater separated from manure, bedding, feed waste, livestock traffic areas, and other potential contamination sources\.":
            "This work will keep clean stormwater separated from agricultural waste sources.",
        r"The planned system should collect and convey clean water away from manure handling and livestock use areas so rainfall and snowmelt remain separated from agricultural waste\.":
            "The planned system will convey clean water away from manure handling and livestock use areas.",
        r"The planned work is intended to improve day-to-day management, keep clean water separated from agricultural waste where applicable, and reduce the potential for manure, sediment, nutrients, leachate, or other contaminants to move from the farmstead\.":
            "The planned work will improve management and reduce contaminant movement from the farmstead.",
    }
    for pattern, repl in replacements.items():
        text = re.sub(pattern, repl, text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _dedupe_meaningful(text):
    text = _trim_redundant_boilerplate(text)
    sentences = re.findall(r"[^.!?]+[.!?]|[^.!?]+$", str(text or "").strip())
    output = []
    seen = []
    stop = {"the", "a", "an", "and", "or", "to", "of", "for", "in", "on", "at", "as", "by", "with", "from", "this", "that", "these", "those", "it", "is", "are", "was", "were", "be", "been", "being", "should", "continue", "continued", "planned", "work", "intended", "practices", "practice"}
    for s in sentences:
        s = _sent(s)
        words = [w for w in re.sub(r"[^a-z0-9\s]", " ", s.lower()).split() if w not in stop]
        key = set(words)
        if not key:
            continue
        duplicate = False
        for old in seen:
            if len(key & old) / max(len(key), 1) > 0.72 and len(key & old) >= 6:
                duplicate = True
                break
        if not duplicate:
            output.append(s)
            seen.append(key)
    return " ".join(output).strip()


def quality_check(self, text):
    text = str(text or "").strip()
    if not text:
        return ""
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
            line = _dedupe_meaningful(_sent(line))
            if bullet:
                line = line.rstrip(".")
            lines.append(bullet + line if line else "")
        return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    text = _dedupe_meaningful(_sent(text))
    return text


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
        parts.append("If conditions change, the area should be reevaluated and an appropriate BMP should be planned.")
    else:
        lead = "To address the identified resource concerns" if concern_yes else "To maintain current conditions and prevent future resource concerns"
        first = f"{lead}, it is recommended that {practice_sentence} be implemented"
        if loc:
            first += f" at {loc}"
        parts.append(first + ".")
        if dimension_sentence:
            parts.append(dimension_sentence)
        elif size:
            parts.append(f"The planned work should be sized to approximately {size}, subject to final design and field verification.")
        parts.append(_section_purpose(section))
        parts.append(_design_sentence(section, rows))
        if om_language:
            parts.append(_sent(om_language))

    return quality_check(self, " ".join(p for p in parts if p))


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
App.quality_check_narrative = quality_check
App.clean_paragraph = quality_check
App.proofread_text = quality_check
App.generate_docx = safe_generate_docx


if __name__ == "__main__":
    App().mainloop()
