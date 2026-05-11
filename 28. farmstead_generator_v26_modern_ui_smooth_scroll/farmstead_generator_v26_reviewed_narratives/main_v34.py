"""Farmstead Existing Conditions Evaluation Generator v34.

This version keeps v30 as the stable logic base, keeps the v33 UI refresh, and
adds EC-focused narrative and proofreading improvements. It intentionally stays
focused on Existing Conditions and Planned Actions evaluations, not full CNMP
document assembly.
"""

import os
import re
import runpy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V33 = os.path.join(BASE_DIR, "main_v33.py")

base = runpy.run_path(BASE_V33, run_name="farmstead_generator_v33")
App = base["App"]

_METHOD_GLOBALS = App.create_ui.__globals__
SECTIONS = _METHOD_GLOBALS.get("SECTIONS", [
    "Basic Info", "Barnyard", "Clean Stormwater", "Livestock Manure Storage",
    "Pasture", "Animal Mortalities", "Fuel / Petroleum", "Air Quality",
    "Driveways", "Feed Storage",
])
QUICK_BUILD_SECTIONS = _METHOD_GLOBALS.get("QUICK_BUILD_SECTIONS", [s for s in SECTIONS if s != "Basic Info"])

_PREV_CREATE_UI = App.create_ui
_PREV_PLANNED_NARRATIVE = App.make_planned_narrative
_PREV_EXPORT = App.generate_docx.__globals__.get("_ORIGINAL_GENERATE_DOCX", App.generate_docx)

KNOWN_REFERENCE_FARMS = ["Yousett Farm", "Ketch Farm", "A&M Hideaway", "A and M Hideaway", "Big Z Beef", "Maple Row Farm"]

COMMON_TYPOS = {
    "cattlee": "cattle", "cattleee": "cattle", "horss": "horses", "witner": "winter",
    "inclimate": "inclement", "suplimentaly": "supplementally", "supplementaly": "supplementally",
    "suplimental": "supplemental", "mnnure": "manure", "approprite": "appropriate",
    "acces": "access", "draned": "drained", "tnaks": "tanks", "mortaliy": "mortality",
    "inbetween": "between", "landbase": "land base", "drive ways": "driveways",
}
TERM_FIXES = {
    "BMP's": "BMPs", "bmp's": "BMPs", "OandM": "O&M", "oandm": "O&M",
    "clean rain water": "clean rainwater", "storm water": "stormwater", "run off": "runoff",
    "manure-laden run off": "manure-laden runoff",
}
UNIT_FIXES = {
    "sqf": "square feet", "sf": "square feet", "sq ft": "square feet",
    "sq. ft": "square feet", "sq. ft.": "square feet", "sqft": "square feet",
    "ln ft": "linear feet", "lin ft": "linear feet", "lf": "linear feet", "l.f.": "linear feet",
}
STOP_WORDS = {
    "the", "a", "an", "and", "or", "to", "of", "for", "in", "on", "at", "as", "by",
    "with", "from", "this", "that", "these", "those", "it", "is", "are", "was", "were",
    "be", "been", "being", "should", "continue", "continued", "planned", "work", "intended",
    "practice", "practices", "area", "farm",
}

RESOURCE_LANGUAGE = {
    "Barnyard": "Under the conditions described, livestock concentration, unstable surfaces, manure accumulation, or exposed soil can create a source of sediment, nutrients, pathogens, and manure-laden runoff during rainfall or snowmelt events.",
    "Clean Stormwater": "This should be treated as a resource concern when clean roof runoff or upslope stormwater is allowed to contact manure, bedding, feed waste, temporary manure piles, or livestock traffic areas before leaving the farmstead.",
    "Livestock Manure Storage": "This should be treated as a resource concern when manure or bedding is stored in an uncontrolled location, exposed to precipitation, or located where leachate or manure-laden runoff can reach ditches, streams, wetlands, wells, poorly drained soils, or other sensitive areas.",
    "Pasture": "This should be treated as a resource concern when grazing pressure, poor livestock distribution, bare soil, high-use areas, or stream access can allow sediment and nutrients to move from the pasture.",
    "Animal Mortalities": "This should be treated as a resource concern when mortalities do not have a defined management method or location that limits leachate, runoff, scavenger access, and impacts to wells or water resources.",
    "Fuel / Petroleum": "This should be treated as a resource concern when tanks or containers lack containment, spill response, protection from traffic, or adequate separation from soil and water resources.",
    "Air Quality": "This should be treated as a resource concern when manure handling, livestock concentration, field application, feed handling, or traffic areas create odor, dust, ammonia, methane, smoke, or nuisance impacts.",
    "Driveways": "This should be treated as a resource concern when access lanes are rutted, muddy, unstable, poorly drained, or tracking manure and sediment toward drainageways or sensitive areas.",
    "Feed Storage": "This should be treated as a resource concern when spoiled feed, feed runoff, silage leachate, or clean stormwater contact with feed waste can move nutrients or organic material from the storage area.",
}
NO_CONCERN_MONITORING = {
    "Barnyard": "The farm should continue to monitor livestock traffic, surface stability, manure buildup, and runoff during wet weather and high-use periods.",
    "Clean Stormwater": "The farm should continue to inspect and maintain gutters, downspouts, outlets, and discharge areas so clean water remains separated from manure and livestock use areas.",
    "Livestock Manure Storage": "The farm should continue to monitor manure handling areas for leachate, ponding, runoff, clean water contact, and changes in storage needs.",
    "Pasture": "The farm should continue to monitor forage cover, livestock distribution, high-use areas, water locations, and any drainage or stream access concerns.",
    "Animal Mortalities": "The farm should continue to follow applicable New York State mortality disposal requirements and maintain separation from wells and water resources.",
    "Fuel / Petroleum": "The farm should continue to inspect petroleum storage areas and maintain spill response, housekeeping, and containment practices appropriate for the site.",
    "Air Quality": "The farm should continue proper manure handling and application practices and remain mindful of nearby residences, public roads, and windy conditions.",
    "Driveways": "The farm should continue to grade, gravel, and maintain driveways and clean manure from traffic areas as needed.",
    "Feed Storage": "The farm should continue routine operation and maintenance of feed storage areas and monitor for spoiled feed, leachate, runoff contact, and muddy conditions.",
}


def _capitalize_sentences(text):
    return re.sub(r"(^|[.!?]\s+)([a-z])", lambda m: m.group(1) + m.group(2).upper(), text)


def _clean(text, sentence_end=False):
    text = str(text or "").strip()
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\bw/\b", "with", text, flags=re.I)
    text = re.sub(r"\bUO\b", "underground outlet", text)
    for bad, good in COMMON_TYPOS.items():
        text = re.sub(r"\b" + re.escape(bad) + r"\b", good, text, flags=re.I)
    for bad, good in TERM_FIXES.items():
        text = re.sub(r"\b" + re.escape(bad) + r"\b", good, text, flags=re.I)
    for bad, good in sorted(UNIT_FIXES.items(), key=lambda item: len(item[0]), reverse=True):
        text = re.sub(r"\b" + re.escape(bad) + r"\b", good, text, flags=re.I)
    text = re.sub(r"\bNo current resource concern\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\bNo structural BMP\s*-\s*continue O&M\.?\s*", "", text, flags=re.I)
    text = re.sub(r"\[item #\]|\[units\]", "", text, flags=re.I)
    text = text.replace("County County", "County").replace("Basin Basin", "Basin").replace("Watershed Watershed", "Watershed")
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([.!?]){2,}", r"\1", text)
    text = re.sub(r"\bthe the\b", "the", text, flags=re.I)
    text = re.sub(r"\b([A-Za-z]{4,})\s+\1\b", r"\1", text, flags=re.I)
    text = re.sub(r"\.\s+And\s+", ". ", text)
    text = re.sub(r"\.\s+But\s+", ". However, ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = _capitalize_sentences(text)
    if sentence_end and text and text[-1] not in ".!?:":
        text += "."
    return text


def _sent(text):
    return _clean(text, sentence_end=True)


def _sentences(text):
    return [s.strip() for s in re.findall(r"[^.!?]+[.!?]|[^.!?]+$", str(text or "").strip()) if s.strip()]


def _sentence_key(text):
    key = re.sub(r"\([^)]*\)", "", text.lower())
    key = re.sub(r"[^a-z0-9\s]", " ", key)
    return {word for word in key.split() if word not in STOP_WORDS}


def _dedupe_sentences(text):
    output, seen = [], []
    for raw in _sentences(text):
        cleaned = _sent(raw)
        key = _sentence_key(cleaned)
        if not key:
            continue
        duplicate = any(key == old or (len(key & old) >= 7 and len(key & old) / max(len(key), 1) > 0.82) for old in seen)
        if not duplicate:
            output.append(cleaned)
            seen.append(key)
    return " ".join(output).strip()


def _strip_placeholder(value):
    text = str(value or "").strip()
    if not text or text.lower() in {"[item #]", "[units]", "[farm name]", "[address]", "[county]", "none", "n/a", "na"}:
        return ""
    return text


def _parse_standard_rows(raw, keep_incomplete=False):
    rows = []
    for line in str(raw or "").splitlines():
        line = line.strip()
        if not line or line.lower().startswith("standard"):
            continue
        if "|" not in line:
            if keep_incomplete:
                rows.append({"raw": line, "standard": line, "item": "", "desc": "", "units": ""})
            continue
        parts = [part.strip() for part in line.split("|")]
        parts += [""] * (4 - len(parts))
        standard, item, desc, units = parts[:4]
        rows.append({"raw": line, "standard": _clean(standard).strip(" ."), "item": _clean(item).strip(" ."), "desc": _clean(desc).strip(" ."), "units": _clean(units).strip(" .")})
    return rows


def _known_standard_map(app):
    for method_name in ["build_quick_build_wizard", "create_ui", "generate_docx"]:
        method = getattr(app, method_name, None)
        if method and hasattr(method, "__globals__"):
            values = method.__globals__.get("ALL_BMP_STANDARDS")
            if values:
                return {name.lower(): str(code) for name, code in values}
    return {}


def _practice_name_and_code(standard):
    match = re.search(r"^(.*?)\s*\((\d{3})\)", standard or "")
    if not match:
        return (re.sub(r"\s+", " ", standard or "").strip().lower(), "")
    return (match.group(1).strip().lower(), match.group(2).strip())


def _first_widget(self, section, names):
    for name in names:
        value = _strip_placeholder(self.get_widget_value(section, name))
        if value:
            return value
    return ""


def _section_intro(self, section, subject, summary):
    name = _first_widget(self, "Quick Build Wizard", [f"{section} - Condition Name"])
    loc = _first_widget(self, "Quick Build Wizard", [f"{section} - Location"])
    size = _first_widget(self, "Quick Build Wizard", [f"{section} - Size"])
    pieces = []
    if name and name.lower() not in summary.lower():
        pieces.append(name)
    if loc and loc.lower() not in summary.lower():
        pieces.append(f"at {loc}")
    if size and size.lower() not in summary.lower():
        pieces.append(f"covering approximately {size}")
    if not pieces:
        return ""
    return _sent(f"The {subject.lower()} {' '.join(pieces)} was evaluated during the farmstead assessment")


def improved_proofread_text(self, text):
    text = str(text or "").strip()
    if not text:
        return ""
    if "\n" in text or "\u2022" in text:
        lines = []
        for raw in text.splitlines():
            line = raw.strip()
            if not line:
                continue
            bullet = line.startswith("\u2022") or re.match(r"^[-*]\s+", line)
            line = re.sub(r"^[\u2022\-*]\s*", "", line).strip()
            cleaned = _clean(line, sentence_end=not bullet)
            if cleaned:
                lines.append(("\u2022 " if bullet else "") + cleaned.rstrip("." if bullet else ""))
        return "\n".join(lines)
    text = re.sub(r"\s*;\s*", ". ", text)
    return _dedupe_sentences(_sent(text))


def improved_quality_check(self, text):
    text = improved_proofread_text(self, text)
    if "\n" in text:
        return text
    weak_patterns = [
        r"The planned work is intended to improve day-to-day management, keep clean water separated from agricultural waste where applicable, and reduce the potential for manure, sediment, nutrients, leachate, or other contaminants to move from the farmstead\.",
        r"Final layout, grades, outlets, and construction details should be confirmed during design and installation so the practice functions as intended and meets applicable NRCS standards\.",
    ]
    for pattern in weak_patterns:
        text = re.sub(pattern, "", text, flags=re.I)
    return _dedupe_sentences(text)


def improved_existing_narrative(self, section, subject, summary, concern_yes, concern_language, no_concern_om):
    summary = improved_proofread_text(self, summary)
    if not summary:
        return ""
    parts = []
    intro = _section_intro(self, section, subject, summary)
    if intro:
        parts.append(intro)
    parts.append(summary)
    if concern_yes:
        parts.append(RESOURCE_LANGUAGE.get(section, concern_language))
        parts.append("The concern should be reviewed with the planned action for this section and with the applicable NRCS practice standards listed in the evaluation.")
    else:
        parts.append("Based on the conditions described, no significant resource concern was identified for this section at this time.")
        parts.append(NO_CONCERN_MONITORING.get(section, no_concern_om))
    return improved_quality_check(self, " ".join(part for part in parts if part))


def improved_planned_narrative(self, section, subject, practices, concern_yes, loc, size, action_language, om_language):
    return improved_quality_check(self, _PREV_PLANNED_NARRATIVE(self, section, subject, practices, concern_yes, loc, size, action_language, om_language))


def improved_proofread_all_report_text(self):
    for field in ["Introduction Narrative", "Animal Housing and Feed Management Narrative", "Manure Management and Application Narrative"]:
        value = self.get_widget_value("Basic Info", field)
        if value:
            self.set_widget_value("Basic Info", field, improved_quality_check(self, value))
    overview = self.get_widget_value("Basic Info", "Overview Narrative")
    if overview:
        self.set_widget_value("Basic Info", "Overview Narrative", improved_proofread_text(self, overview))
    for section in SECTIONS[1:]:
        for field in ["Existing Conditions narrative", "Planned Actions narrative"]:
            value = self.get_widget_value(section, field)
            if value:
                self.set_widget_value(section, field, improved_quality_check(self, value))


def _report_item(original, issue, suggested, reason):
    return {"original": _clean(original), "issue": _clean(issue), "suggested": _clean(suggested), "reason": _clean(reason)}


def _format_report_item(item):
    return f"- Original text: {item['original']}\n  Issue: {item['issue']}\n  Suggested correction: {item['suggested']}\n  Reason: {item['reason']}"


def _format_report(report):
    lines = []
    for title, key in [("CRITICAL CORRECTIONS", "critical"), ("GENERAL CORRECTIONS", "general"), ("CONSISTENCY CHECKS", "consistency"), ("VERIFY", "verify")]:
        lines.append(title)
        lines.extend([_format_report_item(item) for item in report[key]] or ["- None found."])
        lines.append("")
    return "\n".join(lines).strip()


def _analyze_text_block(label, text, farm_name, report):
    text = str(text or "").strip()
    if not text:
        report["critical"].append(_report_item(f"{label}: [blank]", "Missing section text.", "Add farm-specific notes and regenerate the narrative, or confirm this section is not applicable.", "Blank narratives should be reviewed before the Existing Conditions Evaluation is exported."))
        return
    for placeholder in ["[Farm Name]", "[address]", "[town]", "[county]", "[watershed]", "[units]", "[item #]", "[No narrative entered.]"]:
        if placeholder.lower() in text.lower():
            report["critical"].append(_report_item(f"{label}: {placeholder}", "Placeholder or blank value appears in outgoing text.", "Replace the placeholder with the correct farm-specific value or remove the sentence.", "Placeholders should not appear in a completed Existing Conditions Evaluation."))
    farm_lower = farm_name.lower().strip()
    for known in KNOWN_REFERENCE_FARMS:
        if known.lower() in text.lower() and known.lower() != farm_lower:
            report["critical"].append(_report_item(f"{label}: {known}", "Possible leftover farm name from another evaluation.", f"Replace with {farm_name} if this section belongs to the current farm, or verify the reference is intentional.", "Wrong farm names are critical corrections before export."))
    for typo, correction in COMMON_TYPOS.items():
        if re.search(r"\b" + re.escape(typo) + r"\b", text, flags=re.I):
            report["general"].append(_report_item(f"{label}: {typo}", "Spelling or typing error.", correction, "This appears to be an unintended field-note typo."))
    for phrase in ["the the", "and and", "of the the", "is is", "are are"]:
        if re.search(r"\b" + re.escape(phrase) + r"\b", text, flags=re.I):
            report["general"].append(_report_item(f"{label}: {phrase}", "Repeated word or phrase.", phrase.split()[0], "Repeated wording should be removed for readability."))
    for match in re.finditer(r"\b[A-Za-z]*([A-Za-z])\1{3,}[A-Za-z]*\b", text):
        original = match.group(0)
        report["general"].append(_report_item(f"{label}: {original}", "Possible accidental repeated letters.", re.sub(r"([A-Za-z])\1{2,}", r"\1", original), "Repeated letters usually indicate a typing error."))
    for para in [p.strip() for p in re.split(r"\n+", text) if p.strip()]:
        if not para.startswith("\u2022") and para[-1] not in ".!?:)":
            report["general"].append(_report_item(f"{label}: {para}", "Missing ending punctuation.", para + ".", "Narrative paragraphs should end with punctuation."))
    seen = []
    for sentence in _sentences(text):
        words = sentence.split()
        if len(words) < 4 and not sentence.startswith("\u2022"):
            report["verify"].append(_report_item(f"{label}: {sentence}", "Possible incomplete sentence.", "Verify whether this fragment should be expanded or removed.", "Short fragments may not communicate a complete observation."))
        if len(words) > 48 or sentence.count(",") >= 4:
            report["verify"].append(_report_item(f"{label}: {sentence}", "Long or complex sentence.", "Review for clarity and split if the technical meaning can be preserved.", "Long comma-list sentences make the evaluation harder to review."))
        key = _sentence_key(sentence)
        if key and any(key == old or (len(key & old) >= 7 and len(key & old) / max(len(key), 1) > 0.82) for old in seen):
            report["general"].append(_report_item(f"{label}: {sentence}", "Repeated or near-identical sentence.", "Remove this sentence or combine it with the earlier sentence.", "Duplicate language makes the output sound automated."))
        elif key:
            seen.append(key)
    for unit in ["Sqf", "SQF", "SF", "sq ft", "Sq Ft", "sq. ft", "Ln Ft", "lin ft", "LF", "lf"]:
        if re.search(r"\b" + re.escape(unit) + r"\b", text):
            report["consistency"].append(_report_item(f"{label}: {unit}", "Inconsistent unit style.", _clean(unit), "Use consistent unit terminology throughout the evaluation."))
    if re.search(r"\bstorm water\b", text, flags=re.I):
        report["consistency"].append(_report_item(f"{label}: storm water", "Inconsistent terminology.", "stormwater", "Use one spelling throughout the evaluation."))
    for fact_label, pattern in [("manure storage statement", r"\b(no permanent manure storage|temporary manure piles?|all manure is exported|exported off farm|spread on rented land)\b"), ("animal number or rate", r"\b\d+(\.\d+)?\s*(cows|calves|heifers|bulls|horses|goats|sheep|acres|gallons|tons|cubic feet|square feet|linear feet)\b"), ("separation distance", r"\b\d+(\.\d+)?\s*(feet|foot|ft|miles|yards)\b")]:
        if re.search(pattern, text, flags=re.I):
            report["verify"].append(_report_item(f"{label}: {fact_label}", "Farm-specific fact requires confirmation.", "Verify this statement against the entered field notes and supporting documents.", "Technical facts should be flagged for review instead of automatically changed."))


def _analyze_standards_block(app, label, text, planned_text, report):
    rows = _parse_standard_rows(text, keep_incomplete=True)
    if not rows:
        if planned_text and "no additional structural" not in planned_text.lower():
            report["verify"].append(_report_item(f"{label}: [blank]", "No NRCS standards table rows are listed for a planned action.", "Add standards rows if structural BMPs are planned, or confirm this is O&M only.", "The standards table ties planned actions to the practice items in the evaluation."))
        return
    known = _known_standard_map(app)
    for row in rows:
        raw = row.get("raw", "")
        if "|" not in raw:
            report["consistency"].append(_report_item(f"{label}: {raw}", "Standards row does not follow the expected table format.", "Standard | Item # | Description | Units", "Consistent formatting helps the Word export and proofreader read the table correctly."))
            continue
        if not row.get("item") or row.get("item", "").lower() in {"item #", "[item #]"}:
            report["critical"].append(_report_item(f"{label}: {raw}", "Missing item number in standards table.", "Enter the correct item number for this practice.", "Item numbers are needed to connect planned BMPs to the plan."))
        if not row.get("units") or row.get("units", "").lower() in {"units", "[units]"}:
            report["critical"].append(_report_item(f"{label}: {raw}", "Missing units or dimensions in standards table.", "Enter the planned units or dimensions, or mark the row for review.", "Planned-action narratives rely on BMP units and dimensions."))
        practice_name, code = _practice_name_and_code(row.get("standard", ""))
        expected = known.get(practice_name)
        if expected and code and code != expected:
            report["critical"].append(_report_item(f"{label}: {row.get('standard')}", "Practice name and NRCS standard number may not match.", f"Verify whether this should be {practice_name.title()} ({expected}).", "NRCS standard numbers are technical facts and should be reviewed before export."))
    standards_text = " ".join(row.get("standard", "") for row in rows)
    for code in sorted(set(re.findall(r"\((\d{3})\)", planned_text or "")) - set(re.findall(r"\((\d{3})\)", standards_text))):
        report["verify"].append(_report_item(f"{label}: ({code})", "Planned narrative references a practice number not listed in the standards table.", "Verify the standards table and planned action describe the same practices.", "Practice-table mismatches should be reviewed instead of automatically changed."))


def _build_report_data(app):
    app.refresh_front_page_narratives(show_msg=False)
    if "Quick Build Wizard" in app.data_widgets:
        for sec in QUICK_BUILD_SECTIONS:
            app.apply_quick_build_section(sec)
    else:
        for sec in SECTIONS[1:]:
            app.refresh_section_narratives(sec, show_msg=False)
    app.proofread_all_report_text()
    data = app.get_all_data()
    basic = data.get("Basic Info", {})
    farm_name = _strip_placeholder(basic.get("Farm Name")) or "[Farm Name]"
    report = {"critical": [], "general": [], "consistency": [], "verify": []}
    if farm_name == "[Farm Name]":
        report["critical"].append(_report_item("Basic Info - Farm Name: [blank]", "Farm name is missing.", "Enter the farm name exactly as it should appear in the evaluation.", "The proofreader cannot check wrong farm names without the current farm name."))
    for field in ["Introduction Narrative", "Overview Narrative", "Animal Housing and Feed Management Narrative", "Manure Management and Application Narrative"]:
        _analyze_text_block(f"Basic Info - {field}", basic.get(field, ""), farm_name, report)
    if farm_name != "[Farm Name]" and farm_name.lower() not in basic.get("Introduction Narrative", "").lower():
        report["critical"].append(_report_item("Basic Info - Introduction Narrative", "Current farm name does not appear in the introduction.", f"Include {farm_name} in the introduction sentence.", "The completed examples identify the farm clearly on the first page."))
    for sec in SECTIONS[1:]:
        sec_data = data.get(sec, {})
        existing = sec_data.get("Existing Conditions narrative", "")
        planned = sec_data.get("Planned Actions narrative", "")
        _analyze_text_block(f"{sec} - Existing Conditions", existing, farm_name, report)
        _analyze_text_block(f"{sec} - Planned Actions", planned, farm_name, report)
        _analyze_standards_block(app, f"{sec} - NRCS Standards and Units Used", sec_data.get("NRCS Standards and Units Used", ""), planned, report)
        concern = app.get_widget_value("Quick Build Wizard", f"{sec} - Concern").strip().lower()
        if concern.startswith("no") and re.search(r"\b(resource concern|contaminated runoff|leachate|sediment|nutrient)\b", existing, flags=re.I):
            report["verify"].append(_report_item(f"{sec} - Existing Conditions", "Narrative may describe a concern while Quick Build says no concern.", "Verify the concern selection or edit the narrative.", "The proofreader should flag possible contradictions between inputs and generated text."))
    return report


def build_ec_correction_report(app):
    return _format_report(_build_report_data(app))


def open_ec_correction_report_tab(self):
    if hasattr(self, "proofreading_report_text"):
        try:
            self.notebook.tab(self.proofreading_tab_index, text="EC Proofreading Report")
        except Exception:
            pass
        self.notebook.select(self.proofreading_tab_index)
        return
    tab = ttk.Frame(self.notebook, padding=10)
    self.notebook.add(tab, text="EC Proofreading Report")
    self.proofreading_tab_index = self.notebook.index("end") - 1
    ttk.Label(tab, text="Generate a correction report from the outgoing Existing Conditions Evaluation text. Simple wording issues may be cleaned automatically; technical facts are flagged for review.", wraplength=1100, justify="left").pack(anchor="w", pady=(0, 8))
    button_row = ttk.Frame(tab)
    button_row.pack(fill="x", pady=(0, 8))
    ttk.Button(button_row, text="Generate EC Correction Report", command=self.generate_correction_report).pack(side="left", padx=(0, 6))
    ttk.Button(button_row, text="Save Report as TXT", command=self.save_correction_report).pack(side="left")
    self.proofreading_report_text = tk.Text(tab, wrap="word", height=35, font=("Consolas", 10), undo=True)
    self.proofreading_report_text.pack(fill="both", expand=True)
    self.notebook.select(self.proofreading_tab_index)


def generate_ec_correction_report(self):
    open_ec_correction_report_tab(self)
    report_text = build_ec_correction_report(self)
    self.proofreading_report_text.delete("1.0", "end")
    self.proofreading_report_text.insert("1.0", report_text)
    self.notebook.select(self.proofreading_tab_index)
    return report_text


def save_ec_correction_report(self):
    open_ec_correction_report_tab(self)
    report_text = self.proofreading_report_text.get("1.0", "end").strip() or generate_ec_correction_report(self)
    path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Report", "*.txt"), ("All files", "*.*")], initialfile="EC_Evaluation_Correction_Report.txt")
    if not path:
        return
    with open(path, "w", encoding="utf-8") as f:
        f.write(report_text)
    messagebox.showinfo("Saved", f"Correction report saved:\n{path}")


def _rename_cnmp_widgets(widget):
    try:
        text = widget.cget("text")
        if isinstance(text, str) and "CNMP" in text:
            widget.configure(text=text.replace("CNMP", "EC"))
    except Exception:
        pass
    for child in widget.winfo_children():
        _rename_cnmp_widgets(child)


def create_ui_ec(self):
    _PREV_CREATE_UI(self)
    _rename_cnmp_widgets(self)


def generate_docx_ec(self):
    try:
        report = _build_report_data(self)
        open_ec_correction_report_tab(self)
        self.proofreading_report_text.delete("1.0", "end")
        self.proofreading_report_text.insert("1.0", _format_report(report))
        if report["critical"]:
            messagebox.showwarning("Review critical corrections", "The EC proofreader found critical corrections. Review the EC Proofreading Report before finalizing the Word document.")
    except Exception as exc:
        messagebox.showwarning("Correction report warning", f"The EC correction report could not be generated before export.\n\nDetails: {exc}")
    return _PREV_EXPORT(self)


App.make_existing_narrative = improved_existing_narrative
App.make_planned_narrative = improved_planned_narrative
App.proofread_text = improved_proofread_text
App.quality_check_narrative = improved_quality_check
App.clean_paragraph = improved_quality_check
App.proofread_all_report_text = improved_proofread_all_report_text
App.create_ui = create_ui_ec
App.open_correction_report_tab = open_ec_correction_report_tab
App.generate_correction_report = generate_ec_correction_report
App.save_correction_report = save_ec_correction_report
App.generate_docx = generate_docx_ec


if __name__ == "__main__":
    App().mainloop()
