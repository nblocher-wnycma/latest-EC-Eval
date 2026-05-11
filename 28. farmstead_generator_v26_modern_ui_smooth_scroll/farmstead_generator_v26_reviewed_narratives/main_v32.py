"""Farmstead Evaluation Generator v32 runner.

Builds on the current working v30 code and adds a CNMP proofreading/correction
report before export. This does not rewrite the document automatically. It reviews
outgoing generated text and produces a correction report in the format requested:

CRITICAL CORRECTIONS
GENERAL CORRECTIONS
CONSISTENCY CHECKS
VERIFY

Run with Run_Current_Working_Version.bat or Run_Farmstead_Eval_v32.bat.
"""

import os
import re
import runpy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V30 = os.path.join(BASE_DIR, "main_v30.py")

base = runpy.run_path(BASE_V30, run_name="farmstead_generator_v30")
App = base["App"]
SECTIONS = base["SECTIONS"]
QUICK_BUILD_SECTIONS = base["QUICK_BUILD_SECTIONS"]

_ORIGINAL_CREATE_UI = App.create_ui
_ORIGINAL_GENERATE_DOCX = App.generate_docx

KNOWN_REFERENCE_FARMS = [
    "Yousett Farm",
    "Big Z Beef",
    "A&M Hideaway",
    "A and M Hideaway",
    "Ketch Farm",
    "Maple Row Farm",
]

COMMON_TYPO_FIXES = {
    "cattlee": "cattle",
    "cattleee": "cattle",
    "cattleeeee": "cattle",
    "cattleeeeeee": "cattle",
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
}

UNIT_CANONICAL = {
    "sqf": "square feet",
    "sf": "square feet",
    "sq ft": "square feet",
    "sq. ft": "square feet",
    "sq. ft.": "square feet",
    "sqft": "square feet",
    "ln ft": "linear feet",
    "lin ft": "linear feet",
    "lf": "linear feet",
    "l.f.": "linear feet",
    "ft": "feet",
}


def _safe_text(value):
    return str(value or "").strip()


def _sentences(text):
    return [s.strip() for s in re.findall(r"[^.!?]+[.!?]|[^.!?]+$", _safe_text(text)) if s.strip()]


def _normalize_key(sentence):
    key = sentence.lower()
    key = re.sub(r"\([^)]*\)", "", key)
    key = re.sub(r"[^a-z0-9\s]", " ", key)
    key = re.sub(r"\b(the|a|an|and|or|to|of|for|in|on|at|as|by|with|from|this|that|these|those|it|is|are|was|were|be|been|being|should|continue|continued|planned|work|intended|practice|practices)\b", " ", key)
    return re.sub(r"\s+", " ", key).strip()


def _correction_item(original, issue, suggested, reason):
    return {
        "original": _safe_text(original),
        "issue": _safe_text(issue),
        "suggested": _safe_text(suggested),
        "reason": _safe_text(reason),
    }


def _format_item(item):
    return (
        f"- Original text: {item['original']}\n"
        f"  Issue: {item['issue']}\n"
        f"  Suggested correction: {item['suggested']}\n"
        f"  Reason: {item['reason']}"
    )


def _format_report(report):
    lines = []
    sections = [
        ("CRITICAL CORRECTIONS", report["critical"]),
        ("GENERAL CORRECTIONS", report["general"]),
        ("CONSISTENCY CHECKS", report["consistency"]),
        ("VERIFY", report["verify"]),
    ]
    for title, items in sections:
        lines.append(title)
        if not items:
            lines.append("- None found.")
        else:
            for item in items:
                lines.append(_format_item(item))
        lines.append("")
    return "\n".join(lines).strip()


def _suggest_unit_fix(text):
    fixed = text
    # Longer patterns first so 'sq ft' is caught before 'ft'.
    patterns = sorted(UNIT_CANONICAL.items(), key=lambda kv: len(kv[0]), reverse=True)
    for raw, canon in patterns:
        fixed = re.sub(r"\b" + re.escape(raw) + r"\b", canon, fixed, flags=re.I)
    return fixed


def _shorten_long_list_sentence(sentence):
    sentence = _safe_text(sentence)
    replacements = {
        "ditches, wetlands, streams, wells, poorly drained soils, or other sensitive areas": "nearby sensitive areas",
        "manure, bedding, feed waste, livestock traffic areas, and other potential contamination sources": "agricultural waste sources",
        "manure handling and livestock use areas so rainfall and snowmelt remain separated from agricultural waste": "manure handling and livestock use areas",
        "manure, sediment, nutrients, leachate, or other contaminants": "contaminants",
        "odor, dust, ammonia, methane, or nuisance conditions": "air quality concerns",
    }
    fixed = sentence
    for old, new in replacements.items():
        fixed = fixed.replace(old, new)
    return fixed


def _analyze_text_block(label, text, farm_name, report):
    text = _safe_text(text)
    if not text:
        report["critical"].append(_correction_item(
            f"{label}: [blank]",
            "Missing section text.",
            "Add a completed narrative or confirm this section is not applicable.",
            "Blank narratives should be reviewed before export."
        ))
        return

    # Placeholders and blank values.
    for placeholder in ["[Farm Name]", "[address]", "[town]", "[county]", "[watershed]", "[units]", "[item #]", "[No narrative entered.]"]:
        if placeholder.lower() in text.lower():
            report["critical"].append(_correction_item(
                f"{label}: {placeholder}",
                "Placeholder or blank value appears in outgoing text.",
                "Replace the placeholder with the correct farm-specific value or remove the sentence.",
                "Placeholders should not appear in an exported CNMP section."
            ))

    # Leftover farm names from another farm.
    farm_lower = farm_name.lower().strip()
    for known in KNOWN_REFERENCE_FARMS:
        if known.lower() in text.lower() and known.lower() != farm_lower:
            report["critical"].append(_correction_item(
                f"{label}: {known}",
                "Possible leftover farm name from another evaluation.",
                f"Replace with {farm_name} if this section belongs to the current farm, or verify the reference is intentional.",
                "Incorrect farm names are critical corrections before export."
            ))

    # Spelling and obvious typing errors.
    repeated_letter_matches = re.findall(r"\b[A-Za-z]*([A-Za-z])\1{2,}[A-Za-z]*\b", text)
    if repeated_letter_matches:
        for word in sorted(set(re.findall(r"\b[A-Za-z]*[A-Za-z]{1}([A-Za-z])\1{2,}[A-Za-z]*\b", text))):
            pass
        bad_words = sorted(set(re.findall(r"\b[A-Za-z]*([A-Za-z])\1{2,}[A-Za-z]*\b", text)))
    for typo, correction in COMMON_TYPO_FIXES.items():
        if re.search(r"\b" + re.escape(typo) + r"\b", text, flags=re.I):
            report["general"].append(_correction_item(
                f"{label}: {typo}",
                "Spelling or typing error.",
                correction,
                "This appears to be an unintended spelling error in the outgoing text."
            ))

    for match in re.finditer(r"\b[A-Za-z]*([A-Za-z])\1{3,}[A-Za-z]*\b", text):
        original = match.group(0)
        suggested = re.sub(r"([A-Za-z])\1{2,}", r"\1", original)
        report["general"].append(_correction_item(
            f"{label}: {original}",
            "Possible accidental repeated letters.",
            suggested,
            "Words with repeated letters usually indicate a typing error."
        ))

    # Grammar/wording checks.
    for phrase in ["the the", "and and", "of the the", "is is", "are are"]:
        if re.search(r"\b" + re.escape(phrase) + r"\b", text, flags=re.I):
            report["general"].append(_correction_item(
                f"{label}: {phrase}",
                "Repeated word or phrase.",
                phrase.split()[0],
                "Repeated wording should be removed for readability."
            ))

    # Missing punctuation and incomplete sentence checks.
    paragraphs = [p.strip() for p in re.split(r"\n+", text) if p.strip()]
    for para in paragraphs:
        if para.startswith("•"):
            continue
        if para and para[-1] not in ".!?:":
            report["general"].append(_correction_item(
                f"{label}: {para}",
                "Missing ending punctuation.",
                para + ".",
                "Narrative paragraphs should end with punctuation."
            ))

    for sentence in _sentences(text):
        stripped = sentence.strip()
        words = stripped.split()
        if len(words) < 4 and not stripped.startswith("•"):
            report["verify"].append(_correction_item(
                f"{label}: {stripped}",
                "Possible incomplete sentence.",
                "Verify whether this sentence should be expanded or removed.",
                "Very short sentence fragments may not communicate a complete thought."
            ))

        comma_count = stripped.count(",")
        if comma_count >= 4 or len(words) > 42:
            shortened = _shorten_long_list_sentence(stripped)
            if shortened != stripped:
                report["general"].append(_correction_item(
                    f"{label}: {stripped}",
                    "Long comma list or wordy sentence.",
                    shortened,
                    "Shortening long lists improves clarity without removing technical meaning."
                ))
            else:
                report["verify"].append(_correction_item(
                    f"{label}: {stripped}",
                    "Long or complex sentence.",
                    "Verify whether this sentence can be shortened without losing technical meaning.",
                    "Long sentences can make the CNMP narrative harder to review."
                ))

        # Specific awkward/redundant wording seen in generated CNMP narratives.
        if "This creates a resource concern because" in stripped and "This creates a resource concern because" in text[text.find(stripped)+len(stripped):]:
            report["general"].append(_correction_item(
                f"{label}: {stripped}",
                "Repeated resource concern phrasing.",
                "Keep the strongest resource-concern sentence and remove repeated boilerplate.",
                "Repeated concern language makes the narrative sound redundant."
            ))

    # Sentence redundancy.
    seen = {}
    for sentence in _sentences(text):
        key = _normalize_key(sentence)
        if not key:
            continue
        if key in seen:
            report["general"].append(_correction_item(
                f"{label}: {sentence}",
                "Repeated or near-identical sentence.",
                "Remove this sentence or combine it with the earlier sentence.",
                "Duplicate sentences reduce clarity and make the section feel automated."
            ))
        else:
            seen[key] = sentence

    # Unit consistency.
    raw_unit_patterns = ["Sqf", "SQF", "SF", "sq ft", "Sq Ft", "sq. ft", "Ln Ft", "lin ft", "LF", "lf"]
    for unit in raw_unit_patterns:
        if re.search(r"\b" + re.escape(unit) + r"\b", text):
            report["consistency"].append(_correction_item(
                f"{label}: {unit}",
                "Inconsistent unit style.",
                _suggest_unit_fix(unit),
                "Use consistent unit terminology throughout the exported document."
            ))

    # Terminology consistency.
    if re.search(r"\bstorm water\b", text, flags=re.I):
        report["consistency"].append(_correction_item(
            f"{label}: storm water",
            "Inconsistent terminology.",
            "stormwater",
            "Use one consistent spelling throughout the report."
        ))

    # Farm-specific verification flags.
    if re.search(r"\b(no permanent manure storage|does not have a designated manure storage|temporary manure pile)\b", text, flags=re.I):
        report["verify"].append(_correction_item(
            f"{label}: manure storage statement",
            "Farm-specific fact requires confirmation.",
            "Verify this statement matches the current farm conditions.",
            "Manure storage facts vary by farm and should be confirmed before export."
        ))

    if re.search(r"\b(all exported|exported off farm|spread on rented land|applied to fields)\b", text, flags=re.I):
        report["verify"].append(_correction_item(
            f"{label}: manure application/export statement",
            "Farm-specific fact requires confirmation.",
            "Verify manure application/export details match the farm input.",
            "Application and export details directly affect CNMP accuracy."
        ))


def _analyze_standards_block(label, standards_text, report):
    text = _safe_text(standards_text)
    if not text:
        return
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.lower().startswith("standard"):
            continue
        parts = [p.strip() for p in stripped.split("|")]
        if len(parts) < 4:
            report["consistency"].append(_correction_item(
                f"{label}: {stripped}",
                "Standards row does not follow the expected format.",
                "Standard | Item # | Description | Units",
                "Consistent standards-table formatting helps the Word document export correctly."
            ))
            continue
        standard, item, desc, units = parts[:4]
        if not item or item.lower() in {"[item #]", "item #"}:
            report["critical"].append(_correction_item(
                f"{label}: {stripped}",
                "Missing item number in standards table.",
                "Enter the correct item number for this practice.",
                "Item numbers are needed to tie planned BMPs to the plan."
            ))
        if not units or units.lower() in {"[units]", "units"}:
            report["critical"].append(_correction_item(
                f"{label}: {stripped}",
                "Missing units/dimensions in standards table.",
                "Enter the planned units or dimensions for this practice.",
                "Planned-action narratives and standards tables rely on BMP units."
            ))
        if re.search(r"\b(Sqf|SF|sq ft|Sq Ft|sq\. ft|Ln Ft|lin ft|LF|lf)\b", units):
            report["consistency"].append(_correction_item(
                f"{label}: {units}",
                "Inconsistent units in standards table.",
                _suggest_unit_fix(units),
                "Use consistent unit wording in standards and narratives."
            ))


def build_correction_report(app):
    """Build a correction report from the outgoing generated text."""
    # Mirror the export workflow so the report reviews what would be exported.
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
    farm_name = _safe_text(basic.get("Farm Name")) or "[Farm Name]"
    report = {"critical": [], "general": [], "consistency": [], "verify": []}

    for field in [
        "Introduction Narrative",
        "Overview Narrative",
        "Animal Housing and Feed Management Narrative",
        "Manure Management and Application Narrative",
    ]:
        _analyze_text_block(f"Basic Info - {field}", basic.get(field, ""), farm_name, report)

    for sec in SECTIONS[1:]:
        sec_data = data.get(sec, {})
        _analyze_text_block(f"{sec} - Existing Conditions", sec_data.get("Existing Conditions narrative", ""), farm_name, report)
        _analyze_text_block(f"{sec} - Planned Actions", sec_data.get("Planned Actions narrative", ""), farm_name, report)
        _analyze_standards_block(f"{sec} - NRCS Standards and Units Used", sec_data.get("NRCS Standards and Units Used", ""), report)

    return _format_report(report)


def open_correction_report_tab(self):
    if hasattr(self, "proofreading_report_text"):
        self.notebook.select(self.proofreading_tab_index)
        return
    tab = ttk.Frame(self.notebook, padding=10)
    self.notebook.add(tab, text="Proofreading Report")
    self.proofreading_tab_index = self.notebook.index("end") - 1

    instructions = ttk.Label(
        tab,
        text="Generate a correction report from the outgoing document text. This does not rewrite the whole document; it flags issues to review before export.",
        wraplength=1100,
        justify="left",
    )
    instructions.pack(anchor="w", pady=(0, 8))

    button_row = ttk.Frame(tab)
    button_row.pack(fill="x", pady=(0, 8))
    ttk.Button(button_row, text="Generate Correction Report", command=self.generate_correction_report).pack(side="left", padx=(0, 6))
    ttk.Button(button_row, text="Save Report as TXT", command=self.save_correction_report).pack(side="left")

    self.proofreading_report_text = tk.Text(tab, wrap="word", height=35, font=("Consolas", 10), undo=True)
    self.proofreading_report_text.pack(fill="both", expand=True)
    self.notebook.select(self.proofreading_tab_index)


def generate_correction_report(self):
    if not hasattr(self, "proofreading_report_text"):
        open_correction_report_tab(self)
    report_text = build_correction_report(self)
    self.proofreading_report_text.delete("1.0", "end")
    self.proofreading_report_text.insert("1.0", report_text)
    self.notebook.select(self.proofreading_tab_index)


def save_correction_report(self):
    if not hasattr(self, "proofreading_report_text"):
        open_correction_report_tab(self)
    report_text = self.proofreading_report_text.get("1.0", "end").strip()
    if not report_text:
        report_text = build_correction_report(self)
        self.proofreading_report_text.insert("1.0", report_text)
    path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Report", "*.txt"), ("All files", "*.*")],
        initialfile="CNMP_Correction_Report.txt",
    )
    if not path:
        return
    with open(path, "w", encoding="utf-8") as f:
        f.write(report_text)
    messagebox.showinfo("Saved", f"Correction report saved:\n{path}")


def create_ui_with_proofreading(self):
    _ORIGINAL_CREATE_UI(self)
    # Add a simple footer button so the feature is easy to find without changing
    # the working toolbar internals.
    footer = ttk.Frame(self, padding=(12, 0, 12, 8))
    footer.pack(fill="x")
    ttk.Button(footer, text="Open / Generate CNMP Correction Report", command=self.generate_correction_report).pack(side="left")


def generate_docx_with_report_notice(self):
    # Build/update a correction report before export so the user can review it.
    # Do not block export; just make the report available in the tab.
    try:
        if not hasattr(self, "proofreading_report_text"):
            open_correction_report_tab(self)
        report_text = build_correction_report(self)
        self.proofreading_report_text.delete("1.0", "end")
        self.proofreading_report_text.insert("1.0", report_text)
    except Exception as exc:
        messagebox.showwarning("Correction report warning", f"The correction report could not be generated before export.\n\nDetails: {exc}")
    return _ORIGINAL_GENERATE_DOCX(self)


App.create_ui = create_ui_with_proofreading
App.open_correction_report_tab = open_correction_report_tab
App.generate_correction_report = generate_correction_report
App.save_correction_report = save_correction_report
App.generate_docx = generate_docx_with_report_notice


if __name__ == "__main__":
    App().mainloop()
