# Existing Conditions Evaluation Generator v34

This version is focused on improving the Existing Conditions and Planned Actions evaluation generator, not building a full CNMP generator.

## Base used

- `main_v30.py` remains the stable narrative and planned-BMP logic base.
- `main_v33.py` is loaded for the visual UI refresh.
- v31 is not used.

## What changed

- Existing-condition narratives now require farm-specific notes. If the section has no meaningful existing-condition note, the generated section is left blank and the proofreader flags it for review instead of filling the report with generic template language.
- Existing-condition narratives are expanded into a more professional evaluation style: observed condition, resource-concern explanation when selected, and review/monitoring language.
- Planned-action narratives still use the v30 planned BMP table and dimension handling, with an additional cleanup pass.
- The correction report is EC-focused instead of CNMP-focused.
- The proofreader flags critical corrections, general corrections, consistency checks, and items to verify before export.
- Technical facts such as animal numbers, acreages, distances, practice dimensions, manure export/application statements, and NRCS standards are flagged for review rather than automatically changed.
- The current working runner now opens `main_v34.py` on this branch.

## Proofreader checks added or strengthened

- Blank or missing section narratives
- Placeholders such as `[Farm Name]`, `[item #]`, and `[units]`
- Leftover reference farm names from completed examples
- Missing current farm name in the introduction
- Common field-note typos and repeated words
- Long comma-list sentences and incomplete fragments
- Inconsistent units such as `SF`, `sq ft`, `LF`, and `Ln Ft`
- Inconsistent terminology such as `storm water` vs. `stormwater`
- Missing standards-table item numbers or units
- Possible practice name and NRCS standard number mismatches
- Possible contradictions between the Quick Build resource-concern selection and generated text

## How to run

Use either:

- `Run_Current_Working_Version.bat`
- `Run_Farmstead_Eval_v34.bat`

Both run `main_v34.py` on this branch.
