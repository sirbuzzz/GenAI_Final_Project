---
name: lab-report-analyzer
description: Reads a patient lab results document (PDF, Word, or Excel) using Claude Vision, extracts all biomarker values and the patient's biological sex, matches biomarker names to a health coach's optimal reference sheet using LLM-based semantic matching with LLM-as-judge validation, and produces a color-coded PDF report with gender-specific optimal ranges where applicable. Also generates a plain-English coach summary highlighting key patterns and priority concerns.
---

## When to use this skill
- A patient lab document has been uploaded and a comparison report is needed
- The user asks to "analyze my labs", "run the report", or "compare my labs to the reference"
- The lab document is a PDF, Word (.docx), or Excel (.xlsx) file — all formats supported

## When NOT to use this skill
- No lab document has been provided
- The user is asking for medical advice, diagnoses, or treatment recommendations
- The file provided is not a lab results document

## Inputs
- **Lab results document** — uploaded by the user at runtime (PDF, .docx, or .xlsx). Any PDF format supported — text-based, scanned, or vector PDF.
- **Reference Excel** — `references/Optimal Lab Ranges.xlsx` (static, already in the skill folder)

## Workflow

### Step 1 — Convert document to images
Use PyMuPDF to render each page as a high-resolution PNG image (2x zoom).
All three file formats (PDF, DOCX, XLSX) are converted through the same image pipeline,
making them all compatible with Claude Vision.

### Step 2 — Claude Vision: Extract biomarkers + patient sex (Call 1)
Send all page images to Claude. Claude reads the document visually and returns structured JSON:
```json
{
  "sex": "male" | "female" | null,
  "panels": {
    "Panel Name": [
      {"name": "...", "value": 0.0, "flag": null, "lab_low": 0.0, "lab_high": 0.0, "units": "..."}
    ]
  }
}
```
Patient biological sex is read directly from the document demographics header — no manual entry needed.
This handles all PDF formats, including vector-path PDFs where text parsers return empty output.

### Step 3 — Claude: Semantic name matching (Call 2)
The extracted biomarker names and the full reference sheet name list are sent to Claude.
Claude matches each name by biological meaning, handling:
- Abbreviations: "AST (SGOT)" → "AST", "ALT (SGPT)" → "ALT"
- Short forms: "Hemoglobin" → "HGB", "Hematocrit" → "HCT"
- Alternate names: "Carbon Dioxide, Total" → "CO2/BICARBONATE/CARBON DIOXIDE"
- Verbose names: "Alanine Aminotransferase" → "ALT"

### Step 4 — Claude: LLM-as-judge validation (Call 3)
Claude reviews its own matches and assigns a confidence label to each:
- `confident` — clearly the same biomarker, name difference is formatting only
- `uncertain` — possible match but methodological or semantic ambiguity exists
- `wrong` — different biomarkers

Only `confident` matches are passed to the report. `uncertain` matches are surfaced in the UI for manual review.

### Step 5 — Skill: Compare values with gender-specific ranges
`analyze_labs.py` loads the reference Excel and classifies each confirmed match.
For biomarkers with gender-specific optimal ranges (RBC, HGB, HCT, DHEA-S, Ferritin,
ESR/Sed Rate, Uric Acid, SHBG), the correct male or female range is selected based on
the sex detected in Step 2. Every biomarker is always compared against the coach's optimal
range — the conventional lab reference ranges printed on the report are never used as the standard.

Status classifications:
- **Green (In Range)** — within the coach's optimal range
- **Orange (Slightly Off)** — outside optimal but within the lab's standard range
- **Red (Out of Range)** — outside both ranges

### Step 6 — Skill: Generate color-coded PDF report
`analyze_labs.py` builds the PDF report using reportlab, organized by panel.
For gender-specific biomarkers, the Optimal Range column shows the sex-appropriate range
(e.g. "Male: 14.0–15.0" or "Female: 13.5–14.5").

### Step 7 — Claude: Coach summary (Call 4)
Claude receives all flagged biomarkers and writes a 3–5 sentence plain-English summary
highlighting cross-panel patterns and priority concerns for the health coach.

## Output
1. **Color-coded PDF report** — organized by panel, downloadable
2. **Coach summary** — plain-English paragraph shown in the app UI
3. **Status breakdown** — count of In Range / Slightly Off / Out of Range markers
4. **Match validation results** — Confident / Uncertain / Wrong counts shown in UI

## Limitations
- Claude Vision accuracy depends on image quality — very low-res or damaged scans may miss values
- Units are not converted — assumes the lab document and reference sheet use the same units
- LLM matching and validation can produce slightly different confidence labels across runs for borderline cases
- Uncertain matches require manual review before acting on the report
- Does not interpret results medically or make treatment recommendations
- Optimal ranges reflect the coach's reference sheet, not standard clinical ranges
