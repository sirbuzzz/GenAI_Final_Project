# Lab Report Analyzer — Health Coach Workflow

A Streamlit app that reads patient lab reports (PDF, Word, or Excel), extracts biomarker values and patient sex using Claude Vision, validates name matching with an LLM-as-judge, and generates a color-coded PDF report against a functional medicine health coach's optimal reference ranges — including gender-specific ranges where applicable.

---

## Context, User, and Problem

**User:** A functional medicine health coach who reviews patient lab results as part of a health coaching workflow.

**Workflow being improved:** After a patient receives lab results (typically a Labcorp or Quest PDF), the coach manually reviews each biomarker and compares it against their own optimal reference ranges — which differ from the standard clinical ranges printed on the lab report. This involves opening the PDF, scanning each value, looking it up in a separate Excel sheet, and writing notes for the patient.

**Why this is hard:**
- Lab reports use vendor-specific naming conventions. The coach's reference sheet uses different names. "AST (SGOT)" on a Labcorp PDF is just "AST" in the reference; "Carbon Dioxide, Total" maps to "CO2/BICARBONATE/CARBON DIOXIDE." Manual mapping is tedious and error-prone.
- Standard clinical reference ranges and functional medicine optimal ranges are different. A glucose of 95 mg/dL is "normal" by Labcorp standards but above the coach's optimal range of 75–86. The coach needs comparison against *their* ranges, not the lab's.
- Several biomarkers (RBC, HGB, HCT, DHEA-S, Ferritin, ESR, Uric Acid, SHBG) have different optimal ranges for men and women. Using the wrong range misclassifies otherwise healthy values.
- Lab PDFs are often vector-path or image-based — standard text extraction libraries like pdfplumber fail on them silently, returning empty or garbled output.

**Why GenAI is useful here:**
- Claude Vision can read any PDF format visually, the same way a human reads it, without depending on the PDF's text layer. It also reads patient demographics (sex, DOB) from the same document in one pass.
- LLM-based name matching handles the semantic gap between lab naming conventions and the reference sheet — a task that keyword matching or fuzzy string matching cannot reliably solve.
- A language model can synthesize flagged biomarkers into a coherent coach-facing summary that highlights patterns across panels, not just a list of out-of-range values.

---

## Solution and Design

### What was built

A Streamlit web app with a reusable skill at its core. The user uploads a lab document, clicks **Generate Report**, and receives a downloadable color-coded PDF report plus a plain-English coach summary.

The app supports PDF, Word (.docx), and Excel (.xlsx) uploads. PyMuPDF converts each document to high-resolution PNG images, making all three formats compatible with the same Vision pipeline.

### Architecture — 4 Claude API calls

```
Upload → Images → [Claude Vision] → [Claude Matching] → [Claude Validation] → Skill → [Claude Summary]
```

**Call 1 — Vision Extraction (`claude-sonnet-4-6`)**
All page images are sent to Claude with a prompt to extract every biomarker result organized by panel and also identify the patient's biological sex from the document header. Returns structured JSON with both panels and sex. Real Labcorp PDFs always include sex on the patient demographics page — Claude reads it in the same pass as the lab values.

**Call 2 — Name Matching (`claude-sonnet-4-6`)**
The extracted PDF names and the full list of reference sheet names are sent to Claude. Claude matches each PDF biomarker to its equivalent in the reference sheet by biological meaning — handling abbreviations, alternate names, and formatting differences.

**Call 3 — LLM-as-Judge Validation (`claude-sonnet-4-6`)**
Claude reviews its own matches and labels each one:
- `confident` — clearly the same biomarker, name difference is formatting only
- `uncertain` — possible match but methodological or semantic ambiguity exists
- `wrong` — different biomarkers

Only `confident` matches are included in the final report. `uncertain` matches are shown in the UI for manual coach review.

**Skill — Report Generation (deterministic)**
`analyze_labs.py` (the pre-built skill) loads the reference Excel and classifies each matched biomarker using the detected patient sex for gender-specific biomarkers:
- **Green (In Range)** — within the coach's optimal range
- **Orange (Slightly Off)** — outside optimal but within Labcorp's standard range
- **Red (Out of Range)** — outside both ranges

The skill generates the formatted color-coded PDF using reportlab.

**Call 4 — Coach Summary (`claude-sonnet-4-6`)**
Claude receives all flagged biomarkers and writes a 3–5 sentence plain-English summary for the coach, highlighting cross-panel patterns and priority concerns.

### Key design choices

- **Claude Vision over text parsing.** pdfplumber was the original approach but fails on vector-path PDFs (which is how most Labcorp PDFs are generated). Claude Vision reads the document visually and is format-agnostic.
- **Sex extracted from the document, not from a form field.** Real lab reports always include patient sex on the demographics page. Rather than asking the coach to enter it manually, Claude reads it from the same document in the Vision call.
- **Gender-specific ranges for 8 biomarkers.** RBC, HGB, HCT, DHEA-S, Ferritin, ESR/Sed Rate, Uric Acid, and SHBG have separate male and female optimal ranges in the reference sheet. The app reads the patient's sex directly from the lab document and applies the correct range. Every matched biomarker is always compared against the coach's optimal range — the conventional lab ranges printed on the report are never used as the standard.
- **Separation of concerns.** GenAI handles what code cannot do well: reading messy documents, resolving name ambiguity, writing summaries. Deterministic code handles what GenAI should not do: numeric range comparison, status classification, PDF layout.
- **Self-validation.** The LLM-as-judge step catches matching errors before they reach the report. This is important because a wrong match (e.g., "Serum Creatinine" matched to "CREATINE") would produce a misleading status classification with no warning.
- **Confidence filtering.** Only confident matches go into the report. Uncertain matches are surfaced for manual review rather than silently included or silently dropped.

---

## Evaluation and Results

### Baseline comparison

The baseline is the coach's current manual workflow:

1. Receive the patient's lab PDF from Labcorp or another provider
2. Open a blank Excel sheet and transcribe each biomarker name and result value by hand
3. Pull out the reference sheet and look up each biomarker one by one against the optimal range
4. Write in the status for each biomarker across all panels (CBC, Metabolic, Lipid, etc.)
5. Recall or look up what out-of-range values mean and write notes for the patient

This process is time-consuming (dozens of biomarkers per patient), error-prone (manual transcription and comparison), inconsistent across sessions, and doesn't scale as the client base grows. It also requires the coach to mentally apply the correct gender-specific range for each applicable biomarker.

This app replaces the entire workflow with a single file upload. The coach gets a color-coded PDF report and a plain-English summary in seconds.

| | Manual baseline | This app |
|---|---|---|
| Biomarker extraction | Manual transcription | Claude Vision — any PDF format |
| Name matching | Coach looks up each name by memory | Claude semantic matching |
| Range comparison | Manual lookup against reference sheet | Automated against optimal ranges |
| Gender-specific ranges | Coach applies from memory | Detected from document, applied automatically |
| Bad match detection | None | LLM-as-judge confidence scoring |
| Coach summary | Coach writes notes by hand | Generated automatically |

A simple code alternative (pdfplumber + keyword matching) was also tested and returned zero biomarkers on a real Labcorp PDF — the PDF uses vector paths rather than a text layer, which text parsers cannot read. Claude Vision extracted all 47 biomarkers from the same file.

### Test cases

Eight synthetic Labcorp-style lab reports were created in PDF, Word (.docx), and Excel (.xlsx) formats — 24 files total. Each includes a realistic patient header with name, DOB, and sex.

| Case | Patient | Sex | Purpose |
|---|---|---|---|
| `normal_1` | James Carter | Male | Clean routine checkup, standard Labcorp names |
| `normal_2` | Sarah Mitchell | Female | Clean routine checkup, female-specific panels (DHEA-S) |
| `edge_1` | Robert Davis | Male | Metabolic syndrome — elevated glucose, insulin, triglycerides |
| `edge_2` | Maria Torres | Female | Thyroid dysfunction + iron-deficiency anemia |
| `edge_3` | William Chen | Male | Critical — kidney failure, severe anemia, uncontrolled diabetes |
| `judge_1` | Jennifer Lopez | Female | Alternate names ("Fasting Blood Sugar", "Free Thyroxine") |
| `judge_2` | David Kim | Male | Mixed naming ("LDL-C", "HbA1c", "CRP High Sensitivity") |
| `judge_3` | Ashley Brown | Female | Verbose names ("Alanine Aminotransferase", "Glycated Hemoglobin") |

### Matching accuracy across all 8 cases

Run with `python3 eval_matching.py test_cases/<name>.pdf`:

| Case | Extracted | Confident | Uncertain | Wrong | Unmatched | Sex Detected |
|---|---|---|---|---|---|---|
| normal_1 | 38 | 29 (76%) | 9 | 0 | 0 | Male ✓ |
| normal_2 | 37 | 29 (78%) | 8 | 0 | 0 | Female ✓ |
| edge_1 | 30 | 25 (83%) | 5 | 0 | 0 | Male ✓ |
| edge_2 | 27 | 23 (85%) | 4 | 0 | 0 | Female ✓ |
| edge_3 | 27 | 23 (85%) | 4 | 0 | 0 | Male ✓ |
| judge_1 | 21 | 18 (86%) | 3 | 0 | 0 | Female ✓ |
| judge_2 | 17 | 13 (76%) | 4 | 0 | 0 | Male ✓ |
| judge_3 | 23 | 19 (83%) | 3 | 1 | 0 | Female ✓ |

**Sex detection: 8/8 correct across all cases.**

### Gender-specific range validation

All gender-specific biomarkers were classified using the correct range for the detected sex:

| Case | Biomarker | Value | Range Applied | Status |
|---|---|---|---|---|
| normal_1 (M) | RBC | 5.1 | Male 4.8–5.5 | In Range ✓ |
| normal_1 (M) | HGB | 14.8 | Male 14.0–15.0 | In Range ✓ |
| normal_2 (F) | RBC | 4.5 | Female 4.3–4.8 | In Range ✓ |
| normal_2 (F) | DHEA-S | 310 | Female 275–390 | In Range ✓ |
| edge_1 (M) | DHEA-S | 95 | Male 350–690 | Out of Range ✓ |
| edge_2 (F) | HCT | 33.5 | Female 37–44 | Out of Range ✓ |
| edge_3 (M) | HGB | 9.5 | Male 14.0–15.0 | Out of Range ✓ |
| judge_1 (F) | HGB | 13.5 | Female 13.5–14.5 | In Range ✓ |

### Recurring uncertain flags

The same patterns appear as uncertain across nearly all cases — all are legitimate data quality issues in the reference Excel or genuine methodological ambiguities, not model errors:

- **Neutrophils/Lymphs/Monocytes/Eos/Baso (Absolute) → reference names** — the reference sheet does not clarify absolute count vs. percentage differential, so the judge flags these as ambiguous
- **Creatinine → CREATINE** — typo in the reference sheet (creatinine ≠ creatine)
- **Lymphs (Absolute) → LYMPHOHICYTES** — misspelling of LYMPHOCYTES in reference sheet
- **LDL Chol Calc (NIH) → LDL DIRECT** — calculated vs. directly measured LDL are methodologically different
- **Insulin → INSULIN, FASTING** — fasting status not specified in the PDF label

The confident/uncertain split on borderline cases varies slightly across runs due to LLM stochasticity. Zero wrong matches and zero unmatched are consistent across all runs.

In `judge_3`, "Serum Creatinine → CREATINE" was correctly escalated from uncertain to **wrong** because the explicit word "Creatinine" in the PDF name made the reference sheet typo undeniable. This demonstrates the judge working as intended.

### What worked

- Vision extraction is robust across PDF types — including vector PDFs where text parsers fail.
- Patient sex is read directly from the document header with 100% accuracy across all test cases.
- Semantic name matching handles the full range of Labcorp naming variations with high confidence, including highly verbose names ("Alanine Aminotransferase" → ALT) and alternate forms ("Fasting Blood Sugar" → GLUCOSE).
- The LLM-as-judge step catches wrong matches and escalates ambiguous ones for human review rather than silently including them.
- Gender-specific ranges are applied correctly in all cases.

### What failed and where a human should stay involved

- **Uncertain matches require manual review.** The app surfaces them but cannot resolve them automatically. A coach should review uncertain matches before acting on the report.
- **The reference sheet has real typos** (CREATINE, LYMPHOHICYTES) that cause recurring uncertain flags. These should be fixed in the Excel for a production deployment.
- **Units are not converted.** If a lab reports a value in different units than the reference sheet expects, the comparison will be wrong with no warning.
- **LLM stochasticity.** The matching and validation steps can produce slightly different confidence labels across runs for borderline cases.
- **Very low-resolution or damaged scans** may cause Vision to miss or misread values.

---

## Artifact Snapshot

The app extracts biomarkers, detects patient sex, validates name matches, generates a downloadable PDF report, and displays a coach summary — all from a single file upload.

**App UI output example (edge case — metabolic syndrome pattern, Male):**

- Detected: Patient sex **Male** — gender-specific ranges applied for RBC, HGB, HCT, DHEA-S
- Coach Summary: *"Across multiple panels, this client shows an emerging pattern consistent with early metabolic dysregulation. Fasting insulin at 18.5 and triglycerides at 285 — both above optimal — suggest reduced insulin sensitivity and possible early carbohydrate or fat metabolism stress. DHEA-S at 95 falls well below the male optimal range, potentially indicating adrenal insufficiency or chronic stress load..."*
- Status breakdown: In Range / Slightly Off / Out of Range
- Match validation: 25 Confident / 5 Uncertain / 0 Wrong / 0 Unmatched

**Color-coded PDF report** (sample rows):

| Biomarker | Result | Optimal Range | Status |
|---|---|---|---|
| GLUCOSE | 118 mg/dL | 75–86 mg/dL | Out of Range |
| DHEA-S | 95 ug/dL | Male: 350–690 ug/dL | Out of Range |
| INSULIN, FASTING | 18.5 uIU/mL | 2–5 uIU/mL | Out of Range |
| TRIGLYCERIDES | 285 mg/dL | 70–110 mg/dL | Out of Range |

---

## Setup and Usage

### Requirements

- Python 3.9+
- An Anthropic API key

### Installation

```bash
git clone <repo-url>
cd Final_Project_JHU
pip install -r requirements.txt
```

### API key

Create a `.env` file in the project root:

```
ANTHROPIC_API_KEY=your_key_here
```

Do not commit this file. It is listed in `.gitignore`.

### Run the app

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser. Upload any lab results file (PDF, .docx, or .xlsx) and click **Generate Report**.

### Run the matching eval

```bash
python3 eval_matching.py test_cases/normal_1.pdf
```

You can pass any test document as an argument:

```bash
python3 eval_matching.py test_cases/edge_1.pdf
python3 eval_matching.py test_cases/judge_3.pdf
```

### Project structure

```
Final_Project_JHU/
├── app.py                        # Streamlit app
├── eval_matching.py              # Matching + gender eval script
├── create_test_documents.py      # Generates all 24 test documents
├── requirements.txt
├── test_cases/                   # 24 test documents (8 cases × 3 formats)
├── .agents/
│   └── skills/
│       └── lab-report-analyzer/
│           ├── SKILL.md
│           ├── scripts/
│           │   └── analyze_labs.py   # Core skill: range comparison + PDF generation
│           └── references/
│               └── Optimal Lab Ranges.xlsx
└── .env                          # Not committed — add your API key here
```

---

*This tool is for health coaching purposes only and is not a medical diagnosis. Optimal ranges reflect the coach's reference sheet, not standard clinical ranges.*
