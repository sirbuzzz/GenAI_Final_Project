"""
eval_matching.py

Evaluates biomarker extraction + matching accuracy for a test document.
Runs Claude Vision extraction, name matching, and LLM-as-judge validation,
then prints a table showing every biomarker and its match result.

Usage:
    python eval_matching.py [path/to/test_doc.pdf]

Default: test_cases/normal_case.pdf
"""

import base64
import json
import os
import sys
from pathlib import Path

import fitz  # PyMuPDF
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env", override=True)

sys.path.insert(0, str(Path(__file__).parent / ".agents" / "skills" / "lab-report-analyzer" / "scripts"))
from analyze_labs import load_reference_excel, determine_status, _get_optimal_bounds

import anthropic

REFERENCE_EXCEL = Path(__file__).parent / ".agents" / "skills" / "lab-report-analyzer" / "references" / "Optimal Lab Ranges.xlsx"

CLIENT = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])


def doc_to_images(file_path: str) -> list[str]:
    doc = fitz.open(file_path)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        images.append(base64.standard_b64encode(pix.tobytes("png")).decode("utf-8"))
    doc.close()
    return images


def extract_biomarkers(images: list[str]) -> tuple:
    content = []
    for img in images:
        content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img}})
    content.append({"type": "text", "text": """Extract all biomarker results and the patient's biological sex from this lab report.

Return a JSON object with two keys: "sex" and "panels".

Format:
{
  "sex": "male" or "female" or null,
  "panels": {
    "Panel Name": [
      {"name": "exact name from document", "value": 84.0, "flag": null, "lab_low": 70.0, "lab_high": 99.0, "units": "mg/dL"}
    ]
  }
}

Rules:
- "sex" should be "male", "female", or null if not shown
- Use null for flag if not flagged
- Use null for lab_low or lab_high if not shown
- Skip non-numeric results
- Return ONLY the JSON object, no markdown, no explanation."""})

    msg = CLIENT.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system="You are a clinical lab data extraction specialist. Extract biomarker data exactly as it appears in lab documents.",
        messages=[{"role": "user", "content": content}],
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    parsed = json.loads(raw.strip())
    sex    = parsed.get("sex")
    panels = parsed.get("panels", parsed)
    return sex, panels


def match_biomarkers(pdf_names: list[str], ref_names: list[str]) -> dict:
    msg = CLIENT.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""You are a clinical lab data specialist. Match biomarker names from a lab document to names in a reference Excel sheet.
Common naming differences:
- Abbreviations: "AST (SGOT)" → "AST", "ALT (SGPT)" → "ALT"
- Short forms: "Hemoglobin" → "HGB", "Hematocrit" → "HCT", "Platelets" → "PLT"
- Qualifiers: "Neutrophils (Absolute)" → "NEUTROPHILS", "Eos (Absolute)" → "EOSINOPHILS"
- Alternate names: "Carbon Dioxide, Total" → "CO2/BICARBONATE/CARBON DIOXIDE"
- Insulin: "Insulin" → "INSULIN, FASTING"
- CRP: "C-Reactive Protein, Quant" → "CRP"
Match on biological meaning. Only use exact reference names. Return ONLY a JSON object.""",
        messages=[{"role": "user", "content": f"""Match each lab document biomarker to the correct reference sheet name.

Lab document biomarkers:
{json.dumps(pdf_names, indent=2)}

Reference sheet names:
{json.dumps(ref_names, indent=2)}

Return: {{"PDF name": "REFERENCE NAME", ...}}
Only include confident matches. Return ONLY the JSON."""}],
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def validate_matches(mapping: dict, ref_dict: dict) -> dict:
    pairs = [
        {"pdf_name": pdf, "ref_name": ref, "ref_has_range": bool(ref_dict.get(ref, {}).get("optimal_range"))}
        for pdf, ref in mapping.items()
    ]
    msg = CLIENT.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""You are a clinical lab quality control specialist.
Review biomarker name matches between a lab document and a reference sheet.
Confidence levels:
- "confident": clearly the same biomarker, name difference is just formatting/abbreviation
- "uncertain": possibly the same but there is ambiguity
- "wrong": clearly different biomarkers
Return ONLY a JSON object.""",
        messages=[{"role": "user", "content": f"""Validate these biomarker matches:

{json.dumps(pairs, indent=2)}

Return:
{{
  "PDF name": {{
    "ref_name": "REFERENCE NAME",
    "confidence": "confident|uncertain|wrong",
    "reason": "brief reason if uncertain or wrong, empty string if confident"
  }}
}}

Return ONLY the JSON."""}],
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run_eval(doc_path: str):
    path = Path(doc_path)
    print(f"\n{'='*70}")
    print(f"  BIOMARKER MATCHING EVAL: {path.name}")
    print(f"{'='*70}\n")

    reference = load_reference_excel(REFERENCE_EXCEL)
    ref_names = list(reference.keys())

    print("Step 1/3  Extracting biomarkers with Claude Vision...")
    images = doc_to_images(str(path))
    sex, panels = extract_biomarkers(images)

    pdf_names = [bm["name"] for bms in panels.values() for bm in bms]
    sex_label = sex.capitalize() if sex else "Not detected"
    print(f"          Extracted {len(pdf_names)} biomarkers across {len(panels)} panels.")
    print(f"          Patient sex detected: {sex_label}\n")

    print("Step 2/3  Matching biomarker names to reference sheet...")
    mapping = match_biomarkers(pdf_names, ref_names)
    print(f"          Matched {len(mapping)} of {len(pdf_names)} biomarkers.\n")

    print("Step 3/3  Validating matches with LLM-as-judge...\n")
    validation = validate_matches(mapping, reference)

    # Categorize
    confident = {p: v for p, v in validation.items() if v["confidence"] == "confident"}
    uncertain = {p: v for p, v in validation.items() if v["confidence"] == "uncertain"}
    wrong     = {p: v for p, v in validation.items() if v["confidence"] == "wrong"}
    unmatched = [n for n in pdf_names if n not in mapping]

    # Print results table
    col_w = [35, 35, 12, 30]
    header = f"{'PDF NAME':<{col_w[0]}} {'REFERENCE MATCH':<{col_w[1]}} {'CONFIDENCE':<{col_w[2]}} {'NOTES'}"
    divider = "-" * (sum(col_w) + 6)

    print(header)
    print(divider)

    for panel_name, biomarkers in panels.items():
        print(f"\n  [{panel_name}]")
        for bm in biomarkers:
            pdf_name = bm["name"]
            if pdf_name in confident:
                v = confident[pdf_name]
                conf_label = "CONFIDENT"
                ref_label = v["ref_name"]
                note = ""
            elif pdf_name in uncertain:
                v = uncertain[pdf_name]
                conf_label = "UNCERTAIN"
                ref_label = v["ref_name"]
                note = v.get("reason", "")[:28]
            elif pdf_name in wrong:
                v = wrong[pdf_name]
                conf_label = "WRONG"
                ref_label = v["ref_name"]
                note = v.get("reason", "")[:28]
            else:
                conf_label = "UNMATCHED"
                ref_label = "(not in reference)"
                note = ""

            print(f"  {pdf_name:<{col_w[0]-2}} {ref_label:<{col_w[1]}} {conf_label:<{col_w[2]}} {note}")

    # Summary
    total = len(pdf_names)
    matched_total = len(mapping)
    print(f"\n{divider}")
    print(f"\nSUMMARY")
    print(f"  Total extracted   : {total}")
    print(f"  Matched           : {matched_total}  ({matched_total/total*100:.0f}%)")
    print(f"  Confident         : {len(confident)}  ({len(confident)/total*100:.0f}%)")
    print(f"  Uncertain         : {len(uncertain)}  ({len(uncertain)/total*100:.0f}%)")
    print(f"  Wrong             : {len(wrong)}  ({len(wrong)/total*100:.0f}%)")
    print(f"  Unmatched         : {len(unmatched)}  ({len(unmatched)/total*100:.0f}%)")
    print(f"\n  Included in report: {len(confident)} biomarkers  ({len(confident)/total*100:.0f}% of extracted)\n")

    if uncertain:
        print("UNCERTAIN — review manually:")
        for pdf_name, v in uncertain.items():
            print(f"  {pdf_name}  →  {v['ref_name']}")
            if v.get("reason"):
                print(f"    Reason: {v['reason']}")

    if wrong:
        print("\nWRONG — excluded from report:")
        for pdf_name, v in wrong.items():
            print(f"  {pdf_name}  →  {v['ref_name']}")
            if v.get("reason"):
                print(f"    Reason: {v['reason']}")

    if unmatched:
        print("\nUNMATCHED — not in reference sheet:")
        for name in unmatched:
            print(f"  {name}")

    # Gender-specific range check
    print(f"\n{divider}")
    print(f"\nGENDER-SPECIFIC RANGES  (detected sex: {sex_label})")
    print(f"{'BIOMARKER':<30} {'VALUE':<10} {'RANGE APPLIED':<20} {'STATUS'}")
    print(f"{'─'*70}")

    all_bms = {bm["name"]: bm for bms in panels.values() for bm in bms}
    gender_found = False
    for pdf_name, v in confident.items():
        ref_name = v["ref_name"]
        ref = reference.get(ref_name)
        if not ref or not ref.get("gender_specific"):
            continue
        bm = all_bms.get(pdf_name)
        if not bm:
            continue
        opt_low, opt_high = _get_optimal_bounds(ref, sex)
        status = determine_status(bm, ref, sex)
        if opt_low is not None:
            range_str = f"{opt_low}-{opt_high}"
        else:
            range_str = "sex unknown"
        print(f"  {pdf_name:<28} {str(bm['value']):<10} {range_str:<20} {status}")
        gender_found = True

    if not gender_found:
        print("  No gender-specific biomarkers found in confident matches.")
    print()


if __name__ == "__main__":
    doc = sys.argv[1] if len(sys.argv) > 1 else "test_cases/normal_1.pdf"
    run_eval(doc)
