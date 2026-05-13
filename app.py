"""
Health Coach Lab Report Analyzer
Streamlit app — upload any lab document, get a validated color-coded report.
"""

import base64
import json
import os
import sys
import tempfile
from pathlib import Path

import anthropic
import fitz  # PyMuPDF
import streamlit as st
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env", override=True)

SKILL_DIR = Path(__file__).parent / ".agents" / "skills" / "lab-report-analyzer"
REFERENCE_EXCEL = SKILL_DIR / "references" / "Optimal Lab Ranges.xlsx"

sys.path.insert(0, str(SKILL_DIR / "scripts"))
from analyze_labs import load_reference_excel, generate_report, determine_status, get_findings


def get_client():
    return anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])


# ---------------------------------------------------------------------------
# Step 1 — Convert document to images
# ---------------------------------------------------------------------------

def doc_to_images(file_path: str) -> list[str]:
    """Converts each page of a PDF to a base64 PNG."""
    doc = fitz.open(file_path)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        images.append(base64.standard_b64encode(pix.tobytes("png")).decode("utf-8"))
    doc.close()
    return images


# ---------------------------------------------------------------------------
# Step 2 — Claude Vision: Extract biomarkers
# ---------------------------------------------------------------------------

def extract_biomarkers(images: list[str]) -> tuple:
    """
    Claude reads the lab document and extracts all biomarkers organized by panel.
    Returns panels with original PDF/document names preserved.

    Format:
    {
        "Panel Name": [
            {"name": "original name from doc", "value": 84.0,
             "flag": "Low"|"High"|null, "lab_low": 70.0, "lab_high": 99.0, "units": "mg/dL"}
        ]
    }
    """
    content = []
    for img in images:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img}
        })
    content.append({
        "type": "text",
        "text": """Extract all biomarker results from this lab report, and also identify the patient's biological sex if shown (e.g. Sex: Male, Gender: F, M/F field).

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
- Use null for lab_low or lab_high if not shown or one-sided range
- Skip non-numeric results (e.g. "Will Follow", "Pending")
- Return ONLY the JSON object, no markdown, no explanation."""
    })

    message = get_client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system="You are a clinical lab data extraction specialist. Extract biomarker data exactly as it appears in lab documents.",
        messages=[{"role": "user", "content": content}],
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    parsed = json.loads(raw.strip())
    sex    = parsed.get("sex")
    panels = parsed.get("panels", parsed)  # fallback if model returns old format

    for bms in panels.values():
        for bm in bms:
            bm["value"]    = float(bm.get("value") or 0)
            bm["lab_low"]  = float(bm["lab_low"])  if bm.get("lab_low")  is not None else None
            bm["lab_high"] = float(bm["lab_high"]) if bm.get("lab_high") is not None else None
            bm["flag"]     = bm.get("flag") or None
            bm["units"]    = bm.get("units", "")

    return panels, sex


# ---------------------------------------------------------------------------
# Step 3 — Claude Matching: Match PDF names to reference sheet names
# ---------------------------------------------------------------------------

def match_biomarkers(pdf_names: list[str], ref_names: list[str]) -> dict:
    """
    Claude matches PDF biomarker names to reference Excel names by biological meaning.
    Returns: {"PDF name": "REFERENCE NAME", ...}
    """
    message = get_client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""You are a clinical lab data specialist helping a functional medicine health coach.
Match biomarker names from a lab document to names in the coach's reference Excel sheet.

Common naming differences:
- Abbreviations: "AST (SGOT)" → "AST", "ALT (SGPT)" → "ALT"
- Short forms: "Hemoglobin" → "HGB", "Hematocrit" → "HCT", "Platelets" → "PLT"
- Qualifiers: "Neutrophils (Absolute)" → "NEUTROPHILS", "Eos (Absolute)" → "EOSINOPHILS"
- Alternate names: "Carbon Dioxide, Total" → "CO2/BICARBONATE/CARBON DIOXIDE"
- Insulin: "Insulin" → "INSULIN, FASTING"
- CRP: "C-Reactive Protein, Quant" → "CRP"

Rules:
- Match on biological meaning, not string similarity
- Only use exact reference names from the provided list
- Skip any PDF name with no confident match
- Return ONLY a JSON object""",
        messages=[{"role": "user", "content": f"""Match each lab document biomarker to the correct reference sheet name.

Lab document biomarkers:
{json.dumps(pdf_names, indent=2)}

Reference sheet names:
{json.dumps(ref_names, indent=2)}

Return: {{"PDF name": "REFERENCE NAME", ...}}
Only include confident matches. Return ONLY the JSON."""}]
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ---------------------------------------------------------------------------
# Step 4 — Claude Validation: LLM-as-judge
# ---------------------------------------------------------------------------

def validate_matches(mapping: dict, ref_dict: dict) -> dict:
    """
    Claude validates its own matches.
    Returns:
    {
        "PDF name": {
            "ref_name": "REFERENCE NAME",
            "confidence": "confident"|"uncertain"|"wrong",
            "reason": "explanation if not confident"
        }
    }
    """
    pairs = [
        {"pdf_name": pdf, "ref_name": ref, "ref_has_range": bool(ref_dict.get(ref, {}).get("optimal_range"))}
        for pdf, ref in mapping.items()
    ]

    message = get_client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system="""You are a clinical lab quality control specialist.
Review biomarker name matches between a lab document and a reference sheet.
For each match, assess whether it is correct.

Confidence levels:
- "confident": clearly the same biomarker, name difference is just formatting/abbreviation
- "uncertain": possibly the same but there is ambiguity (different test methods, similar but not identical)
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

Return ONLY the JSON."""}]
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ---------------------------------------------------------------------------
# Step 5 — Coach Summary
# ---------------------------------------------------------------------------

def generate_coach_summary(panels: dict, reference: dict, mapping: dict, sex: str = None) -> str:
    """Plain-English summary of flagged biomarkers for the health coach."""
    flagged = []
    for panel_name, biomarkers in panels.items():
        for bm in biomarkers:
            ref_name = mapping.get(bm["name"])
            ref = reference.get(ref_name) if ref_name else None
            if not ref:
                continue
            status = determine_status(bm, ref, sex)
            if status != "In Range":
                flagged.append({
                    "panel": panel_name,
                    "name": bm["name"],
                    "value": bm["value"],
                    "units": bm["units"],
                    "optimal_range": ref["optimal_range"],
                    "status": status
                })

    if not flagged:
        return "All measured biomarkers fall within the health coach's optimal reference ranges."

    message = get_client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=512,
        system="""You are a functional medicine health coach assistant.
Write clear, concise summaries for health coaches — not patients.
Focus on patterns and priorities. Never diagnose or prescribe.""",
        messages=[{"role": "user", "content": f"""Write a 3-5 sentence summary for the health coach based on these flagged biomarkers.
Highlight the most significant findings, note patterns across panels, and flag priority concerns.

{json.dumps(flagged, indent=2)}

Under 100 words. Paragraph form, no bullet points."""}]
    )
    return message.content[0].text.strip()


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Lab Report Analyzer", page_icon="🧬", layout="centered")

st.title("Lab Report Analyzer")
st.caption(
    "Upload a patient lab report. The app extracts biomarkers, validates matches "
    "against the health coach's reference sheet, and returns a color-coded report."
)
st.markdown("---")

uploaded_file = st.file_uploader("Upload lab results (PDF, Word, or Excel)", type=["pdf", "docx", "xlsx"])

if uploaded_file:
    st.success(f"Uploaded: **{uploaded_file.name}**")

    if st.button("Generate Report", type="primary"):

        suffix = Path(uploaded_file.name).suffix.lower()
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        # Step 1 — Convert to images
        with st.spinner("Converting document to images..."):
            images = doc_to_images(tmp_path)

        # Step 2 — Extract biomarkers
        with st.spinner(f"Extracting biomarkers from {len(images)} page(s)..."):
            reference = load_reference_excel(REFERENCE_EXCEL)
            try:
                panels, patient_sex = extract_biomarkers(images)
            except Exception as e:
                st.error(f"Extraction error: {e}")
                st.stop()

        pdf_names = [bm["name"] for bms in panels.values() for bm in bms]
        total = len(pdf_names)
        sex_label = f" | Patient sex: **{patient_sex.capitalize()}**" if patient_sex else " | Patient sex: not detected"
        st.info(f"Extracted **{total} biomarkers** across **{len(panels)} panels**.{sex_label}")

        # Step 3 — Match to reference
        with st.spinner("Matching biomarkers to reference sheet..."):
            ref_names = list(reference.keys())
            try:
                mapping = match_biomarkers(pdf_names, ref_names)
            except Exception as e:
                st.error(f"Matching error: {e}")
                st.stop()

        # Step 4 — Validate matches
        with st.spinner("Validating matches..."):
            try:
                validation = validate_matches(mapping, reference)
            except Exception as e:
                st.error(f"Validation error: {e}")
                st.stop()

        # Split by confidence
        confident = {pdf: v["ref_name"] for pdf, v in validation.items() if v["confidence"] == "confident"}
        uncertain = {pdf: v for pdf, v in validation.items() if v["confidence"] == "uncertain"}
        wrong     = {pdf: v for pdf, v in validation.items() if v["confidence"] == "wrong"}
        unmatched = [n for n in pdf_names if n not in mapping]

        # Step 5 — Generate report using only confident matches
        with st.spinner("Generating color-coded report..."):
            # Remap panel biomarker names to reference names for confident matches only
            report_panels = {}
            for panel_name, biomarkers in panels.items():
                rows = []
                for bm in biomarkers:
                    ref_name = confident.get(bm["name"])
                    if ref_name:
                        rows.append({**bm, "name": ref_name})
                if rows:
                    report_panels[panel_name] = rows

            identity_mapping = {bm["name"]: bm["name"] for bms in report_panels.values() for bm in bms}

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_out:
                output_path = tmp_out.name
            generate_report(report_panels, reference, identity_mapping, output_path, sex=patient_sex)

            with open(output_path, "rb") as f:
                report_bytes = f.read()

        # Step 6 — Coach summary
        with st.spinner("Generating coach summary..."):
            try:
                summary = generate_coach_summary(panels, reference, confident, sex=patient_sex)
            except Exception:
                summary = None

        # ── UI Output ──────────────────────────────────────────────────────

        # Coach summary
        if summary:
            st.markdown("---")
            st.subheader("Coach Summary")
            st.info(summary)

        # Status breakdown
        st.markdown("---")
        in_range = slightly_off = out_of_range = 0
        for bms in report_panels.values():
            for bm in bms:
                ref = reference.get(bm["name"])
                if ref:
                    s = determine_status(bm, ref, patient_sex)
                    if s == "In Range":       in_range += 1
                    elif s == "Slightly Off": slightly_off += 1
                    else:                     out_of_range += 1

        col1, col2, col3 = st.columns(3)
        col1.metric("In Range", in_range)
        col2.metric("Slightly Off", slightly_off)
        col3.metric("Out of Range", out_of_range)

        # Download
        st.markdown("---")
        st.download_button(
            label="Download Full Report PDF",
            data=report_bytes,
            file_name="lab_report_analysis.pdf",
            mime="application/pdf",
        )

        # Match validation details
        st.markdown("---")
        st.subheader("Match Validation")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Confident", len(confident))
        col2.metric("Uncertain", len(uncertain))
        col3.metric("Wrong", len(wrong))
        col4.metric("Unmatched", len(unmatched))

        if uncertain:
            with st.expander(f"⚠️ Uncertain matches ({len(uncertain)}) — review manually"):
                for pdf_name, v in uncertain.items():
                    st.warning(f"**{pdf_name}** → {v['ref_name']} | {v.get('reason', '')}")

        if wrong:
            with st.expander(f"❌ Wrong matches ({len(wrong)}) — excluded from report"):
                for pdf_name, v in wrong.items():
                    st.error(f"**{pdf_name}** → {v['ref_name']} | {v.get('reason', '')}")

        if unmatched:
            with st.expander(f"🔍 Unmatched biomarkers ({len(unmatched)}) — not in reference sheet"):
                for name in unmatched:
                    st.write(f"- {name}")

        with st.expander(f"✅ Confident matches ({len(confident)})"):
            for pdf_name, ref_name in confident.items():
                label = f"**{pdf_name}**" if pdf_name == ref_name else f"**{pdf_name}** → {ref_name}"
                st.write(f"- {label}")

st.markdown("---")
st.caption(
    "This tool is for health coaching purposes only and is not a medical diagnosis. "
    "Optimal ranges reflect the coach's reference sheet, not standard clinical ranges."
)
