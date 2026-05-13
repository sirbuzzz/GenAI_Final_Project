"""
analyze_labs.py

Parses a patient lab results PDF, compares biomarker values against the health coach's
optimal reference ranges from the reference Excel, and generates a color-coded PDF report
organized by lab panel.

Usage:
    python analyze_labs.py <lab_pdf> <mapping_json> <output_pdf>

Arguments:
    lab_pdf       Path to the patient lab results PDF
    mapping_json  Path to JSON file containing model-provided biomarker name mappings
                  Format: {"Lab PDF Name": "Reference Excel Name", ...}
    output_pdf    Path for the generated output PDF report

Dependencies:
    pip install pdfplumber openpyxl reportlab
"""

import sys
import json
import re
from pathlib import Path

import pdfplumber
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable
)
from reportlab.lib.enums import TA_LEFT

REFERENCE_EXCEL = Path(__file__).parent.parent / "references" / "Optimal Lab Ranges.xlsx"

STATUS_GREEN  = "In Range"
STATUS_ORANGE = "Slightly Off"
STATUS_RED    = "Out of Range"


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------

def parse_lab_pdf(pdf_path: str) -> dict:
    """
    Extracts biomarker data from the lab PDF organized by panel.
    Returns:
        {
            "Panel Name": [
                {
                    "name": "Biomarker Name",
                    "value": 84.0,
                    "flag": "Low" | "High" | None,
                    "lab_low": 70.0,
                    "lab_high": 99.0,
                    "units": "mg/dL"
                },
                ...
            ],
            ...
        }
    """
    panels = {}
    current_panel = None

    # Panel headers are section titles — they must NOT contain digits
    panel_header_pattern = re.compile(
        r"^(CBC With Differential[^0-9]*|Comp\. Metabolic[^0-9]*|Lipid Panel[^0-9]*|"
        r"Testosterone[^0-9]*|Hemoglobin A1c[^0-9]*|Thyroxine[^0-9]*|DHEA[^0-9]*|"
        r"TSH[^0-9]*|Estradiol[^0-9]*|GGT[^0-9]*|Insulin[^0-9]*|C-Reactive[^0-9]*|"
        r"Triiodothyronine[^0-9]*|Apolipoprotein[^0-9]*|Cardiovascular[^0-9]*)$",
        re.IGNORECASE
    )

    footnote_pattern = re.compile(r"[⁰-⁹¹²³°]+$")
    non_numeric_skip = {"will follow", "pending", "not estab.", "see note"}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=True)
            lines = _group_words_into_lines(words)

            for line in lines:
                text = line.strip()

                # Detect panel headers
                if panel_header_pattern.match(text):
                    current_panel = text.strip()
                    if current_panel not in panels:
                        panels[current_panel] = []
                    continue

                if current_panel is None:
                    continue

                # Try to parse a biomarker row
                biomarker = _parse_biomarker_row(text, footnote_pattern, non_numeric_skip)
                if biomarker:
                    panels[current_panel].append(biomarker)

    return panels


def _group_words_into_lines(words: list) -> list:
    """Groups extracted words into lines by their vertical position."""
    if not words:
        return []

    lines = []
    current_line = [words[0]]
    current_top = words[0]["top"]

    for word in words[1:]:
        if abs(word["top"] - current_top) < 3:
            current_line.append(word)
        else:
            lines.append(" ".join(w["text"] for w in current_line))
            current_line = [word]
            current_top = word["top"]

    lines.append(" ".join(w["text"] for w in current_line))
    return lines


def _parse_biomarker_row(text: str, footnote_pattern, skip_set: set):
    """
    Attempts to parse a line as a biomarker row.
    Labcorp format: Name | Current Result [Flag] | Previous | Date | Units | Ref Interval
    """
    # Skip header rows and short lines
    if len(text) < 5 or text.lower().startswith("test"):
        return None

    # Skip non-result lines
    for skip in skip_set:
        if skip in text.lower():
            return None

    # Try to extract a numeric value from the line
    tokens = text.split()
    if not tokens:
        return None

    # Find the first numeric token — that's the current result
    value = None
    value_idx = None
    for i, token in enumerate(tokens):
        cleaned = token.replace(",", "")
        try:
            value = float(cleaned)
            value_idx = i
            break
        except ValueError:
            continue

    if value is None or value_idx == 0:
        return None

    # Biomarker name is everything before the first numeric token
    name_tokens = tokens[:value_idx]
    name = " ".join(name_tokens)
    name = footnote_pattern.sub("", name).strip()

    if not name:
        return None

    # Check for flag (Low/High) immediately after the value
    flag = None
    if value_idx + 1 < len(tokens):
        next_token = tokens[value_idx + 1].lower()
        if next_token in ("low", "high"):
            flag = tokens[value_idx + 1].capitalize()

    # Try to find reference interval at the end of the line (e.g. "70-99" or ">59" or "3.4-10.8")
    ref_pattern = re.compile(r"(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)")
    gt_pattern  = re.compile(r"[>≥]\s*(\d+\.?\d*)")
    lt_pattern  = re.compile(r"[<≤]\s*(\d+\.?\d*)")

    lab_low, lab_high = None, None
    ref_match = ref_pattern.search(text)
    gt_match  = gt_pattern.search(text)
    lt_match  = lt_pattern.search(text)

    if ref_match:
        lab_low  = float(ref_match.group(1))
        lab_high = float(ref_match.group(2))
    elif gt_match:
        lab_low  = float(gt_match.group(1))
        lab_high = float("inf")
    elif lt_match:
        lab_low  = float("-inf")
        lab_high = float(lt_match.group(1))

    # Extract units (common patterns)
    units_pattern = re.compile(
        r"\b(mg/dL|g/dL|mmol/L|mEq/L|IU/L|uIU/mL|ulU/mL|ng/dL|pg/mL|"
        r"ug/dL|%|x10E3/uL|x10E6/uL|mL/min/1\.73|mL/min|nmol/L|pmol/L)\b"
    )
    units_match = units_pattern.search(text)
    units = units_match.group(0) if units_match else ""

    return {
        "name": name,
        "value": value,
        "flag": flag,
        "lab_low": lab_low,
        "lab_high": lab_high,
        "units": units,
    }


# ---------------------------------------------------------------------------
# Excel parsing
# ---------------------------------------------------------------------------

def load_reference_excel(excel_path: Path) -> dict:
    """
    Loads the health coach's reference Excel.
    Returns:
        {
            "BIOMARKER NAME": {
                "optimal_range": "70-99 mg/dL",
                "optimal_low": 70.0,
                "optimal_high": 99.0,
                "high_findings": "...",
                "low_findings": "..."
            },
            ...
        }
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    reference = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        name, optimal_range, high_findings, low_findings = (
            row[0], row[1], row[2], row[3]
        )
        if not name or not optimal_range:
            continue

        # Take only the first line, then strip any inline description
        name = str(name).split("\n")[0].strip()
        # Handle "NAME : description" (space-colon-space)
        if " : " in name:
            name = name.split(" : ")[0].strip()
        # Handle "NAME = description" (equals sign separator)
        elif " = " in name:
            name = name.split(" = ")[0].strip()
        # Handle "NAME: long description" (colon-space where description is long)
        elif ": " in name:
            parts = name.split(": ", 1)
            if len(parts[1]) > 15:
                name = parts[0].strip()
        name = name.upper()
        optimal_range = str(optimal_range).strip()

        # Parse numeric bounds — check for gender-specific ranges first
        male_low, male_high, female_low, female_high = _parse_gender_ranges(optimal_range)
        gender_specific = male_low is not None or female_low is not None

        if gender_specific:
            opt_low, opt_high = None, None
        else:
            opt_low, opt_high = _parse_range_string(optimal_range)

        reference[name] = {
            "optimal_range": optimal_range,
            "optimal_low": opt_low,
            "optimal_high": opt_high,
            "male_low": male_low,
            "male_high": male_high,
            "female_low": female_low,
            "female_high": female_high,
            "gender_specific": gender_specific,
            "high_findings": str(high_findings).strip() if high_findings else "",
            "low_findings": str(low_findings).strip() if low_findings else "",
        }

    return reference


def _parse_range_string(range_str: str) -> tuple:
    """Parses a range string like '70-99' or '>50' or '<100' into (low, high)."""
    range_str = range_str.replace("\xa0", " ").strip()

    # Handle "X - Y" or "X to Y" style ranges
    match = re.search(r"(\d+\.?\d*)\s*(?:[-–]|\bto\b)\s*(\d+\.?\d*)", range_str, re.IGNORECASE)
    if match:
        return float(match.group(1)), float(match.group(2))

    gt_match = re.search(r"[>≥]\s*(\d+\.?\d*)", range_str)
    if gt_match:
        return float(gt_match.group(1)), float("inf")

    lt_match = re.search(r"[<≤]\s*(\d+\.?\d*)", range_str)
    if lt_match:
        return float("-inf"), float(lt_match.group(1))

    return None, None


def _parse_gender_ranges(range_str: str):
    """
    Detects and parses gender-specific ranges like:
      'Male 350-690 ... Female 275-390 ...'
      'Male: 40-48%, Female: 37-44%'
      '20-40 for women, 50-70 for men'
    Returns (male_low, male_high, female_low, female_high) or (None, None, None, None).
    """
    s = range_str.replace("\xa0", " ")

    male_low = male_high = female_low = female_high = None

    # Pattern: "Male[: ] <range>" anywhere in string
    m = re.search(r'\bmale[:\s]+(\d+\.?\d*)\s*(?:[-–]|\bto\b)\s*(\d+\.?\d*)', s, re.IGNORECASE)
    if m:
        male_low, male_high = float(m.group(1)), float(m.group(2))

    # Pattern: "Female[: ] <range>"
    f = re.search(r'\bfemale[:\s]+(\d+\.?\d*)\s*(?:[-–]|\bto\b)\s*(\d+\.?\d*)', s, re.IGNORECASE)
    if f:
        female_low, female_high = float(f.group(1)), float(f.group(2))

    # Pattern: "<range> for men/women"
    if male_low is None:
        m2 = re.search(r'(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)\s+for\s+men', s, re.IGNORECASE)
        if m2:
            male_low, male_high = float(m2.group(1)), float(m2.group(2))
    if female_low is None:
        f2 = re.search(r'(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)\s+for\s+wom[ae]n', s, re.IGNORECASE)
        if f2:
            female_low, female_high = float(f2.group(1)), float(f2.group(2))

    if male_low is not None or female_low is not None:
        return male_low, male_high, female_low, female_high
    return None, None, None, None


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def _get_optimal_bounds(ref: dict, sex: str = None):
    """Returns (opt_low, opt_high) using gender-specific bounds when available."""
    if ref.get("gender_specific"):
        sex_lower = (sex or "").lower()
        if sex_lower == "male" and ref.get("male_low") is not None:
            return ref["male_low"], ref["male_high"]
        if sex_lower == "female" and ref.get("female_low") is not None:
            return ref["female_low"], ref["female_high"]
        # Sex unknown — cannot classify gender-specific biomarker
        return None, None
    return ref.get("optimal_low"), ref.get("optimal_high")


def determine_status(biomarker: dict, ref: dict, sex: str = None) -> str:
    """
    Determines the color-coded status for a biomarker.
    Green  — within coach's optimal range
    Orange — outside optimal range but within Labcorp's reference interval
    Red    — outside both ranges
    Pass sex='male' or sex='female' for gender-specific biomarkers.
    """
    value = biomarker["value"]
    opt_low, opt_high = _get_optimal_bounds(ref, sex)
    lab_low  = biomarker.get("lab_low")
    lab_high = biomarker.get("lab_high")

    in_optimal = (
        opt_low is not None and opt_high is not None
        and opt_low <= value <= opt_high
    )
    if in_optimal:
        return STATUS_GREEN

    in_lab_range = (
        lab_low is not None and lab_high is not None
        and lab_low <= value <= lab_high
    )
    if in_lab_range:
        return STATUS_ORANGE

    return STATUS_RED


def get_findings(biomarker: dict, ref: dict, status: str, sex: str = None) -> str:
    """Returns the relevant findings text based on direction of deviation."""
    if status == STATUS_GREEN:
        return ""

    value    = biomarker["value"]
    opt_low, opt_high = _get_optimal_bounds(ref, sex)

    if opt_high is not None and value > opt_high:
        return ref.get("high_findings", "")
    if opt_low is not None and value < opt_low:
        return ref.get("low_findings", "")
    # Gender-specific with unknown sex — show both findings
    if ref.get("gender_specific") and not sex:
        return ref.get("high_findings", "") or ref.get("low_findings", "")
    return ""


# ---------------------------------------------------------------------------
# PDF report generation
# ---------------------------------------------------------------------------

STATUS_COLORS = {
    STATUS_GREEN:  colors.HexColor("#d4edda"),
    STATUS_ORANGE: colors.HexColor("#fff3cd"),
    STATUS_RED:    colors.HexColor("#f8d7da"),
}

STATUS_TEXT_COLORS = {
    STATUS_GREEN:  colors.HexColor("#155724"),
    STATUS_ORANGE: colors.HexColor("#856404"),
    STATUS_RED:    colors.HexColor("#721c24"),
}

HEADER_BG   = colors.HexColor("#343a40")
SUBHEAD_BG  = colors.HexColor("#f8f9fa")
BORDER      = colors.HexColor("#dee2e6")
MUTED       = colors.HexColor("#6c757d")


def generate_report(panels: dict, reference: dict, mapping: dict, output_path: str, sex: str = None) -> None:
    """Generates the color-coded PDF report using reportlab."""

    name_map = {k.lower(): v.upper() for k, v in mapping.items()}

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", fontSize=16, fontName="Helvetica-Bold", spaceAfter=4)
    sub_style   = ParagraphStyle("sub",   fontSize=8,  fontName="Helvetica", textColor=MUTED, spaceAfter=12)
    panel_style = ParagraphStyle("panel", fontSize=11, fontName="Helvetica-Bold", textColor=colors.white)
    cell_style  = ParagraphStyle("cell",  fontSize=7,  fontName="Helvetica", leading=9)
    bold_style  = ParagraphStyle("bold",  fontSize=7,  fontName="Helvetica-Bold", leading=9)

    col_widths = [1.6*inch, 0.9*inch, 1.5*inch, 0.85*inch, 3.0*inch]

    story = []
    story.append(Paragraph("Lab Report Analysis", title_style))
    story.append(Paragraph(
        "Compared against health coach optimal reference ranges. "
        "This report is not a medical diagnosis.", sub_style
    ))

    # Legend
    legend_data = [[
        Paragraph("<font color='#155724'>■</font> In Range", cell_style),
        Paragraph("<font color='#856404'>■</font> Slightly Off", cell_style),
        Paragraph("<font color='#721c24'>■</font> Out of Range", cell_style),
    ]]
    legend_table = Table(legend_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch])
    legend_table.setStyle(TableStyle([("BOX", (0,0), (-1,-1), 0.5, BORDER)]))
    story.append(legend_table)
    story.append(Spacer(1, 14))

    unmatched = []

    for panel_name, biomarkers in panels.items():
        if not biomarkers:
            continue

        rows = []
        row_statuses = []
        for bm in biomarkers:
            bm_key = name_map.get(bm["name"].lower())
            ref = reference.get(bm_key) if bm_key else None

            if ref is None:
                unmatched.append(bm["name"])
                continue

            status   = determine_status(bm, ref, sex)
            findings = get_findings(bm, ref, status, sex)
            tc       = STATUS_TEXT_COLORS[status]

            status_para = Paragraph(
                f"<font color='#{tc.hexval()[2:]}'><b>{status}</b></font>", cell_style
            )

            rows.append([
                Paragraph(bm["name"], cell_style),
                Paragraph(f"{bm['value']} {bm['units']}", cell_style),
                Paragraph(ref["optimal_range"], cell_style),
                status_para,
                Paragraph(findings[:300] if findings else "—", cell_style),
            ])
            row_statuses.append(status)

        if not rows:
            continue

        # Panel header
        header_table = Table(
            [[Paragraph(panel_name, panel_style)]],
            colWidths=[sum(col_widths)]
        )
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), HEADER_BG),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(header_table)

        # Column headers
        col_headers = [["Biomarker", "Result", "Optimal Range", "Status", "Findings"]]
        data = col_headers + rows

        # Determine row background colors
        style_cmds = [
            ("BACKGROUND", (0,0), (-1,0), SUBHEAD_BG),
            ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",   (0,0), (-1,-1), 7),
            ("GRID",       (0,0), (-1,-1), 0.5, BORDER),
            ("VALIGN",     (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING",(0,0), (-1,-1), 5),
            ("RIGHTPADDING",(0,0),(-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ]

        for row_idx, status in enumerate(row_statuses, start=1):
            style_cmds.append(("BACKGROUND", (3, row_idx), (3, row_idx), STATUS_COLORS[status]))

        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle(style_cmds))
        story.append(table)
        story.append(Spacer(1, 14))

    # Unmatched section
    if unmatched:
        header_table = Table(
            [[Paragraph("Unmatched Biomarkers", panel_style)]],
            colWidths=[sum(col_widths)]
        )
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), MUTED),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(header_table)

        unmatched_data = [["Biomarker", "Note"]] + [
            [name, "Not found in reference sheet"] for name in unmatched
        ]
        u_table = Table(unmatched_data, colWidths=[2.5*inch, 5.35*inch])
        u_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), SUBHEAD_BG),
            ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",   (0,0), (-1,-1), 7),
            ("GRID",       (0,0), (-1,-1), 0.5, BORDER),
            ("TEXTCOLOR",  (1,1), (1,-1), MUTED),
        ]))
        story.append(u_table)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
    )
    doc.build(story)
    print(f"Report saved to: {output_path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 4:
        print("Usage: python analyze_labs.py <lab_pdf> <mapping_json> <output_pdf>")
        sys.exit(1)

    lab_pdf_path   = sys.argv[1]
    mapping_path   = sys.argv[2]
    output_pdf     = sys.argv[3]

    print("Parsing lab PDF...")
    panels = parse_lab_pdf(lab_pdf_path)

    print("Loading reference Excel...")
    reference = load_reference_excel(REFERENCE_EXCEL)

    print("Loading biomarker name mapping...")
    with open(mapping_path, "r") as f:
        mapping = json.load(f)

    print("Generating report...")
    generate_report(panels, reference, mapping, output_pdf)


if __name__ == "__main__":
    main()
