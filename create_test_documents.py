"""
Creates synthetic Labcorp-style lab result documents (PDF, DOCX, XLSX).

Cases:
  normal_1  — Male,   35 — Clean routine checkup
  normal_2  — Female, 42 — Clean routine checkup, female panels
  edge_1    — Male,   58 — Metabolic syndrome pattern
  edge_2    — Female, 45 — Thyroid/hormonal dysfunction + iron-deficiency anemia
  edge_3    — Male,   67 — Critical: kidney failure, severe anemia, diabetes
  judge_1   — Female, 38 — Alternate biomarker names (Free Thyroxine, Fasting Blood Sugar...)
  judge_2   — Male,   50 — Mixed naming (LDL-C, HbA1c, CRP High Sensitivity...)
  judge_3   — Female, 29 — Verbose names (Alanine Aminotransferase, Total White Blood Cells...)
"""

from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

Path("test_cases").mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Dataset definitions
# Each panel row: (name, value, flag, units, reference_interval)
# ---------------------------------------------------------------------------

NORMAL_1 = {
    "label": "Routine Health Panel",
    "patient": "James Carter",
    "dob": "03/14/1990",
    "sex": "Male",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential/Platelet", [
            ("WBC",                    "5.5",  "",     "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "5.1",  "",     "x10E6/uL",    "4.14-5.80"),
            ("Hemoglobin",             "14.8", "",     "g/dL",        "13.0-17.7"),
            ("Hematocrit",             "44.0", "",     "%",           "37.5-51.0"),
            ("MCV",                    "88",   "",     "fL",          "79-97"),
            ("MCH",                    "29.0", "",     "pg",          "26.6-33.0"),
            ("MCHC",                   "33.2", "",     "g/dL",        "31.5-35.7"),
            ("RDW",                    "12.8", "",     "%",           "11.6-15.4"),
            ("Platelets",              "245",  "",     "x10E3/uL",    "150-450"),
            ("Neutrophils (Absolute)", "3.2",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "1.9",  "",     "x10E3/uL",    "0.7-3.1"),
            ("Monocytes(Absolute)",    "0.4",  "",     "x10E3/uL",    "0.1-0.9"),
            ("Eos (Absolute)",         "0.2",  "",     "x10E3/uL",    "0.0-0.4"),
            ("Baso (Absolute)",        "0.0",  "",     "x10E3/uL",    "0.0-0.2"),
            ("Immature Grans (Abs)",   "0.0",  "",     "x10E3/uL",    "0.0-0.1"),
        ]),
        ("Comp. Metabolic Panel (14)", [
            ("Glucose",                "82",   "",     "mg/dL",       "70-99"),
            ("BUN",                    "13",   "",     "mg/dL",       "6-20"),
            ("Creatinine",             "0.96", "",     "mg/dL",       "0.76-1.27"),
            ("eGFR",                   "105",  "",     "mL/min/1.73", ">59"),
            ("BUN/Creatinine Ratio",   "14",   "",     "",            "9-20"),
            ("Sodium",                 "140",  "",     "mmol/L",      "134-144"),
            ("Potassium",              "4.3",  "",     "mmol/L",      "3.5-5.2"),
            ("Chloride",               "103",  "",     "mmol/L",      "96-106"),
            ("Carbon Dioxide, Total",  "26",   "",     "mmol/L",      "20-29"),
            ("Calcium",                "9.3",  "",     "mg/dL",       "8.7-10.2"),
            ("Protein, Total",         "7.4",  "",     "g/dL",        "6.0-8.5"),
            ("Albumin",                "4.7",  "",     "g/dL",        "4.1-5.1"),
            ("Globulin, Total",        "2.7",  "",     "g/dL",        "1.5-4.5"),
            ("AST (SGOT)",             "22",   "",     "IU/L",        "0-40"),
            ("ALT (SGPT)",             "24",   "",     "IU/L",        "0-44"),
        ]),
        ("Lipid Panel", [
            ("Cholesterol, Total",     "182",  "",     "mg/dL",       "100-199"),
            ("Triglycerides",          "85",   "",     "mg/dL",       "0-149"),
            ("HDL Cholesterol",        "58",   "",     "mg/dL",       ">39"),
            ("VLDL Cholesterol Cal",   "17",   "",     "mg/dL",       "5-40"),
            ("LDL Chol Calc (NIH)",    "107",  "High", "mg/dL",       "0-99"),
        ]),
        ("Hemoglobin A1c", [
            ("Hemoglobin A1c",         "5.3",  "",     "%",           "4.8-5.6"),
        ]),
        ("TSH", [
            ("TSH",                    "1.8",  "",     "uIU/mL",      "0.450-4.500"),
        ]),
        ("Insulin", [
            ("Insulin",                "4.5",  "",     "uIU/mL",      "2.6-24.9"),
        ]),
    ]
}

NORMAL_2 = {
    "label": "Routine Health Panel",
    "patient": "Sarah Mitchell",
    "dob": "07/22/1983",
    "sex": "Female",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential/Platelet", [
            ("WBC",                    "5.8",  "",     "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "4.5",  "",     "x10E6/uL",    "3.77-5.28"),
            ("Hemoglobin",             "13.8", "",     "g/dL",        "11.1-15.9"),
            ("Hematocrit",             "41.0", "",     "%",           "34.0-46.6"),
            ("MCV",                    "87",   "",     "fL",          "79-97"),
            ("MCH",                    "30.2", "",     "pg",          "26.6-33.0"),
            ("MCHC",                   "33.6", "",     "g/dL",        "31.5-35.7"),
            ("RDW",                    "13.1", "",     "%",           "11.6-15.4"),
            ("Platelets",              "268",  "",     "x10E3/uL",    "150-450"),
            ("Neutrophils (Absolute)", "3.5",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "1.7",  "",     "x10E3/uL",    "0.7-3.1"),
            ("Monocytes(Absolute)",    "0.3",  "",     "x10E3/uL",    "0.1-0.9"),
            ("Eos (Absolute)",         "0.2",  "",     "x10E3/uL",    "0.0-0.4"),
            ("Baso (Absolute)",        "0.0",  "",     "x10E3/uL",    "0.0-0.2"),
            ("Immature Grans (Abs)",   "0.0",  "",     "x10E3/uL",    "0.0-0.1"),
        ]),
        ("Comp. Metabolic Panel (14)", [
            ("Glucose",                "84",   "",     "mg/dL",       "70-99"),
            ("BUN",                    "11",   "",     "mg/dL",       "6-20"),
            ("Creatinine",             "0.82", "",     "mg/dL",       "0.57-1.00"),
            ("eGFR",                   "98",   "",     "mL/min/1.73", ">59"),
            ("Sodium",                 "139",  "",     "mmol/L",      "134-144"),
            ("Potassium",              "4.1",  "",     "mmol/L",      "3.5-5.2"),
            ("Chloride",               "101",  "",     "mmol/L",      "96-106"),
            ("Carbon Dioxide, Total",  "25",   "",     "mmol/L",      "20-29"),
            ("Calcium",                "9.1",  "",     "mg/dL",       "8.7-10.2"),
            ("Protein, Total",         "7.2",  "",     "g/dL",        "6.0-8.5"),
            ("Albumin",                "4.6",  "",     "g/dL",        "4.1-5.1"),
            ("AST (SGOT)",             "18",   "",     "IU/L",        "0-40"),
            ("ALT (SGPT)",             "16",   "",     "IU/L",        "0-44"),
        ]),
        ("Lipid Panel", [
            ("Cholesterol, Total",     "175",  "",     "mg/dL",       "100-199"),
            ("Triglycerides",          "78",   "",     "mg/dL",       "0-149"),
            ("HDL Cholesterol",        "65",   "",     "mg/dL",       ">39"),
            ("VLDL Cholesterol Cal",   "16",   "",     "mg/dL",       "5-40"),
            ("LDL Chol Calc (NIH)",    "94",   "",     "mg/dL",       "0-99"),
        ]),
        ("Hemoglobin A1c", [
            ("Hemoglobin A1c",         "5.1",  "",     "%",           "4.8-5.6"),
        ]),
        ("TSH", [
            ("TSH",                    "2.1",  "",     "uIU/mL",      "0.450-4.500"),
        ]),
        ("Insulin", [
            ("Insulin",                "3.8",  "",     "uIU/mL",      "2.6-24.9"),
        ]),
        ("DHEA-Sulfate", [
            ("DHEA-Sulfate",           "310.0","",     "ug/dL",       "45.0-270.0"),
        ]),
    ]
}

EDGE_1 = {
    "label": "Comprehensive Metabolic Panel",
    "patient": "Robert Davis",
    "dob": "09/05/1967",
    "sex": "Male",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential/Platelet", [
            ("WBC",                    "7.8",  "",     "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "4.6",  "",     "x10E6/uL",    "4.14-5.80"),
            ("Hemoglobin",             "13.8", "",     "g/dL",        "13.0-17.7"),
            ("Hematocrit",             "41.5", "",     "%",           "37.5-51.0"),
            ("MCV",                    "86",   "",     "fL",          "79-97"),
            ("Platelets",              "310",  "",     "x10E3/uL",    "150-450"),
            ("Neutrophils (Absolute)", "5.2",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "1.8",  "",     "x10E3/uL",    "0.7-3.1"),
        ]),
        ("Comp. Metabolic Panel (14)", [
            ("Glucose",                "118",  "High", "mg/dL",       "70-99"),
            ("BUN",                    "18",   "",     "mg/dL",       "6-20"),
            ("Creatinine",             "1.15", "",     "mg/dL",       "0.76-1.27"),
            ("eGFR",                   "72",   "",     "mL/min/1.73", ">59"),
            ("Sodium",                 "140",  "",     "mmol/L",      "134-144"),
            ("Potassium",              "4.5",  "",     "mmol/L",      "3.5-5.2"),
            ("Chloride",               "102",  "",     "mmol/L",      "96-106"),
            ("Carbon Dioxide, Total",  "24",   "",     "mmol/L",      "20-29"),
            ("Calcium",                "9.2",  "",     "mg/dL",       "8.7-10.2"),
            ("Albumin",                "4.4",  "",     "g/dL",        "4.1-5.1"),
            ("AST (SGOT)",             "38",   "",     "IU/L",        "0-40"),
            ("ALT (SGPT)",             "52",   "High", "IU/L",        "0-44"),
        ]),
        ("Lipid Panel", [
            ("Cholesterol, Total",     "228",  "High", "mg/dL",       "100-199"),
            ("Triglycerides",          "285",  "High", "mg/dL",       "0-149"),
            ("HDL Cholesterol",        "38",   "Low",  "mg/dL",       ">39"),
            ("VLDL Cholesterol Cal",   "57",   "High", "mg/dL",       "5-40"),
            ("LDL Chol Calc (NIH)",    "133",  "High", "mg/dL",       "0-99"),
        ]),
        ("Hemoglobin A1c", [
            ("Hemoglobin A1c",         "6.2",  "High", "%",           "4.8-5.6"),
        ]),
        ("TSH", [
            ("TSH",                    "2.3",  "",     "uIU/mL",      "0.450-4.500"),
        ]),
        ("Insulin", [
            ("Insulin",                "18.5", "High", "uIU/mL",      "2.6-24.9"),
        ]),
        ("DHEA-Sulfate", [
            ("DHEA-Sulfate",           "95.0", "Low",  "ug/dL",       "138.5-475.2"),
        ]),
        ("C-Reactive Protein, Quant", [
            ("C-Reactive Protein, Quant", "4.8", "High", "mg/L",      "0-10"),
        ]),
    ]
}

EDGE_2 = {
    "label": "Comprehensive Health Panel",
    "patient": "Maria Torres",
    "dob": "11/18/1980",
    "sex": "Female",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential/Platelet", [
            ("WBC",                    "6.2",  "",     "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "3.80", "Low",  "x10E6/uL",    "3.77-5.28"),
            ("Hemoglobin",             "10.8", "Low",  "g/dL",        "11.1-15.9"),
            ("Hematocrit",             "33.5", "Low",  "%",           "34.0-46.6"),
            ("MCV",                    "72",   "Low",  "fL",          "79-97"),
            ("MCH",                    "24.8", "Low",  "pg",          "26.6-33.0"),
            ("MCHC",                   "30.1", "Low",  "g/dL",        "31.5-35.7"),
            ("RDW",                    "15.8", "High", "%",           "11.6-15.4"),
            ("Platelets",              "395",  "",     "x10E3/uL",    "150-450"),
            ("Neutrophils (Absolute)", "3.8",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "1.5",  "",     "x10E3/uL",    "0.7-3.1"),
        ]),
        ("Comp. Metabolic Panel (14)", [
            ("Glucose",                "88",   "",     "mg/dL",       "70-99"),
            ("BUN",                    "12",   "",     "mg/dL",       "6-20"),
            ("Creatinine",             "0.78", "",     "mg/dL",       "0.57-1.00"),
            ("eGFR",                   "95",   "",     "mL/min/1.73", ">59"),
            ("Sodium",                 "138",  "",     "mmol/L",      "134-144"),
            ("Potassium",              "3.8",  "",     "mmol/L",      "3.5-5.2"),
            ("Calcium",                "9.0",  "",     "mg/dL",       "8.7-10.2"),
            ("Albumin",                "4.3",  "",     "g/dL",        "4.1-5.1"),
            ("AST (SGOT)",             "24",   "",     "IU/L",        "0-40"),
            ("ALT (SGPT)",             "20",   "",     "IU/L",        "0-44"),
        ]),
        ("Thyroxine (T4) Free, Direct", [
            ("T4,Free(Direct)",        "2.15", "High", "ng/dL",       "0.82-1.77"),
        ]),
        ("TSH", [
            ("TSH",                    "0.12", "Low",  "uIU/mL",      "0.450-4.500"),
        ]),
        ("Triiodothyronine (T3), Free", [
            ("Triiodothyronine (T3), Free", "4.8", "High", "pg/mL",  "2.0-4.4"),
        ]),
        ("DHEA-Sulfate", [
            ("DHEA-Sulfate",           "58.0", "Low",  "ug/dL",       "45.0-270.0"),
        ]),
        ("Insulin", [
            ("Insulin",                "8.2",  "",     "uIU/mL",      "2.6-24.9"),
        ]),
        ("GGT", [
            ("GGT",                    "22",   "",     "IU/L",        "0-65"),
        ]),
    ]
}

EDGE_3 = {
    "label": "Comprehensive Panel",
    "patient": "William Chen",
    "dob": "02/28/1958",
    "sex": "Male",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential/Platelet", [
            ("WBC",                    "12.8", "High", "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "3.20", "Low",  "x10E6/uL",    "4.14-5.80"),
            ("Hemoglobin",             "9.5",  "Low",  "g/dL",        "13.0-17.7"),
            ("Hematocrit",             "28.8", "Low",  "%",           "37.5-51.0"),
            ("MCV",                    "71",   "Low",  "fL",          "79-97"),
            ("Platelets",              "142",  "Low",  "x10E3/uL",    "150-450"),
            ("Neutrophils (Absolute)", "9.8",  "High", "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "0.8",  "",     "x10E3/uL",    "0.7-3.1"),
        ]),
        ("Comp. Metabolic Panel (14)", [
            ("Glucose",                "248",  "High", "mg/dL",       "70-99"),
            ("BUN",                    "58",   "High", "mg/dL",       "6-20"),
            ("Creatinine",             "4.20", "High", "mg/dL",       "0.76-1.27"),
            ("eGFR",                   "14",   "Low",  "mL/min/1.73", ">59"),
            ("Sodium",                 "132",  "Low",  "mmol/L",      "134-144"),
            ("Potassium",              "5.8",  "High", "mmol/L",      "3.5-5.2"),
            ("Chloride",               "94",   "Low",  "mmol/L",      "96-106"),
            ("Carbon Dioxide, Total",  "16",   "Low",  "mmol/L",      "20-29"),
            ("Calcium",                "7.8",  "Low",  "mg/dL",       "8.7-10.2"),
            ("Albumin",                "2.9",  "Low",  "g/dL",        "4.1-5.1"),
            ("AST (SGOT)",             "88",   "High", "IU/L",        "0-40"),
            ("ALT (SGPT)",             "124",  "High", "IU/L",        "0-44"),
        ]),
        ("Lipid Panel", [
            ("Cholesterol, Total",     "285",  "High", "mg/dL",       "100-199"),
            ("Triglycerides",          "420",  "High", "mg/dL",       "0-149"),
            ("HDL Cholesterol",        "28",   "Low",  "mg/dL",       ">39"),
            ("LDL Chol Calc (NIH)",    "173",  "High", "mg/dL",       "0-99"),
        ]),
        ("Hemoglobin A1c", [
            ("Hemoglobin A1c",         "9.8",  "High", "%",           "4.8-5.6"),
        ]),
        ("Insulin", [
            ("Insulin",                "28.5", "High", "uIU/mL",      "2.6-24.9"),
        ]),
        ("C-Reactive Protein, Quant", [
            ("C-Reactive Protein, Quant", "18.5", "High", "mg/L",     "0-10"),
        ]),
    ]
}

JUDGE_1 = {
    "label": "Annual Lab Panel",
    "patient": "Jennifer Lopez",
    "dob": "05/12/1987",
    "sex": "Female",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("Complete Blood Count", [
            ("White Blood Cell Count",   "6.2",  "",     "x10E3/uL",    "3.4-10.8"),
            ("Red Blood Cell Count",     "4.55", "",     "x10E6/uL",    "3.77-5.28"),
            ("Blood Hemoglobin",         "13.5", "",     "g/dL",        "11.1-15.9"),
            ("Blood Hematocrit",         "40.5", "",     "%",           "34.0-46.6"),
            ("Absolute Neutrophil Count","3.8",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Absolute Lymphocyte Count","1.6",  "",     "x10E3/uL",    "0.7-3.1"),
            ("Platelet Count",           "255",  "",     "x10E3/uL",    "150-450"),
        ]),
        ("Basic Metabolic Panel", [
            ("Fasting Blood Sugar",      "91",   "",     "mg/dL",       "70-99"),
            ("Blood Urea Nitrogen",      "14",   "",     "mg/dL",       "6-20"),
            ("Kidney Creatinine",        "0.85", "",     "mg/dL",       "0.57-1.00"),
            ("Glomerular Filtration Rate","96",  "",     "mL/min/1.73", ">59"),
            ("Serum Sodium",             "139",  "",     "mmol/L",      "134-144"),
            ("Serum Potassium",          "4.0",  "",     "mmol/L",      "3.5-5.2"),
            ("Serum Calcium",            "9.2",  "",     "mg/dL",       "8.7-10.2"),
            ("Liver ALT",                "22",   "",     "IU/L",        "0-44"),
            ("Liver AST",                "20",   "",     "IU/L",        "0-40"),
        ]),
        ("Thyroid Function Tests", [
            ("Thyroid Stimulating Hormone","2.4","",    "uIU/mL",      "0.450-4.500"),
            ("Free Thyroxine",           "1.18", "",     "ng/dL",       "0.82-1.77"),
            ("Free Triiodothyronine",    "3.2",  "",     "pg/mL",       "2.0-4.4"),
        ]),
        ("Metabolic Markers", [
            ("Fasting Insulin Level",    "5.8",  "",     "uIU/mL",      "2.6-24.9"),
            ("Hemoglobin A1c",           "5.4",  "",     "%",           "4.8-5.6"),
        ]),
    ]
}

JUDGE_2 = {
    "label": "Executive Health Panel",
    "patient": "David Kim",
    "dob": "08/30/1975",
    "sex": "Male",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("CBC With Differential", [
            ("WBC",                    "7.1",  "",     "x10E3/uL",    "3.4-10.8"),
            ("RBC",                    "4.95", "",     "x10E6/uL",    "4.14-5.80"),
            ("Hemoglobin",             "14.5", "",     "g/dL",        "13.0-17.7"),
            ("Hematocrit",             "43.8", "",     "%",           "37.5-51.0"),
            ("Neutrophils (Absolute)", "4.2",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Lymphs (Absolute)",      "2.0",  "",     "x10E3/uL",    "0.7-3.1"),
            ("Platelets",              "280",  "",     "x10E3/uL",    "150-450"),
        ]),
        ("Cardiovascular Panel", [
            ("Total Cholesterol",      "205",  "High", "mg/dL",       "100-199"),
            ("LDL-C",                  "128",  "High", "mg/dL",       "0-99"),
            ("HDL-C",                  "45",   "",     "mg/dL",       ">39"),
            ("Serum Triglycerides",    "162",  "High", "mg/dL",       "0-149"),
        ]),
        ("Diabetes Markers", [
            ("HbA1c",                  "5.8",  "High", "%",           "4.8-5.6"),
            ("Fasting Insulin",        "9.2",  "",     "uIU/mL",      "2.6-24.9"),
        ]),
        ("Inflammation", [
            ("CRP High Sensitivity",   "5.2",  "High", "mg/L",        "0-10"),
        ]),
        ("Liver Function", [
            ("ALT (SGPT)",             "35",   "",     "IU/L",        "0-44"),
            ("AST (SGOT)",             "28",   "",     "IU/L",        "0-40"),
        ]),
        ("Hormones", [
            ("DHEA Sulfate",           "185",  "",     "ug/dL",       "138.5-475.2"),
        ]),
    ]
}

JUDGE_3 = {
    "label": "Wellness Lab Panel",
    "patient": "Ashley Brown",
    "dob": "01/15/1996",
    "sex": "Female",
    "collected": "04/23/2026",
    "reported": "04/25/2026",
    "panels": [
        ("Complete Metabolic Panel", [
            ("Serum Glucose Fasting",    "86",   "",     "mg/dL",       "70-99"),
            ("Blood Urea Nitrogen (BUN)","10",   "",     "mg/dL",       "6-20"),
            ("Serum Creatinine",         "0.75", "",     "mg/dL",       "0.57-1.00"),
            ("Estimated GFR",            "108",  "",     "mL/min/1.73", ">59"),
            ("Serum Sodium Level",       "141",  "",     "mmol/L",      "134-144"),
            ("Serum Potassium Level",    "4.2",  "",     "mmol/L",      "3.5-5.2"),
            ("Total Serum Protein",      "7.1",  "",     "g/dL",        "6.0-8.5"),
            ("Serum Albumin",            "4.5",  "",     "g/dL",        "4.1-5.1"),
            ("Alanine Aminotransferase", "18",   "",     "IU/L",        "0-44"),
            ("Aspartate Aminotransferase","16",  "",     "IU/L",        "0-40"),
        ]),
        ("Blood Count Panel", [
            ("Total White Blood Cells",  "5.9",  "",     "x10E3/uL",    "3.4-10.8"),
            ("Total Red Blood Cells",    "4.62", "",     "x10E6/uL",    "3.77-5.28"),
            ("Hemoglobin Concentration", "13.9", "",     "g/dL",        "11.1-15.9"),
            ("Hematocrit Value",         "41.5", "",     "%",           "34.0-46.6"),
            ("Mean Corpuscular Volume",  "86",   "",     "fL",          "79-97"),
            ("Absolute Neutrophil Count","3.4",  "",     "x10E3/uL",    "1.4-7.0"),
            ("Absolute Lymphocyte Count","1.8",  "",     "x10E3/uL",    "0.7-3.1"),
        ]),
        ("Lipid Profile", [
            ("Total Cholesterol Level",  "168",  "",     "mg/dL",       "100-199"),
            ("Triglyceride Level",       "92",   "",     "mg/dL",       "0-149"),
            ("High Density Lipoprotein", "72",   "",     "mg/dL",       ">39"),
            ("Low Density Lipoprotein",  "78",   "",     "mg/dL",       "0-99"),
        ]),
        ("Endocrine Tests", [
            ("Thyroid Stimulating Hormone","1.9","",    "uIU/mL",      "0.450-4.500"),
            ("Glycated Hemoglobin",      "5.0",  "",     "%",           "4.8-5.6"),
        ]),
    ]
}

ALL_CASES = [
    (NORMAL_1, "normal_1"),
    (NORMAL_2, "normal_2"),
    (EDGE_1,   "edge_1"),
    (EDGE_2,   "edge_2"),
    (EDGE_3,   "edge_3"),
    (JUDGE_1,  "judge_1"),
    (JUDGE_2,  "judge_2"),
    (JUDGE_3,  "judge_3"),
]


# ---------------------------------------------------------------------------
# PDF generator — Labcorp-style layout
# ---------------------------------------------------------------------------

def create_pdf(data: dict, filename: str):
    HEADER_BG  = colors.HexColor("#1a1a2e")
    PANEL_BG   = colors.HexColor("#2c3e50")
    SUBHEAD_BG = colors.HexColor("#f0f0f0")
    FLAG_RED   = colors.HexColor("#f8d7da")
    BORDER     = colors.HexColor("#cccccc")
    MUTED      = colors.HexColor("#666666")

    styles = {
        "title":  ParagraphStyle("title",  fontSize=12, fontName="Helvetica-Bold", spaceAfter=2),
        "meta":   ParagraphStyle("meta",   fontSize=8,  fontName="Helvetica", textColor=MUTED, spaceAfter=2),
        "demo":   ParagraphStyle("demo",   fontSize=8,  fontName="Helvetica-Bold", spaceAfter=8),
        "panel":  ParagraphStyle("panel",  fontSize=10, fontName="Helvetica-Bold", textColor=colors.white),
        "cell":   ParagraphStyle("cell",   fontSize=7,  fontName="Helvetica", leading=9),
        "hcell":  ParagraphStyle("hcell",  fontSize=7,  fontName="Helvetica-Bold", leading=9),
        "flag":   ParagraphStyle("flag",   fontSize=7,  fontName="Helvetica-Bold", leading=9,
                                 textColor=colors.HexColor("#c0392b")),
    }

    col_widths = [2.4*inch, 1.0*inch, 0.6*inch, 1.0*inch, 1.4*inch]
    page_width = sum(col_widths)

    story = []

    # Top header bar
    header_data = [[
        Paragraph(f"<font color='white'><b>labcorp</b></font>", styles["title"]),
        Paragraph(
            f"<font color='white'>Date Collected: {data['collected']}   "
            f"Date Reported: {data['reported']}   Fasting: Yes</font>",
            styles["meta"]
        ),
    ]]
    header_table = Table(header_data, colWidths=[1.5*inch, page_width - 1.5*inch])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), HEADER_BG),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 6))

    # Patient demographics
    story.append(Paragraph(
        f"Patient: {data['patient']}     DOB: {data['dob']}     Sex: {data['sex']}",
        styles["demo"]
    ))
    story.append(Paragraph(f"Ordered Items: {data['label']}", styles["meta"]))
    story.append(Spacer(1, 10))

    for panel_name, biomarkers in data["panels"]:
        # Panel header
        panel_row = [[Paragraph(panel_name, styles["panel"])]]
        panel_table = Table(panel_row, colWidths=[page_width])
        panel_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), PANEL_BG),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))
        story.append(panel_table)

        # Column headers
        col_headers = [[
            Paragraph("Test",                      styles["hcell"]),
            Paragraph("Current Result and Flag",   styles["hcell"]),
            Paragraph("Flag",                      styles["hcell"]),
            Paragraph("Units",                     styles["hcell"]),
            Paragraph("Reference Interval",        styles["hcell"]),
        ]]

        rows = []
        flag_rows = []
        for i, (name, value, flag, units, ref) in enumerate(biomarkers):
            value_para = Paragraph(f"<b>{value}</b>" if flag else value, styles["cell"])
            flag_para  = Paragraph(flag, styles["flag"] if flag else styles["cell"])
            rows.append([
                Paragraph(name,  styles["cell"]),
                value_para,
                flag_para,
                Paragraph(units, styles["cell"]),
                Paragraph(ref,   styles["cell"]),
            ])
            if flag in ("Low", "High"):
                flag_rows.append(i + 1)

        table_data = col_headers + rows
        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        style_cmds = [
            ("BACKGROUND", (0,0), (-1,0), SUBHEAD_BG),
            ("FONTSIZE", (0,0), (-1,-1), 7),
            ("GRID", (0,0), (-1,-1), 0.5, BORDER),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]
        for r in flag_rows:
            style_cmds.append(("BACKGROUND", (0,r), (-1,r), FLAG_RED))
        t.setStyle(TableStyle(style_cmds))
        story.append(t)
        story.append(Spacer(1, 8))

    doc = SimpleDocTemplate(filename, pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch)
    doc.build(story)
    print(f"  Created: {filename}")


# ---------------------------------------------------------------------------
# Word (.docx) generator
# ---------------------------------------------------------------------------

def create_docx(data: dict, filename: str):
    doc = Document()

    # Title
    title = doc.add_heading("labcorp", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Header info
    p = doc.add_paragraph()
    p.add_run(f"Date Collected: {data['collected']}   Date Reported: {data['reported']}   Fasting: Yes").font.size = Pt(8)

    # Patient demographics
    p2 = doc.add_paragraph()
    run = p2.add_run(f"Patient: {data['patient']}     DOB: {data['dob']}     Sex: {data['sex']}")
    run.font.size = Pt(9)
    run.font.bold = True

    doc.add_paragraph(f"Ordered Items: {data['label']}").runs[0].font.size = Pt(8)
    doc.add_paragraph("")

    for panel_name, biomarkers in data["panels"]:
        heading = doc.add_heading(panel_name, level=1)
        heading.runs[0].font.color.rgb = RGBColor(44, 62, 80)
        heading.runs[0].font.size = Pt(11)

        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"

        headers = ["Test", "Current Result and Flag", "Flag", "Units", "Reference Interval"]
        hdr_row = table.rows[0]
        for i, h in enumerate(headers):
            cell = hdr_row.cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(7)

        for name, value, flag, units, ref in biomarkers:
            row = table.add_row()
            for i, val in enumerate([name, value, flag, units, ref]):
                cell = row.cells[i]
                cell.text = val
                run = cell.paragraphs[0].runs[0]
                run.font.size = Pt(7)
                if flag in ("Low", "High") and i in (1, 2):
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(192, 57, 43)

        doc.add_paragraph("")

    doc.save(filename)
    print(f"  Created: {filename}")


# ---------------------------------------------------------------------------
# Excel (.xlsx) generator
# ---------------------------------------------------------------------------

def create_xlsx(data: dict, filename: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lab Results"

    header_fill  = PatternFill("solid", fgColor="1A1A2E")
    panel_fill   = PatternFill("solid", fgColor="2C3E50")
    subhead_fill = PatternFill("solid", fgColor="F0F0F0")
    flag_fill    = PatternFill("solid", fgColor="F8D7DA")
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # Title row
    ws["A1"] = "labcorp"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:E1")

    # Header info
    ws["A2"] = f"Date Collected: {data['collected']}   Date Reported: {data['reported']}   Fasting: Yes"
    ws["A2"].font = Font(size=8)
    ws.merge_cells("A2:E2")

    # Patient demographics
    ws["A3"] = f"Patient: {data['patient']}     DOB: {data['dob']}     Sex: {data['sex']}"
    ws["A3"].font = Font(bold=True, size=9)
    ws.merge_cells("A3:E3")

    ws["A4"] = f"Ordered Items: {data['label']}"
    ws["A4"].font = Font(size=8, italic=True)
    ws.merge_cells("A4:E4")

    row = 6
    for panel_name, biomarkers in data["panels"]:
        # Panel header
        ws.merge_cells(f"A{row}:E{row}")
        ws[f"A{row}"] = panel_name
        ws[f"A{row}"].font = Font(bold=True, color="FFFFFF", size=10)
        ws[f"A{row}"].fill = panel_fill
        ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")
        row += 1

        # Column headers
        col_headers = ["Test", "Current Result and Flag", "Flag", "Units", "Reference Interval"]
        for col, h in enumerate(col_headers, 1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = Font(bold=True, size=8)
            cell.fill = subhead_fill
            cell.border = thin
            cell.alignment = Alignment(horizontal="center")
        row += 1

        # Data rows
        for name, value, flag, units, ref in biomarkers:
            vals = [name, value, flag, units, ref]
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = Font(size=8, bold=(flag in ("Low","High") and col in (2,3)),
                                 color="C0392B" if (flag in ("Low","High") and col in (2,3)) else "000000")
                cell.border = thin
                if flag in ("Low", "High"):
                    cell.fill = flag_fill
            row += 1

        row += 1

    # Column widths
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 18

    wb.save(filename)
    print(f"  Created: {filename}")


# ---------------------------------------------------------------------------
# Generate all files
# ---------------------------------------------------------------------------

for data, name in ALL_CASES:
    print(f"\n{name} — {data['sex']}, {data['patient']}")
    create_pdf(  data, f"test_cases/{name}.pdf")
    create_docx( data, f"test_cases/{name}.docx")
    create_xlsx( data, f"test_cases/{name}.xlsx")

print(f"\nDone — {len(ALL_CASES) * 3} files created in test_cases/")
