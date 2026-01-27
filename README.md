# Beam Report Generator Using PyLaTeX

## Overview
This project generates a professional engineering PDF report for a simply supported beam using Python and PyLaTeX. The force data is read from an Excel file, and the report is created automatically without manual LaTeX editing.

The report includes a title page, table of contents, introduction with an embedded beam image, input force table recreated using LaTeX tabular, and analytically generated Shear Force Diagram (SFD) and Bending Moment Diagram (BMD) using TikZ/pgfplots.

---

## Objective
- Read beam force data from an Excel file
- Generate a structured PDF report using PyLaTeX
- Recreate the force table as selectable LaTeX text
- Plot SFD and BMD using TikZ/pgfplots (no image-based plots)

---

## Project Structure
```
Beam-Report-Generator-Using-PyLaTeX/
├── main.py
├── requirements.txt
├── README.md
│
├── data/
│   └── forces.xlsx
│
├── images/
│   └── beam.png
│
├── output/
│   └── report.pdf
│
└── docs/
    └── code_explanation.pdf
```

---

## Input Files
- **forces.xlsx**: Excel file containing beam force data
- **beam.png**: Image of the simply supported beam

---

## Report Contents
1. Title Page  
2. Table of Contents  
3. Introduction  
   - Beam description with embedded image  
   - Data source information  
4. Input Data  
   - Force table recreated using LaTeX Tabular  
5. Analysis  
   - Shear Force Diagram (TikZ/pgfplots)  
   - Bending Moment Diagram (TikZ/pgfplots)

---

## Requirements
- Python 3.x
- LaTeX distribution (MiKTeX or TeX Live)

Install Python dependencies:
```bash
pip install -r requirements.txt
```

---

## How to Run
```bash
python main.py
```

After execution, the generated PDF report will be available in the `output/` folder.

---
