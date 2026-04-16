# BRIEF — PBIXAnalyzer

## Overview

**PBIXAnalyzer** is a Windows desktop tool that analyzes Power BI files (`.pbix`) and generates an Excel report with a column-level impact assessment.

## Goal

Allow data analysts and Power BI developers to understand, before making any model changes, **where and how each column is used** — in visuals, filters, and Power Query — preventing accidental breakage of reports.

## Target Users

- Data analysts and Business Intelligence professionals
- Power BI developers
- Teams maintaining Power BI reports in production

## Tech Stack

| Component | Technology |
|---|---|
| Language | Python 3.x |
| GUI framework | Tkinter (built-in) |
| Excel generation | openpyxl |
| Analyzed format | .pbix (internal ZIP) |
| Platform | Windows |

## Key Files

| File | Role |
|---|---|
| `pbix_analyzer.py` | Core application logic (~1,200 lines) |
| `Analisar PBIX.bat` | Application launcher |
| `1_INSTALAR_PYTHON.bat` | Dependency installer |

## Main Features

1. **PBIX extraction and parsing** — supports old format (pre-2024) and new format (2024+)
2. **Visual analysis** — identifies each column in each visual, page, and role
3. **Filter analysis** — visual, page, and report-level filters
4. **Power Query analysis** — column references in M code
5. **Data model extraction** — tables, columns, DAX measures
6. **Automatic risk assessment** — HIGH / MEDIUM / LOW per column
7. **Multi-sheet Excel report** — up to 7 sheets with cross-referenced data

## Output

`.xlsx` file with the following sheets:
1. **Column Impact** — main risk analysis (color-coded)
2. **Visuals - Details** — granular column ↔ visual mapping
3. **Power Query - Columns** — references in M code
4. **Power Query Code** — full M code
5. **Data Model** — complete schema
6. **Model Tables** — tables and their usage
7. **DAX Queries** — saved DAX queries (when present)

## Risk Criteria

| Level | Criterion | Color |
|---|---|---|
| HIGH | Used in Power Query AND multiple visuals | Red |
| MEDIUM | Used in 3+ visuals | Yellow |
| LOW | Used in 1–2 visuals | Green |
| PQ Only | Referenced only in Power Query | Grey |
