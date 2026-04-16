# DECISIONS — PBIXAnalyzer

Log of technical and product decisions made throughout the project.

---

## Use `.ai` folder for structured documentation

**Date:** 2026-04-16

| Field | Detail |
|---|---|
| Decision | Use Markdown files in the `.ai` folder for internal documentation |
| Alternatives considered | No structured documentation |
| Reason | Preserve context between AI-assisted work sessions and maintain a reasoning history |
| Impact | Low — documentation only, does not affect the code |

---

## Convert entire codebase to English

**Date:** 2026-04-16

| Field | Detail |
|---|---|
| Decision | Convert 100% of source code, UI strings, Excel output, and documentation to English |
| Alternatives considered | Keep Portuguese (original language) |
| Reason | Project is being made public; English maximises accessibility for contributors and users |
| Impact | Medium — all user-facing strings, Excel sheet names, column headers, log messages, comments, and docstrings updated. Output Excel files will now have English sheet names and headers. The `.bat` file names (`Analisar PBIX.bat`, `1_INSTALAR_PYTHON.bat`) retain their original names but their content is now fully in English. |

---

_Add new decisions above this line, with date and context._
