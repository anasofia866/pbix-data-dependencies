# NOTES — PBIXAnalyzer

Technical notes, observations, and relevant details discovered during development.

---

## PBIX Internal Structure

**Date:** 2026-04-16

- A `.pbix` file is internally a ZIP archive.
- The format changed in 2024: the new format uses individual JSON files per page under `Report/definition/pages/`, while the old format uses a single `Report/Layout` file.
- The parser automatically detects the format and routes to the appropriate method.

## Encoding

**Date:** 2026-04-16

- The `Report/Layout` file uses **UTF-16 LE/BE** — different from the standard UTF-8.
- UTF-8, UTF-8 with BOM, and Windows-1252 are also supported as fallbacks.
- Null bytes are stripped before parsing.
- The Windows console requires explicit configuration to support UTF-8 (emojis and accented characters).

## Power Query (M Code)

**Date:** 2026-04-16

- Column references in M code are extracted via regex: bracket notation `[ColumnName]` and functions such as `Table.SelectColumns`, `Table.RenameColumns`, etc.
- M code is compressed inside the `Mashup` component within the PBIX.

## GUI Threading

**Date:** 2026-04-16

- Analysis runs on a separate thread to prevent the Tkinter UI from freezing during processing.

## Deduplication

**Date:** 2026-04-16

- Usage records (`UsageRecord`) use a hash-based system to avoid duplicates during aggregation.

## Language

**Date:** 2026-04-16

- Project converted to 100% English (source code, UI, Excel output, documentation) to support public release.

---

_Add new notes above this line._
