# TODO — Platform SOB Agent

## Pending

- [ ] **Get Google Sheet ID for `[Platform PC2]`**
  - Workbook: `[Reg Commercial Team] Platform PC 2 Data`
  - Tab: `SHP & TTS PC2`
  - Once obtained, update in:
    1. `references.md` → Sheet Registry table (currently `TODO`)
    2. `gsheets_util.py` → `SHEETS` dict (add `"Platform PC2": "<ID>"`)

---

## Recently Completed (2026-05-28)

- [x] Corrected table descriptions in `references.md`:
  - `SHP/TTS ADG ADO`: clarified 4 separate tables (SHP ADG, SHP ADO, TTS ADG, TTS ADO), removed red section reference, added note as main SOB ingredient source
  - `SHP/TTS Clusters`: corrected from "SOB table" to "ADG & ADO absolutes table"
  - `(Final) Data from CF excel`: corrected from "PC2 values" to "ADG SOB & ADO SOB cross-check values"
- [x] Added new `[Platform PC2]` workbook (actual PC2 source) to both `SKILL.md` and `references.md`
- [x] Updated PC2 Step 1 in `SKILL.md` to reference correct source workbook/tab
