---
name: google-sheets-intelligence
description: Read, analyze, and update Google Sheets with full understanding of table structure, formulas, and cell dependencies. Accept natural language instructions to update spreadsheets while preserving formula integrity.
version: 1.0.0
author: Hermes Agent
tags: [google-sheets, spreadsheets, automation, data, formulas]
---

# Google Sheets Intelligence

Analyzes Google Sheets at a semantic level — understands table structure (headers, columns, data types), every formula in every cell, how cells reference each other, named ranges, and what's protected. Accepts natural language update instructions.

**Prerequisite:** Google Workspace OAuth must be set up first (one-time). The script shares the same auth token as the `google-workspace` skill.

## Safety Mechanism (Mandatory)

**You must ALWAYS ask for explicit permission before doing ANY spreadsheet operation — including reads/observations AND writes/edits.** Never silently look up spreadsheet structure, preview data, check dependencies, or update cells without the user's explicit go-ahead.

The permission protocol:

1. **User says something about a spreadsheet** → Ask: "Can I look at [spreadsheet name/URL] to see its structure?" Wait for a clear "yes" before running any command.
2. **You propose an edit** → Show the user exactly what will change (cells, old values, new values, downstream impact on formulas). Ask: "Shall I apply this?" Wait for a clear "yes" before executing.
3. **Batch/multi-cell edits** → Always list every change with before/after. Ask explicitly for permission.
4. **Natural-language update requests** → First inspect structure, draft the changes, present them for review, then execute only after approval.

**One question at a time.** Don't ask about structure, dependencies, and permissions in a single wall of text. Walk through it step by step.

## Quick Start

Set a shorthand:

```bash
GSI="python ~/.hermes/skills/productivity/google-sheets-intelligence/scripts/sheets_intelligence.py"
```

### Understand a spreadsheet

Get a full structural overview (headers, column types, formula count, key formulas):

```bash
$GSI structure SPREADSHEET_ID

# Or for a specific sheet:
$GSI structure SPREADSHEET_ID --sheet "Sheet1"
```

See a terminal-friendly table preview:

```bash
$GSI preview SPREADSHEET_ID --rows 15
```

Detect likely stacked table sections before editing dashboard-style sheets:

```bash
$GSI scan-sections SPREADSHEET_ID
$GSI scan-sections SPREADSHEET_ID --sheet "Sheet1"
```

List named ranges:

```bash
$GSI named-ranges SPREADSHEET_ID
```

### See how cells link together

Shows every formula and what cells it references, plus reverse dependencies (what references a given cell):

```bash
$GSI dependencies SPREADSHEET_ID
```

### Update cells

Update a single cell (if value starts with `=`, it's stored as a formula):

```bash
$GSI update SPREADSHEET_ID "Sheet1!B2" "42" --dry-run
$GSI update SPREADSHEET_ID "Sheet1!B2" "42"
$GSI update SPREADSHEET_ID "Sheet1!C2" "=B2*1.1"
```

By default, `update` refuses to overwrite a cell that currently contains a formula. Only use this after explicit user approval:

```bash
$GSI update SPREADSHEET_ID "Sheet1!C2" "123" --allow-formula-overwrite
```

Update a range:

```bash
$GSI update-range SPREADSHEET_ID "Sheet1!A1:C3" '[[1,2,3],[4,5,6]]' --dry-run
$GSI update-range SPREADSHEET_ID "Sheet1!A1:C3" '[[1,2,3],[4,5,6]]'
```

Append rows:

```bash
$GSI append SPREADSHEET_ID "Sheet1!A:C" '[[1,2,3]]' --dry-run
$GSI append SPREADSHEET_ID "Sheet1!A:C" '[[1,2,3]]'
```

Batch update multiple cells at once:

```bash
$GSI batch SPREADSHEET_ID '{"Sheet1!A1": "42", "Sheet1!C2": "=B2*1.1"}' --dry-run
$GSI batch SPREADSHEET_ID '{"Sheet1!A1": "42", "Sheet1!C2": "=B2*1.1"}'
```

`batch` uses the Google Sheets `values.batchUpdate` endpoint, so changes apply as one API request. It also blocks formula overwrites unless `--allow-formula-overwrite` is provided.

### Lightweight Direct API Pattern (when sheets_intelligence.py is slow)

The `sheets_intelligence.py` script performs deep analysis (checking every cell, formula, named range, protection, dependencies) and can **time out (120s+) on large spreadsheets with 10+ tabs**. For simple value reads/writes on known ranges, use a custom Python script with the Google Sheets API directly — it completes in <3s.

**Use this when:** you know the exact range and just need to read or write values. Skip `structure`/`preview` and go straight to the API:

```python
import json, os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Auth (same token as sheets_intelligence.py)
token_path = os.path.expanduser("~/.hermes/google_token.json")
with open(token_path) as f:
    token_data = json.load(f)
creds = Credentials.from_authorized_user_info(token_data)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

service = build("sheets", "v4", credentials=creds)
SID = "SPREADSHEET_ID"

# Read values
result = service.spreadsheets().values().get(
    spreadsheetId=SID,
    range="'SheetName'!D3:H6"
).execute()
values = result.get("values", [])

# Write values (valueInputOption="USER_ENTERED" handles % and formulas)
body = {"values": [["67.1%", "6.84%", "72.49%", "48.09%", "28.77%"]]}
result = service.spreadsheets().values().update(
    spreadsheetId=SID,
    range="'SheetName'!D3:H6",
    valueInputOption="USER_ENTERED",
    body=body
).execute()
print(f"Updated {result.get('updatedCells')} cells")
```

**When `creds.refresh(Request())` hangs:** The `google-auth` library's `Request()` can time out on some networks. If this happens:
1. The token may still be valid — check `creds.expired` first
2. Set `socket.setdefaulttimeout(10)` before calling `refresh()`
3. As a fallback, just try building the service directly — a valid unexpired token works fine without refresh

**macOS Python SSL certificate error (`CERTIFICATE_VERIFY_FAILED`):** Python 3.x on macOS often can't find root CA certificates, causing all HTTPS calls to fail with `SSL: CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate`. Fix:

```python
import ssl, certifi

# Create an SSL context with certifi's CA bundle
SSL_CTX = ssl.create_default_context(cafile=certifi.where())

# Then pass it to urllib:
with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
    ...
```

If `certifi` isn't installed: `pip install certifi`. Alternatively, run the `Install Certificates.command` that ships with the Python installer (located in `/Applications/Python 3.x/`).

**Token refresh fallback (when google-auth library hangs):** If `creds.refresh(Request())` hangs but `urllib` works, refresh the token directly:

```python
import urllib.request, urllib.parse, json, os, ssl
import certifi

SSL_CTX = ssl.create_default_context(cafile=certifi.where())

with open(os.path.expanduser("~/.hermes/google_token.json")) as f:
    creds = json.load(f)

refresh_data = urllib.parse.urlencode({
    "client_id": creds["client_id"],
    "client_secret": creds["client_secret"],
    "refresh_token": creds["refresh_token"],
    "grant_type": "refresh_token",
}).encode()

req = urllib.request.Request(creds["token_uri"], data=refresh_data,
    headers={"Content-Type": "application/x-www-form-urlencoded"})
with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
    token_resp = json.loads(resp.read())
    creds["token"] = token_resp["access_token"]
    # Optionally save back to file
    with open(os.path.expanduser("~/.hermes/google_token.json"), "w") as f:
        json.dump(creds, f, indent=2)
```

**To find the numeric sheet ID** (needed for formatting/merge operations):
```python
spreadsheet = service.spreadsheets().get(
    spreadsheetId=SID,
    fields="sheets.properties"
).execute()
for s in spreadsheet["sheets"]:
    if s["properties"]["title"] == "SheetName":
        sheet_id = s["properties"]["sheetId"]
```

**Key difference from sheets_intelligence.py:** The raw API `values().update()` does NOT check for formula overwrites or protected cells. Always verify with the user before writing.

### Full analysis (for deep inspection)

Dumps everything — every cell value/formula, all metadata:

```bash
$GSI analyze SPREADSHEET_ID --pretty
```

## ⚠️ Stacked Vertical Table Sections (Critical Pattern)

Many dashboard spreadsheets stack multiple market/region tables vertically in the same sheet. Each table has its own header row, identical column layout, but potentially **different date ranges** and **different column positions** for the same logical field.

**This is a critical source of errors** — column operations (insert/delete) are sheet-wide, not table-scoped.

### How to detect stacked tables

The `structure` and `preview` commands use row 1 as the main table shape and include a `detected_sections` hint. **Never assume row 1 represents the whole sheet.** First run:

```bash
$GSI scan-sections SPREADSHEET_ID --sheet "SheetName"
```

For deeper custom inspection, scan formatted values for additional section headers:

```python
result = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    ranges=["SheetName!A1:ZZ500"],
    fields="sheets.data.rowData.values(formattedValue)",
    includeGridData=True
).execute()
rows = result["sheets"][0]["data"][0]["rowData"]

for ri, row in enumerate(rows):
    for ci, cell in enumerate(row.get("values", [])):
        fv = cell.get("formattedValue", "")
        if fv and any(kw in str(fv) for kw in ["CNLS", "Dashboard", "Top", "NPC"]):
            print(f"Section header at row {ri+1}: {fv}")
```

### Common stacked-table pattern (found in the wild)

| Row Range | Market | Section Header | Dates | "28-Nov" at column |
|-----------|--------|---------------|-------|--------------------|
| 7-33 | MY | "MY CNLS TT S2S NPC" | 4-Dec to 28-Nov | O (col 14) |
| 34-61 | ID | "ID CNLS S2S NPC Dashboard" | 3-Dec to 27-Nov | N (col 13) |
| 62-89 | TH | "TH CNLS Top 20 Seller" | 7-Dec to 1-Dec | (not present) |
| 90-111 | VN | "VN CB S2S NPC Dashboard" | 4-Dec to 28-Nov | O (col 14) |
| 112-127 | VN v2 | "VN CNLS Top 12 Seller" | 11-Dec to 5-Dec | (not present) |
| 128-150 | TH v2 | "TH CNLS Top 20 Seller" | 16-Dec to 10-Dec | (not present) |

### Finding a specific value across all rows

Use `formattedValue` search to find where a specific column label appears across all table sections:

```python
# Search formatted values for a specific date/header
hits = []
for ri, row in enumerate(rows):
    for ci, cell in enumerate(row.get("values", [])):
        fv = cell.get("formattedValue", "")
        if fv and "28-Nov" in str(fv):
            col_letter = chr(65 + ci) if ci < 26 else chr(64 + ci//26) + chr(65 + ci%26)
            hits.append(f"{col_letter}{ri+1}: {fv}")
```

### ⚠️ The Sheet-Wide Trap: Column Operations

**Inserting or deleting a column affects ALL rows in the sheet.** If you insert a column at position O (index 14), every row from 1 to 150 shifts — including rows in other table sections. This is almost never what you want in a stacked-table layout.

**DO NOT** use `insertDimension` / `deleteDimension` (COLUMNS) to add/remove columns in a sheet with stacked tables. This will misalign every table below the insertion point.

**Alternatives for stacked-table edits:**
- **To add a column to only one table section:** Instead of inserting a column, update existing empty cells in that section's row range. Or copy the table to a new sheet, edit there, then replace.
- **To update values across all tables:** Write cell values directly using `updateCells` or `batchUpdate` with `updateCells` requests (not `insertDimension`). Target specific row ranges per table.
- **To add a new column header + data to every table:** You need one `updateCells` call per table section, each targeting that section's row range.

### ⚠️ Manual Column Insert Within a Table (No Sheet-Wide Insert)

**When to use:** You need to add a column (e.g., a new date column "23-Jan" after "28-Nov") to specific table sections in a stacked-table sheet. Since column insertion is sheet-wide and would misalign other tables, you must manually shift data within each table's row range.

**Full workflow (trial-and-error proven):**

```
Phase 1: Unmerge all section header merges in the shift range
Phase 2: copyPaste to shift data right, write new header, re-merge with expanded ranges
Phase 3: Clear leftover data in the source column
```

#### Detailed steps:

**Step 1: Identify the correct column ranges**
- Find where "28-Nov" (or your target marker) lives in each table using `formattedValue` search
- Note that different tables may have the marker in **different columns** (e.g., MY at O, ID at N, VN at O)
- The shift starts at `nov_col_idx + 1` (the column immediately after the marker)

**Step 2: Get exact merge boundaries**
```python
result = service.spreadsheets().get(
    spreadsheetId=SID, ranges=[],
    fields='sheets(properties,merges)'
).execute()
# Output like: Merge: P9:T9 (startCol=15, endCol=20)
```
**Critical:** Adjacent merged ranges like P9:T9 (15-19) and U9:AB9 (20-27) have back-to-back boundaries. The end of one is the start of the next. Use exact values — off-by-one causes "You must select all cells in a merged range" errors.

**Step 3: Unmerge → CopyPaste → Clear → Re-Merge**
```python
# Per table section:
requests = []

# 1. Unmerge all section header merges in the shift range
for ms, me in [(15, 20), (20, 28), (28, 31)]:  # exact merge boundaries
    requests.append({
        'unmergeCells': {
            'range': {
                'sheetId': SID,
                'startRowIndex': section_row_0idx,
                'endRowIndex': section_row_0idx + 1,
                'startColumnIndex': ms,
                'endColumnIndex': me
            }
        }
    })

# 2. CopyPaste: shift all columns right by 1 within the table's row range
requests.append({
    'copyPaste': {
        'source': {
            'sheetId': SID,
            'startRowIndex': section_row_0idx,
            'endRowIndex': data_end_exclusive,
            'startColumnIndex': shift_start,
            'endColumnIndex': last_col + 1       # e.g. AL = col 37 → 38
        },
        'destination': {
            'sheetId': SID,
            'startRowIndex': section_row_0idx,
            'endRowIndex': data_end_exclusive,
            'startColumnIndex': shift_start + 1,
            'endColumnIndex': last_col + 2
        },
        'pasteType': 'PASTE_NORMAL'              # copies values, formulas, formatting
    }
})

# 3. Write new column header with matching formatting
requests.append({
    'updateCells': {
        'range': {
            'sheetId': SID,
            'startRowIndex': header_row_0idx,
            'endRowIndex': header_row_0idx + 1,
            'startColumnIndex': new_col_idx,
            'endColumnIndex': new_col_idx + 1
        },
        'rows': [{
            'values': [{
                'userEnteredValue': {'stringValue': '23-Jan'},
                'userEnteredFormat': {
                    'backgroundColor': {'red': 0.937, 'green': 0.937, 'blue': 0.937},
                    'horizontalAlignment': 'CENTER',
                    'textFormat': {'bold': True, 'fontSize': 10}
                }
            }]
        }],
        'fields': 'userEnteredValue,userEnteredFormat.backgroundColor,'
                  'userEnteredFormat.horizontalAlignment,'
                  'userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize'
    }
})

# 4. Re-merge with expanded ranges (each merge expands by 1 column to the right)
# Before: P:T (15-19), U:AB (20-27), AC:AE (28-30)
# After:  Q:U (16-20), V:AC (21-28), AD:AF (29-31)
requests.append({
    'mergeCells': {
        'range': {
            'sheetId': SID,
            'startRowIndex': section_row_0idx,
            'endRowIndex': section_row_0idx + 1,
            'startColumnIndex': 16, 'endColumnIndex': 21  # Q:U
        },
        'mergeType': 'MERGE_ALL'
    }
})
requests.append({
    'mergeCells': {
        'range': {'startColumnIndex': 21, 'endColumnIndex': 29},  # V:AC
        'mergeType': 'MERGE_ALL'
    }
})
requests.append({
    'mergeCells': {
        'range': {'startColumnIndex': 29, 'endColumnIndex': 32},  # AD:AF
        'mergeType': 'MERGE_ALL'
    }
})
```

**Step 4: Clear leftover source column data**

`copyPaste` copies data — it does NOT move it. The source column retains its original values:

```python
# Clear data rows in the new column (remove leftover from copyPaste)
requests.append({
    'updateCells': {
        'range': {
            'sheetId': SID,
            'startRowIndex': header_row_0idx + 1,  # first data row
            'endRowIndex': data_end_exclusive,
            'startColumnIndex': new_col_idx,
            'endColumnIndex': new_col_idx + 1
        },
        'fields': 'userEnteredValue'   # empty rows = clear
    }
})

# Also clear leftover section header text from the original position
# (P9 will still have "SHP ADO NPC (Weekly)" since unmerge left the text there)
requests.append({
    'updateCells': {
        'range': {
            'sheetId': SID,
            'startRowIndex': section_row_0idx,
            'endRowIndex': section_row_0idx + 1,
            'startColumnIndex': 15,  # original P column
            'endColumnIndex': 16
        },
        'fields': 'userEnteredValue'
    }
})
```

#### Pitfalls to avoid

| Mistake | Symptom | Fix |
|---------|---------|-----|
| Wrong merge boundary (off by 1) | "You must select all cells in a merged range" | Check exact start/end from API's `merges` field |
| Not unmerging before copyPaste | "You can't perform a paste that partially intersects a merge" | Always unmerge ALL adjacent merges in the shift range first |
| Not clearing source column | Old data appears in the new "empty" column | Use `updateCells` with empty rows after copyPaste |
| Assuming copyPaste moves data | Source column has duplicate values | Explicitly clear source cells after the copy |
| Batch unmerge + merge in same call | Merges created at wrong positions | Do Phase 1 (unmerge) as a separate batchUpdate, then Phase 2 (copy + write + re-merge) |
| Same merge boundaries for all tables | Tables with different date ranges get wrong column placement | The shift-start column differs per table (MY at P, ID at O, etc.) — but the RE-MERGE positions (Q:U, V:AC, AD:AF) are always the same because the shift is always +1 |

### Identifying table boundaries

Look for these patterns to find where one table ends and the next begins:

1. **"Updated on" rows** — Often mark the start of a new table section (e.g., "Updated on" with a date in the next column)
2. **Blank separator rows** — One or more empty rows between tables
3. **Section title rows** — Rows with section names like "MY CNLS..." or "ID CNLS..."
4. **Date range label rows** — Rows containing things like "28Nov - 4Dec" or "Dec 1 - Dec 7"

Script to extract table boundaries:

```python
tables = []
current_table = None
for ri, row in enumerate(rows):
    first_cell = (row.get("values", [{}])[0]).get("formattedValue", "") if row.get("values") else ""
    # Detect section start
    if first_cell and any(kw in first_cell for kw in ["Updated on", "MY ", "ID ", "TH ", "VN "]):
        if current_table and current_table["data_end"]:
            tables.append(current_table)
        current_table = {"start_row": ri+1, "name": first_cell[:30], "data_end": None}
    # Detect data row
    if current_table and row.get("values", [{}])[0].get("formattedValue", ""):
        current_table["data_end"] = ri+1
```

## Cell Formatting (Colors, Styles) — NOT in sheets_intelligence.py

The `sheets_intelligence.py` script **does not support reading or writing cell formatting** (background colors, font colors, borders, etc.). It only handles cell *values* and *formulas*.

### Reading cell background colors

Use the Google Sheets API directly with the existing OAuth token. Write a Python script:

```python
import json, os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

token_path = os.path.expanduser("~/.hermes/google_token.json")
with open(token_path) as f:
    token_data = json.load(f)
creds = Credentials.from_authorized_user_info(token_data)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

service = build("sheets", "v4", credentials=creds)

result = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    ranges=["SheetName"],
    fields="sheets.data.rowData.values(effectiveFormat.backgroundColor,effectiveValue,formattedValue)",
    includeGridData=True
).execute()

row_data = result["sheets"][0]["data"][0]["rowData"]
```

**Color data structure:** Each cell's `effectiveFormat.backgroundColor` returns an RGB dict, e.g. `{"red": 0.96, "green": 0.80, "blue": 0.80}`. Empty/uncolored cells return `{}`. Values are floats 0.0–1.0. Some cells use `backgroundColorStyle.rgbColor` instead — check both.

**Common color values found in sheets:**
| Color | RGB |
|-------|-----|
| White (default) | (1.00, 1.00, 1.00) |
| Light gray | (0.80, 0.80, 0.80) |
| Medium gray | (0.40, 0.40, 0.40) |
| Light pink/red | (0.96, 0.80, 0.80) |
| Light yellow | (1.00, 0.90, 0.60) |
| Yellow | (1.00, 1.00, 0.00) |
| Light blue | (0.82, 0.89, 0.95) |
| Light green | (0.73, 0.86, 0.69) |
| Orange | (0.90, 0.57, 0.22) |
| Dark blue text | (0.03, 0.22, 0.39) |

**Tip:** Always iterate ALL cells with colors, group by row, and present the full picture to the user before making changes. Sheets can have 15+ unique background colors across thousands of cells.

### Reading cell border colors

The `effectiveFormat.borders` field contains per-side border styling. Each side (top, bottom, left, right) has its own color and style.

```python
# In the fields parameter, request borders instead of backgroundColor:
result = service.spreadsheets().get(
    spreadsheetId=SID,
    ranges=["'Tab Name'"],
    fields="sheets.data.rowData.values(effectiveFormat.borders,effectiveValue,formattedValue)",
    includeGridData=True
).execute()
```

**Border data structure per cell:**
```python
{
    "top":    {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "bottom": {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "left":   {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "right":  {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}}
}
```

Possible styles: `SOLID`, `DOTTED`, `DASHED`, `DOUBLE`, `SOLID_MEDIUM`, `SOLID_THICK`, `NONE`.

Some cells use `colorStyle.rgbColor` instead of `color` — always check both:
```python
color = side_data.get('color', side_data.get('colorStyle', {}).get('rgbColor', {}))
```

**Common use case: find a "red-bordered section" (e.g., cells marked by red borders as copy targets):**

```python
def is_red_border(borders, threshold=0.8):
    \"\"\"Check if any border side has a red-ish color.\"\"\"
    if not borders:
        return False
    for side in ('top', 'bottom', 'left', 'right'):
        side_data = borders.get(side, {})
        if not side_data:
            continue
        color = side_data.get('color', side_data.get('colorStyle', {}).get('rgbColor', {}))
        # Threshold: red > 0.8, green and blue < 0.3
        if color.get('red', 0) > threshold and color.get('green', 0) < 0.3 and color.get('blue', 0) < 0.3:
            return True
    return False
```

**⚠️ Trap: borders vs background colors.** Users may say "red cells" or "red section" meaning either red borders OR red background fill. These are different API fields (`effectiveFormat.borders` vs `effectiveFormat.backgroundColor`) and are queried with different `fields` parameters. When in doubt, read both and present the findings to the user to clarify which they mean.

### Reading cell border colors

The `effectiveFormat.borders` field contains per-side border styling. Each side (top, bottom, left, right) has its own color and style, accessible via `.color` (Color) or `.colorStyle.rgbColor` (ColorStyle).

```python
result = service.spreadsheets().get(
    spreadsheetId=SID,
    ranges=["'Tab Name'"],
    fields="sheets.data.rowData.values(effectiveFormat.borders,effectiveValue,formattedValue)",
    includeGridData=True
).execute()
```

**Border data structure per cell:**
```python
# Example: cell with red borders on all 4 sides
{
    "top":    {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "bottom": {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "left":   {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}},
    "right":  {"style": "SOLID", "color": {"red": 1.0, "green": 0.0, "blue": 0.0}}
}
```

Possible styles: `SOLID`, `DOTTED`, `DASHED`, `DOUBLE`, `SOLID_MEDIUM`, `SOLID_THICK`, `NONE`.

Some cells use `colorStyle.rgbColor` instead of `color` — always check both:
```python
color = side_data.get('color', side_data.get('colorStyle', {}).get('rgbColor', {}))
```

**Common use case: find a "red-bordered section"** (e.g. cells highlighted by red borders to mark them as copy targets):
```python
def is_red_border(borders, threshold=0.8):
    """Check if any border side has a red-ish color."""
    if not borders:
        return False
    for side in ('top', 'bottom', 'left', 'right'):
        side_data = borders.get(side, {})
        if not side_data:
            continue
        color = side_data.get('color', side_data.get('colorStyle', {}).get('rgbColor', {}))
        # Threshold: red > threshold, green and blue < 0.3
        if color.get('red', 0) > threshold and color.get('green', 0) < 0.3 and color.get('blue', 0) < 0.3:
            return True
    return False

# Scan all rows for red-bordered ranges
for ri, row in enumerate(row_data):
    for ci, cell in enumerate(row.get("values", [])):
        borders = cell.get("effectiveFormat", {}).get("borders")
        if is_red_border(borders):
            print(f"Red border at row {ri+1}, col {ci+1}")
```

**⚠️ Trap: confusing borders with background colors.** Users may say "red section" meaning either red borders OR red background fill. These are different API fields (`effectiveFormat.borders` vs `effectiveFormat.backgroundColor`) and require different `fields` parameters. When in doubt, read both and present the findings — or simply ask the user which visual marker they mean.

### Writing cell background colors

To change a cell's background color, use the `batchUpdate` endpoint with `repeatCell` requests:

```python
requests = []
for cell_ref in cell_refs:
    sheet_name, cell = cell_ref.split("!")
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": SHEET_ID,  # numeric sheet ID
                "startRowIndex": ROW-1,
                "endRowIndex": ROW,
                "startColumnIndex": COL-1,
                "endColumnIndex": COL
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 1.0  # blue
                    }
                }
            },
            "fields": "userEnteredFormat.backgroundColor"
        }
    })

body = {"requests": requests}
service.spreadsheets().batchUpdate(
    spreadsheetId=SPREADSHEET_ID,
    body=body
).execute()
```

**Column letter to index:** `ord(letter) - 65` for A-Z. For AA+, strip prefix and use `(ord(prefix)-64)*26 + ord(letter)-65`.

**Find numeric sheet ID:** 
```python
spreadsheet = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    fields="sheets.properties"
).execute()
for s in spreadsheet["sheets"]:
    if s["properties"]["title"] == "SheetName":
        sheet_id = s["properties"]["sheetId"]
```

**Key fields for `repeatCell`:**
| Aspect | Field path |
|--------|-----------|
| Background color | `userEnteredFormat.backgroundColor` |
| Font color | `userEnteredFormat.textFormat.foregroundColor` |
| Bold | `userEnteredFormat.textFormat.bold` |
| Font size | `userEnteredFormat.textFormat.fontSize` |
| Horizontal align | `userEnteredFormat.horizontalAlignment` |
| All formatting at once | `userEnteredFormat` (resets unspecified to default) |

### ⚠️ Merged Cells — Critical

**Merged cells are a common trap.** When a cell you want to format is part of a merged range, you MUST apply the formatting to the entire merged range, not just the individual cell. Formatting only the top-left cell of a merged range may NOT produce a visible change.

**To discover merged ranges:**

```python
result = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    ranges=[],
    fields="sheets(properties,merges)"
).execute()

for s in result["sheets"]:
    if s["properties"]["title"] == "Sheet1":
        for m in s.get("merges", []):
            print(m)
            # e.g. {"sheetId": 123, "startRowIndex": 8, "endRowIndex": 9,
            #       "startColumnIndex": 28, "endColumnIndex": 31}
            # This means row 9, cols AC:AE are merged
```

**Merge data structure:** `startRowIndex` and `endRowIndex` are 0-indexed (exclusive end). `startColumnIndex`/`endColumnIndex` are 0-indexed column indices (exclusive end). So `startRowIndex=8, endRowIndex=9, startColumnIndex=28, endColumnIndex=31` = AC9:AE9 (row 9, columns AC through AE, 3 columns).

**When applying repeatCell to a merged range, use the exact merge range boundaries:**

```python
"range": {
    "sheetId": SHEET_ID,
    "startRowIndex": merge["startRowIndex"],
    "endRowIndex": merge["endRowIndex"],
    "startColumnIndex": merge["startColumnIndex"],
    "endColumnIndex": merge["endColumnIndex"]
}
```

**For maximum compatibility, set BOTH `backgroundColor` AND `backgroundColorStyle`:**

```python
"userEnteredFormat": {
    "backgroundColor": {"red": 0.1, "green": 0.2, "blue": 0.7},
    "backgroundColorStyle": {
        "rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.7}
    }
}
"fields": "userEnteredFormat.backgroundColor,userEnteredFormat.backgroundColorStyle"
```

The `backgroundColor` (Color) field is deprecated but still read by some API versions. The `backgroundColorStyle` (ColorStyle with `rgbColor`) is the modern replacement. Setting both ensures widest compatibility.

**Row heights and hidden rows:** Some rows may have `pixelSize=0` or `hiddenByUser=true` in `rowMetadata`. Formatting changes on hidden rows won't be visible. Check row metadata before applying formatting:

```python
result = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    fields="sheets.properties,sheets.rowMetadata"
).execute()
for s in result["sheets"]:
    if s["properties"]["title"] == "Sheet1":
        for ri, meta in enumerate(s.get("rowMetadata", [])):
            if meta.get("hiddenByUser"):
                print(f"Row {ri+1} is hidden")
```

### ⚠️ Conditional Formatting Overrides

Conditional formatting rules apply AFTER `userEnteredFormat` and take visual precedence. This means:

- You can set `userEnteredFormat.backgroundColor` to blue via API
- But `effectiveFormat.backgroundColor` might still show pink/red from a conditional rule
- **The user won't see your color change**

**How to detect this:**
After applying formatting, always check if `effectiveFormat` matches your change:
```python
result = service.spreadsheets().get(
    spreadsheetId=SPREADSHEET_ID,
    ranges=["Sheet1!A1"],
    fields="sheets.data.rowData.values(userEnteredFormat.backgroundColor,effectiveFormat.backgroundColor)",
    includeGridData=True
).execute()
```

If `userEnteredFormat.backgroundColor` shows your color but `effectiveFormat.backgroundColor` shows something different, conditional formatting is overriding you.

**To fix:** You need to either remove the conditional formatting rules, or add the cells to the rule's exclusion list.

### ⚠️ Browser Cache

Google Sheets caches the sheet rendering in your browser. After API-based formatting changes, the user may need to:

1. **Hard refresh:** `Cmd+Shift+R` (Mac) or `Ctrl+F5` (Windows/Linux)
2. **Open in incognito/private window** — guaranteed fresh load from server
3. **Switch tabs and back** — sometimes triggers a re-render

Don't assume API success = user visibility. Always verify from the user's perspective.

### ⚠️ Other Caveats

- The `sheets_intelligence.py` script cannot read or write formatting — you must use the raw Sheets API via a custom Python script.
- Setting any format field via `repeatCell` resets unspecified formatting fields to default. Use precise `fields` parameter to scope changes.
- `sheets_intelligence.py` preview and structure commands show blank headers for unlabeled columns — use the custom API approach for accurate cell-level data.

## Important Rules

1. **Permission first, always.** Never read or write spreadsheet data without the user's explicit go-ahead. Ask clearly and wait for a "yes".
2. **Never overwrite formula cells with plain values** unless the user explicitly asks. If a cell contains `=SUM(A1:A10)`, write `=SUM(...)` not a static number.
3. **Preserve cell references** when updating formula arguments. Changing `=B2*1.1` to `=C2*1.1` impacts anything depending on that cell.
4. **Always confirm before batch updates** — show the user what will change and how it affects dependents.
5. **Named ranges are safer than raw ranges** — prefer them when available.
6. **Protected sheets/cells cannot be written to** — check `is_protected` in the structure output.
7. **⚠️ Conditional formatting overrides manual cell colors.** When you set `userEnteredFormat.backgroundColor` via the API, conditional formatting rules on those cells will silently override it. The `effectiveFormat.backgroundColor` will still show the conditional format's color, not yours. To actually change the visual color, you must either:
   - Remove the conditional formatting rules that target those cells, OR
   - Add the cells to the conditional format rule's exclusion list
   Always verify by reading `effectiveFormat.backgroundColor` after making changes — if it doesn't match what you set, conditional formatting is interfering.
8. **Formatting changes only affect the visual appearance** — cell values, formulas, and data integrity are preserved. Background color changes via `repeatCell` with `fields: "userEnteredFormat.backgroundColor"` are purely cosmetic.

## Adding New Month Rows to Formula-Driven Stacked Tables (Critical Pattern)

When a user asks you to add a new month row (e.g., Jun'26 below May'26) to a sheet with stacked formula-driven tables, follow this workflow. This has a critical pitfall: **replicate the formula pattern, not the values.**

### The Formula Trap

When the user says "follow the same pattern as previous months," they mean the **formula structure**, not the numeric values. Reading with `FORMATTED_VALUE` shows the computed result (e.g., `4.03 x`) -- you need `FORMULA` render option to see the actual references (e.g., `=C73/C112`).

**Always read formulas first:**
```python
result = service.spreadsheets().values().get(
    spreadsheetId=SID, range="Sheet1!A15:W15",
    valueRenderOption='FORMULA'  # Shows "=C74/C113" not "3.96 x"
).execute()
```

### The Workflow

1. **Map reference patterns**: Read a few existing month rows with `FORMULA` to understand how each column is calculated.

2. **Insert rows bottom-to-top**: Use `insertDimension` from largest row index to smallest to avoid position shifts.

3. **Write formulas with incremented references**: Build formula strings with row references pointing to the new month's data rows (e.g., May `=C74/C113` becomes Jun `=C75/C114`). Use `valueInputOption='USER_ENTERED'`.

4. **Copy formatting**: Use `copyPaste` with `pasteType: 'PASTE_FORMAT'` to copy formatting without overwriting formulas.

5. **Clearing cells safely**: When the user asks to delete numbers, read with `FORMULA` first -- only clear cells that are static numbers (no `=` prefix), never clear formulas.

```python
def is_static_number(val_str):
    if not val_str or not str(val_str).strip(): return False
    s = str(val_str).strip()
    if s.startswith('='): return False
    try: float(s.replace(',', '')); return True
    except: return False
```

## Creating Cross-Sheet Formula Dashboards (INDEX/MATCH)

When a user asks you to pull data from one sheet/tab into another using INDEX MATCH formulas, follow this workflow. This is a non-trivial multi-step operation that requires understanding the source data structure, generating formulas programmatically, and handling formatting correctly.

### Step 1: Understand the source data structure

Source data in "flattened" formats (common in dashboards and exports) typically has two parts:

- **Dimension columns:** Identification columns per row — market, seller type, metric name, etc.
- **Time series columns:** Monthly/quarterly data spanning many columns (e.g., R through AO = 24 months)

Use the `structure` command to get an overview, then write a Python script to enumerate unique dimension combinations:

```python
# After getting all rows with values()...
seen = set()
for i, row in enumerate(rows):
    if len(row) >= 17:
        col_n = str(row[13]).strip()  # Market column
        col_o = str(row[14]).strip()  # Seller type column
        col_q = str(row[16]).strip()  # Metric column
        key = (col_n, col_o, col_q)
        if key not in seen:
            seen.add(key)
```

### Step 2: Design the target table layout

Common patterns:
- **Pattern A (metrics as top-level groups):** Row 1 = ADG | ADO | ABS (each merged across 6 regions), Row 2 = MY SG TH VN PH ID (repeated per metric). Good for comparing one metric across regions.
- **Pattern B (regions as top-level groups, preferred for this user):** Row 1 = MY | SG | TH | VN | PH | ID (each merged across 3 metrics), Row 2 = ADG | ADO | ABS (repeated per region). Good for seeing all metrics for one region together.

### Step 3: Generate INDEX MATCH formulas programmatically

Key formula pattern for multi-condition lookup:

```
=IFERROR(INDEX('Source Sheet'!$DATA_RANGE$5:$LAST_ROW$LAST_COL, 
  MATCH(1, ('Source Sheet'!$DIM_COL_A$5:$DIM_COL_A$LAST_ROW="VALUE_A")*
           ('Source Sheet'!$DIM_COL_B$5:$DIM_COL_B$LAST_ROW="VALUE_B")*
           ('Source Sheet'!$DIM_COL_C$5:$DIM_COL_C$LAST_ROW="VALUE_C"), 0),
  MATCH(DATE(YEAR,MONTH,1), 'Source Sheet'!$MONTH_ROW$ROW:$MONTH_ROW$ROW, 0)), "")
```

**Critical components:**
- `IFERROR(..., "")` wraps the formula — returns blank if no match instead of #N/A
- Multi-condition MATCH uses array multiplication (`*`) as AND logic — works directly in Google Sheets without ARRAYFORMULA
- The column MATCH finds the correct monthly column by matching DATE(Y,M,1) against the date header row in the source
- `$` anchors for copy-paste safety (absolute references to source ranges, but row-relative for the target table)

**Dynamic month reference from target column A:**
```python
# Instead of hardcoding DATE(2025,5,1), reference the month label in column A:
DATE(VALUE(LEFT($A3,4)), VALUE(MID($A6,2)), 1)
```

**Calculated metrics (e.g., ABS = ADG/ADO):**
```python
# Use cell references within the same target sheet
f'=IFERROR(ROUND({adg_cell_ref}/{ado_cell_ref}, 1), "")'
```
This keeps the spreadsheet self-calculating and avoids duplicating lookups.

### Step 4: Write all formulas via batchUpdate

Always use `values().batchUpdate()` with `valueInputOption: "USER_ENTERED"` to write all formulas in a single API call:

```python
data_updates = []
for each cell that needs a formula:
    data_updates.append({
        "range": f"Sheet1!{col_letter}{row}",
        "values": [[formula_string]]
    })

body = {"valueInputOption": "USER_ENTERED", "data": data_updates}
service.spreadsheets().values().batchUpdate(spreadsheetId=SID, body=body).execute()
```

**⚠️ Critical pitfall: batchUpdate overwrites ALL cells in the range.** If you write formulas for columns B-S but set column A to "" (empty string), you'll erase the month labels in column A. Always re-write month labels (or any pre-existing data) in a separate `values().update()` call afterward.

### Step 5: Apply formatting

After writing formulas, apply number formats via `batchUpdate` with `repeatCell`:

```python
# Comma format for whole numbers
{"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}

# 1 decimal place
{"numberFormat": {"type": "NUMBER", "pattern": "0.0"}}

# Percentage
{"numberFormat": {"type": "NUMBER", "pattern": "0.00%"}}
```

Apply these per-column-range using `repeatCell`:

```python
requests.append({
    "repeatCell": {
        "range": {"sheetId": SHEET_ID, "startRowIndex": DATA_START, "endRowIndex": DATA_END,
                  "startColumnIndex": COL, "endColumnIndex": COL + 1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
        "fields": "userEnteredFormat.numberFormat"
    }
})
```

### Step 6: Verify

Read back a few cells using `FORMATTED_VALUE` (shows how the value actually displays) and `FORMULA` (verifies the formula string):

```python
result = service.spreadsheets().values().batchGet(
    spreadsheetId=SID,
    ranges=["Sheet1!B3:G3", "Sheet1!H3:M3"],
    valueRenderOption="FORMATTED_VALUE"
).execute()
```

### Common Pitfalls

| Problem | Symptom | Fix |
|---------|---------|------|
| Month shows as serial number (45778) | Google auto-converts "2025-05" to a date | Use `valueInputOption: "RAW"` for month labels, not `USER_ENTERED` |
| BatchUpdate erases month labels | Empty strings in batchUpdate range overwrite data | Write month labels in a separate call after formulas |
| Unmerge + merge in same batch | Merges created at wrong positions | Split into two separate batchUpdate calls |
| Formula returns #N/A for valid data | Column index mismatch (wrong date header row) | Use row 3 (not row 4) as the date header in MATCH; verify the date header row has correct serial dates starting from Jan 1st of the first period |
| Number format not showing as expected | Formula rounds but display still shows full decimals | Apply both ROUND() in formula AND numberFormat pattern in repeatCell |
| "8.0" displays as "8" | ROUND formula returns 8.0 but number format is "General" | Apply number format pattern "0.0" via repeatCell, not just ROUND() |

## Permission Workflow for Updates

When the user says something about a spreadsheet, follow this sequence one step at a time (never pack multiple questions into one message):

### Step 1: Ask permission to inspect

Ask the user: *"Can I look at [spreadsheet name/URL] to see its structure?"* Wait for their answer.

### Step 2: Understand the structure

After they say yes, run:

```bash
$GSI structure $SPREADSHEET_ID
```

### Step 3: Check dependencies if updating calculated cells

If the update affects cells involved in formulas:

```bash
$GSI dependencies $SPREADSHEET_ID --sheet "Sheet1"
```

### Step 4: Draft the changes and present for review

Summarize what will change:
- Which cells will be modified
- Old value → New value
- Any formulas that will be affected downstream

Ask: *"Shall I apply this?"* Wait for a clear "yes" before executing.

### Step 5: Apply the update

Use the appropriate command (`update`, `update-range`, `append`, or `batch`).

### Step 6: Verify

```bash
$GSI preview $SPREADSHEET_ID --rows 5
```

The spreadsheet ID is in the URL:
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=...
```

Or find it via Drive search:

```bash
GAPI="python ~/.hermes/skills/productivity/google-workspace/scripts/google_api.py"
$GAPI drive search "filename" --max 10
```

## Cell Formatting (Background Colors)


## First-Time Setup

The setup scripts need `PYTHONPATH` to find Hermes internal modules:

```bash
export PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH"
GSETUP="python3 $HOME/.hermes/skills/productivity/google-workspace/scripts/setup.py"
```

### Step 1: Install dependencies
```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### Step 2: Create OAuth credentials
Go to https://console.cloud.google.com/apis/library and enable:
- Google Sheets API
- Google Drive API (needed to search for spreadsheets by name)

Then go to https://console.cloud.google.com/apis/credentials
- Create Credentials → OAuth 2.0 Client ID → **Desktop app**
- Download the JSON file (`client_secret_XXXXX.json`)

If the app is in "Testing" status, add your email as a test user at:
https://console.cloud.google.com/auth/audience

### Step 3: Register the client secret
```bash
PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH" $GSETUP --client-secret /path/to/client_secret.json
```

### Step 3b: (Optional) Narrow the scopes

By default `setup.py` requests ALL scopes (gmail, calendar, drive, sheets, docs,
contacts). If your OAuth consent screen doesn't have all of these configured,
edit the `SCOPES` list in `setup.py` to only include what you need.

Find `SCOPES = [...]` near the top of:
```
~/.hermes/skills/productivity/google-workspace/scripts/setup.py
```

Delete any scopes not configured in your consent screen. For Sheets + Drive only:
```python
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
```

Restore the full list after generating the URL.

### Step 4: Get auth URL

**Clear any stale pending auth first:**
```bash
rm -f ~/.hermes/google_oauth_pending.json
```

Then:
```bash
PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH" $GSETUP --auth-url
```

**CRITICAL: Verify the redirect URI matches your client secret.** Read the `redirect_uris`
array in your `client_secret.json` file. The default in setup.py is `http://localhost`.
If your client secret uses a different URI (e.g. `http://localhost:1`), update
`REDIRECT_URI` in `setup.py` to match before running `--auth-url`.

**Note:** The setup.py `--services` and `--format` flags are NOT supported.
The URL includes PKCE parameters (code_challenge, state) which are required
by Google. Do NOT craft a manual URL without these — it will fail with
"Required parameter is missing: response_type".

If the chat platform mangles the URL (missing parameters, truncated):
- Save the URL to a text file on the user's desktop (e.g., `~/Desktop/google_auth_url.txt`) and ask them to copy-paste it into their browser.
- Even better: create a self-contained HTML file with a clickable button and auto-redirect meta tag. This avoids chat-app URL corruption entirely. Example:
  ```html
  <meta http-equiv="refresh" content="0;url=AUTH_URL_HERE">
  <a href="AUTH_URL_HERE" style="padding:14px 28px;background:#1a73e8;color:white;">
    Sign in with Google
  </a>
  ```
  Save as `~/Desktop/google_auth_redirect.html` and ask the user to open it.

### Step 5: User authorizes
Send the URL to the user. They need to:
1. Open it in a browser
2. Sign in and consent
3. Get redirected to `http://localhost/?code=...` (page will fail to load — expected)
4. Copy the **entire URL** from the address bar (or just the `code=...` part)

### Step 6: Exchange code for token
```bash
PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH" $GSETUP --auth-code "CODE_OR_URL_FROM_STEP_5"
```

If the code expired: delete `~/.hermes/google_oauth_pending.json`, re-run `--auth-url`,
and have the user try again with the fresh URL.

### Step 7: Verify
```bash
PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH" $GSETUP --check
# Should print AUTHENTICATED
```

## Troubleshooting

| Problem | Fix |
|---------|------|
| `No Google token found` | Run the setup steps above |
| `HttpError 403` | Enable Google Sheets API in Cloud Console |
| `ModuleNotFoundError: hermes_constants` | Prefix commands with `PYTHONPATH="$HOME/.hermes/hermes-agent:$PYTHONPATH"` |
| Formula stored as text | Ensure value starts with `=` |
| `namedRanges` not found | Check that named ranges exist in Data → Named ranges |
| `Access blocked / invalid_request` | Use the PKCE URL from setup.py (not a manual URL); check redirect URI matches `client_secret.json` |
|| `Error 400: invalid_request (missing params)` | Chat platform likely mangled the URL — save to a local HTML file instead of sending the link via chat; also ensure PKCE params (code_challenge, state) are present |

### All API calls hang/timeout (terminal sandbox)

**Symptom:** Every Google Sheets API call (via `sheets_intelligence.py`, raw API with `urllib`, or google client library) hangs for 20-90 seconds then times out. DNS resolves (`sheets.googleapis.com` resolves), but HTTPS connections never complete. Even `curl https://www.google.com` times out.

**Root cause:** The Hermes agent's terminal sandbox has **no outbound HTTPS access** to external hosts. The terminal runs in an isolated VM that cannot reach the public internet. The `web_search` and `web_extract` tools work because they route through the host machine's network, but those tools cannot authenticate to Google APIs.

**How to diagnose:**
```bash
# Quick connectivity test
curl -s -o /dev/null -w "%{http_code} %{time_total}s" https://sheets.googleapis.com/v4/spreadsheets/ --connect-timeout 5 --max-time 10
# -> 000 5.00s  =  no connectivity

# Check token expiry
python3 -c "import json,os; t=json.load(open(os.path.expanduser('~/.hermes/google_token.json'))); print('Expired:', t.get('expiry'))"
```

**What to do:**
- The token is stored at `~/.hermes/google_token.json`. It has a refresh token, but refreshing also requires HTTPS to `oauth2.googleapis.com` which is also blocked.
- **If you need to read/write Google Sheets right now:** Ask the user to share the relevant values/cell references directly, or provide the data they want written. You can present the intended action and have the user execute it manually.
- **For long-term cron jobs:** Any workflow that reads/writes Google Sheets via the terminal sandbox WILL fail. Design cron prompts that either (a) skip the terminal-based API calls, or (b) ask the user to manually trigger the sheet operations.
- **Check if network access has been restored:** Re-run the curl connectivity test above. If it returns `200` instead of `000`, the sandbox network is working again.
