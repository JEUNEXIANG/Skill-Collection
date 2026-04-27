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
$GSI update SPREADSHEET_ID "Sheet1!B2" "42"
$GSI update SPREADSHEET_ID "Sheet1!C2" "=B2*1.1"
```

Update a range:

```bash
$GSI update-range SPREADSHEET_ID "Sheet1!A1:C3" '[[1,2,3],[4,5,6]]'
```

Append rows:

```bash
$GSI append SPREADSHEET_ID "Sheet1!A:C" '[[1,2,3]]'
```

Batch update multiple cells at once:

```bash
$GSI batch SPREADSHEET_ID '{"Sheet1!A1": "42", "Sheet1!C2": "=B2*1.1"}'
```

### Full analysis (for deep inspection)

Dumps everything — every cell value/formula, all metadata:

```bash
$GSI analyze SPREADSHEET_ID --pretty
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
| `Error 400: invalid_request (missing params)` | Chat platform likely mangled the URL — save to a local HTML file instead of sending the link via chat; also ensure PKCE params (code_challenge, state) are present |
