"""
gsheets_util.py — Shared Google Sheets auth + helpers for Platform SOB Agent.
All API calls in this directory should import from here.
Handles:
  - macOS SSL cert fix (certifi)
  - Token refresh when expired
  - Token save-back
"""
import json, os, urllib.request, urllib.parse, ssl, datetime
import certifi

TOKEN_PATH = os.path.expanduser("~/.hermes/google_token.json")
SSL_CTX = ssl.create_default_context(cafile=certifi.where())

SHEETS = {
    "Weekly Live View": "1m4bZ11zDEHpVzQP1I4zxtFKs-vzXEYBk7NgMNJK5dpo",  # Updated 2026-05-28
    "Reg Commercial Team": "1BmRS6VjIP5_RRfQs22pgm9Ap49G_EBZJIjJDCAhmAcg",  # Updated 2026-05-28
    "Reg CNLS copy": "1cN29heWI-7trzBLMvEpznXslDmlCEmLHUcKDjT9uTqg",
    "Archive": "1F99kNADGaRxiuxkxG2Gq3A7Xvoh_bG0EM2C0xYi8Ocs",
    "Platform PC2": "11Qqg42jx_jAhfmkjr8JghVa4zwiVXdAsuo3aOGxvVKo",  # Updated 2026-05-28
}


def get_auth_headers():
    """Get authorization headers, auto-refreshing token if needed."""
    with open(TOKEN_PATH) as f:
        creds = json.load(f)

    # Check if expired
    expiry_str = creds.get("expiry", "")
    if expiry_str:
        try:
            expiry = datetime.datetime.fromisoformat(expiry_str.replace("Z", "+00:00"))
            if datetime.datetime.now(datetime.timezone.utc) >= expiry - datetime.timedelta(minutes=2):
                print("[auth] Token expired or expiring soon — refreshing...", flush=True)
                refresh_data = urllib.parse.urlencode({
                    "client_id": creds["client_id"],
                    "client_secret": creds["client_secret"],
                    "refresh_token": creds["refresh_token"],
                    "grant_type": "refresh_token",
                }).encode()
                req = urllib.request.Request(
                    creds["token_uri"], data=refresh_data,
                    headers={"Content-Type": "application/x-www-form-urlencoded"}
                )
                with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
                    token_resp = json.loads(resp.read())
                creds["token"] = token_resp["access_token"]
                new_expiry = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(
                    seconds=token_resp.get("expires_in", 3600))
                creds["expiry"] = new_expiry.isoformat().replace("+00:00", "Z")
                with open(TOKEN_PATH, "w") as f:
                    json.dump(creds, f, indent=2)
                print("[auth] Token refreshed.", flush=True)
        except Exception as e:
            print(f"[auth] Token check/refresh failed: {e}", flush=True)

    return {"Authorization": f"Bearer {creds['token']}", "Accept": "application/json"}


def api_get(sheet_id, fields):
    """Call Google Sheets GET API with auth + SSL."""
    headers = get_auth_headers()
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}?fields={urllib.parse.quote(fields)}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def api_batch_update(sheet_id, requests_list):
    """Call Google Sheets POST batchUpdate with auth + SSL."""
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}:batchUpdate"
    body = json.dumps({"requests": requests_list}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def values_get(sheet_id, range_):
    """Read cell values from a range."""
    headers = get_auth_headers()
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{urllib.parse.quote(range_, safe='!')}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def values_update(sheet_id, range_, values, input_option="USER_ENTERED"):
    """Write values to a range."""
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{urllib.parse.quote(range_, safe='!')}?valueInputOption={input_option}"
    body = json.dumps({"values": values}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="PUT")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def get_cell_formats(sheet_id, tab_name):
    """Fetch cell-level format data (borders, values) for a tab using includeGridData.
    Returns the raw sheets.data.rowData list. Use is_red_border() on each cell's borders.
    Always query the DESTINATION tab, not the source.
    """
    headers = get_auth_headers()
    quoted = urllib.parse.quote(f"'{tab_name}'", safe="'")
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
           f"?ranges={quoted}&fields=sheets.data.rowData.values(effectiveFormat.borders,effectiveValue,formattedValue)")
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        data = json.loads(resp.read())
    sheets = data.get("sheets", [])
    return sheets[0]["data"][0].get("rowData", []) if sheets else []


def is_red_border(borders):
    """Return True if any side of a cell's borders has a red-ish color (red>0.8, green<0.3, blue<0.3)."""
    if not borders:
        return False
    for side in ("top", "bottom", "left", "right"):
        side_data = borders.get(side, {})
        if not side_data:
            continue
        color = side_data.get("color", side_data.get("colorStyle", {}).get("rgbColor", {}))
        if color.get("red", 0) > 0.8 and color.get("green", 0) < 0.3 and color.get("blue", 0) < 0.3:
            return True
    return False


def values_get_formula(sheet_id, range_):
    """Read cell values with FORMULA render option — returns raw formula strings instead of computed values.
    Use when you need to inspect or preserve formulas rather than their results.
    """
    headers = get_auth_headers()
    quoted = urllib.parse.quote(range_, safe="'!")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{quoted}?valueRenderOption=FORMULA"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def values_batch_write(sheet_id, batch_data, input_option="USER_ENTERED"):
    """Write multiple non-contiguous ranges in one API call (values:batchUpdate).
    Distinct from api_batch_update which handles structural changes (tab duplication etc).
    batch_data: list of {"range": "...", "majorDimension": "ROWS", "values": [[...]]}
    """
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values:batchUpdate"
    body = json.dumps({"valueInputOption": input_option, "data": batch_data}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def clean_val(v):
    """Strip 'x' suffix and convert to float. Used when pasting SOB values.
    Handles both text strings ('6.26x') and format-driven suffixes (FORMATTED_VALUE returns '6.26x').
    Returns float if convertible, otherwise returns original value unchanged.
    """
    if v:
        s = str(v)
        if s.endswith("x"):
            try:
                return float(s[:-1])
            except ValueError:
                pass
        try:
            return float(s)
        except ValueError:
            pass
    return v


def find_newest_tab(sheet_id, prefix):
    """Find the tab whose name starts with `prefix` and has the lowest sheet index.
    In Google Sheets, newest-first ordering means lowest index = most recent tab.
    Do NOT sort by sheetId — those are arbitrary numbers, not positional.
    Returns (tab_name, sheet_id, tab_index), or None if no match.
    Use tab_index + 1 as insertSheetIndex to place a new tab right after it.
    """
    data = api_get(sheet_id, "sheets.properties")
    tabs = [
        (s["properties"]["title"], s["properties"]["sheetId"], s["properties"]["index"])
        for s in data.get("sheets", [])
    ]
    matching = [(name, sid, idx) for name, sid, idx in tabs if name.startswith(prefix)]
    if not matching:
        return None
    matching.sort(key=lambda x: x[2])  # ascending index = newest first
    return matching[0]


def sheet_copy_to(src_spreadsheet_id, src_sheet_id, dest_spreadsheet_id):
    """Copy a sheet from one spreadsheet to another (the only cross-spreadsheet way to carry formatting).
    Returns the new sheet's properties dict (including sheetId) in the destination spreadsheet.
    Used in PC2 Step 1 to bring source formatting into Archive before copyPaste PASTE_FORMAT.
    """
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{src_spreadsheet_id}/sheets/{src_sheet_id}:copyTo"
    body = json.dumps({"destinationSpreadsheetId": dest_spreadsheet_id}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def parse_tab_date(tab_name):
    """Extract sortable YYYYMMDD string from a tab name.
    Handles both pre-2026 MM/DD (e.g. '12/25 ADG ADO') and
    2026+ YY/MM/DD (e.g. '26/05/14 By cluster') formats.
    Uses re.match so it works for any tab suffix, not just ADG ADO.
    Returns None if no date pattern found at start of name.
    """
    import re
    m = re.match(r'^(\d{1,2})/(\d{2})(?:/(\d{2}))?', tab_name)
    if not m:
        return None
    if m.group(3):  # YY/MM/DD
        return f"20{m.group(1)}{m.group(2)}{m.group(3)}"
    else:            # MM/DD (pre-2026, assume 2025)
        return f"2025{m.group(1).zfill(2)}{m.group(2)}"


def scan_nonzero(sheet_id, range_):
    """
    Read a range and return (errors, nonzero) where:
      errors  — list of (cell_addr, value) for #VALUE!/#REF! cells
      nonzero — list of (cell_addr, float) for cells where abs(value) > 0.001

    values_get returns FORMATTED_VALUE strings ("0.00", "-31.56", "#VALUE!"),
    so this converts to float before comparing.
    """
    start_row = int(range_.split("!")[1].split(":")[0][1:]) if "!" in range_ else 1
    start_col = range_.split("!")[1][0] if "!" in range_ else "A"
    col_offset = ord(start_col.upper()) - ord("A")

    rows = values_get(sheet_id, range_).get("values", [])
    errors, nonzero = [], []
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            s = str(v).strip()
            addr = f"{chr(ord('A') + col_offset + j)}{start_row + i}"
            if "VALUE" in s.upper() or "REF" in s.upper():
                errors.append((addr, s))
                continue
            try:
                num = float(s.replace(",", ""))
                if abs(num) > 0.001:
                    nonzero.append((addr, num))
            except ValueError:
                pass
    return errors, nonzero


# ── Step 2 layout notes (ADG ADO copy) ────────────────────────────────────────
#
# Source layout ([Weekly Live View] ADG ADO tab):
#   SHP ADG data: cols D-K (indices 3-10), rows ~4-39
#   TTS ADG data: cols N-V (indices 13-21), SAME rows as SHP
#
# Destination layout ([Reg CNLS copy] today tab):
#   SHP ADG section: rows ~4-39   (stacked top)
#   TTS ADG section: rows ~44-79  (stacked bottom)
#
# When reading TTS values, use the same source rows but shift to cols N-V:
#   shp_rows = values_get(SID_WLV, f"'{src_tab}'!D4:K39")["values"]
#   tts_rows = values_get(SID_WLV, f"'{src_tab}'!N4:V39")["values"]  # same rows, different cols
#
# ADO sections: both source and destination use col C as month labels (same format).
