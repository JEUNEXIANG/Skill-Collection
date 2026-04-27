#!/usr/bin/env python3
"""
Google Sheets Intelligence — analyze, understand, and update spreadsheets.

Understands table structure (headers, columns, data types), formulas,
and cell relationships. Accepts natural-language-friendly update instructions.

Prerequisites:
  pip install google-api-python-client google-auth-oauthlib google-auth-httplib2

Auth:
  Uses the same token file as the google-workspace skill:
  ~/.hermes/google_token.json (OAuth2 credentials)

Usage:
  # Analyze a spreadsheet (structure + all data with formulas)
  python sheets_intelligence.py analyze SPREADSHEET_ID [--sheet "Sheet1"]

  # Get structural overview (column names, types, formula count, dependencies)
  python sheets_intelligence.py structure SPREADSHEET_ID

  # Update a single cell (value or formula)
  python sheets_intelligence.py update SPREADSHEET_ID "Sheet1!A1" "42"
  python sheets_intelligence.py update SPREADSHEET_ID "Sheet1!B2" "=A1*2"

  # Update a range (JSON array of arrays)
  python sheets_intelligence.py update-range SPREADSHEET_ID "Sheet1!A1:C3" '[[1,2,3],[4,5,6]]'

  # Append rows
  python sheets_intelligence.py append SPREADSHEET_ID "Sheet1!A:C" '[[1,2,3]]'

  # Batch update multiple cells
  python sheets_intelligence.py batch SPREADSHEET_ID '{"Sheet1!A1": "42", "Sheet1!B2": "=A1*2"}'

  # Show table preview in terminal-friendly format
  python sheets_intelligence.py preview SPREADSHEET_ID [--rows 10]

  # List all named ranges
  python sheets_intelligence.py named-ranges SPREADSHEET_ID

  # Get cell dependencies (which cells reference which)
  python sheets_intelligence.py dependencies SPREADSHEET_ID [--sheet "Sheet1"]
"""

import argparse
import json
import os
import re
import sys
import textwrap
from collections import defaultdict
from pathlib import Path

# ─── Auth ───────────────────────────────────────────────────────────────────

TOKEN_PATH = Path.home() / ".hermes" / "google_token.json"



def get_credentials():
    """Load and refresh credentials from the google-workspace token file."""
    if not TOKEN_PATH.exists():
        print("ERROR: No Google token found. Run the google-workspace setup first.", file=sys.stderr)
        print(f"  Expected at: {TOKEN_PATH}", file=sys.stderr)
        print(file=sys.stderr)
        print("Setup steps (one-time):", file=sys.stderr)
        print("  1. Create OAuth credentials at https://console.cloud.google.com/apis/credentials", file=sys.stderr)
        print("  2. Enable the Google Sheets API", file=sys.stderr)
        print(f"  3. Run: python {Path(__file__).parent.parent.parent / 'google-workspace' / 'scripts' / 'setup.py'} --check", file=sys.stderr)
        sys.exit(1)

    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH))
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_PATH.write_text(json.dumps(json.loads(creds.to_json()), indent=2))
    if not creds.valid:
        print("ERROR: Token is invalid. Re-run google-workspace setup.", file=sys.stderr)
        sys.exit(1)
    return creds


def get_service():
    from googleapiclient.discovery import build
    return build("sheets", "v4", credentials=get_credentials())


# ─── Formula Analysis ────────────────────────────────────────────────────────

# Match cell references like A1, $A$1, A$1, $A1, Sheet2!A1, 'Sheet 2'!A1
CELL_REF_RE = re.compile(
    r"(?:(?P<sheet>(?:'[^']+')|[A-Za-z_][\w.]*)!)?"  # optional sheet name
    r"(?P<ref>\$?[A-Z]+\$?\d+)"                       # cell like $A$1 or B2
)

# Match ranges like A1:B5
RANGE_RE = re.compile(
    r"(?:(?P<sheet>(?:'[^']+')|[A-Za-z_][\w.]*)!)?"
    r"(?:\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)"
)

# Google Sheets functions to identify
FUNCTIONS = {
    "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "IF", "IFS",
    "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "FILTER",
    "QUERY", "ARRAYFORMULA", "ARRAY_CONSTRAIN", "SORT", "UNIQUE",
    "FLATTEN", "SPLIT", "JOIN", "CONCATENATE", "TEXT", "TO_TEXT",
    "DATE", "TIME", "NOW", "TODAY", "YEAR", "MONTH", "DAY",
    "WEEKDAY", "DATEDIF", "EDATE", "EOMONTH", "NETWORKDAYS",
    "WORKDAY", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "MOD",
    "ABS", "CEILING", "FLOOR", "POWER", "SQRT", "LOG", "EXP",
    "SUBTOTAL", "CHOOSE", "OFFSET", "INDIRECT", "ADDRESS",
    "COLUMN", "ROW", "COLUMNS", "ROWS", "TRANSPOSE",
    "IMPORTRANGE", "IMPORTHTML", "IMPORTDATA", "IMPORTXML",
    "GOOGLEFINANCE", "GOOGLETRANSLATE", "IMAGE",
    "ISBLANK", "ISERROR", "ISNUMBER", "ISTEXT", "ISDATE",
    "IFERROR", "IFNA", "SWITCH", "AND", "OR", "NOT", "XOR",
    "TO_DATE", "TO_PURE_NUMBER", "TO_DOLLARS", "TO_PERCENT",
    "N", "ENCODEURL", "REGEXEXTRACT", "REGEXREPLACE", "REGEXMATCH",
    "SUBSTITUTE", "REPLACE", "LEFT", "RIGHT", "MID", "LEN",
    "FIND", "SEARCH", "TRIM", "UPPER", "LOWER", "PROPER",
    "HYPERLINK", "CELL", "TYPE", "RANK",
    "DAYS", "HOUR", "MINUTE", "SECOND", "TIMEVALUE", "DATEVALUE",
}


def parse_cell_refs(formula: str) -> list[dict]:
    """Extract cell references and ranges from a formula string."""
    refs = []
    for m in CELL_REF_RE.finditer(formula):
        refs.append({
            "ref": m.group("ref"),
            "sheet": m.group("sheet").strip("'") if m.group("sheet") else None,
        })
    return refs


def parse_functions(formula: str) -> list[str]:
    """Extract function names from a formula string."""
    return [f for f in FUNCTIONS if re.search(rf"\b{f}\s*\(", formula, re.IGNORECASE)]


def a1_to_row_col(a1: str) -> tuple[int, int] | None:
    """Convert A1 notation to 0-based (row, col). Returns None if invalid."""
    m = re.match(r"\$?([A-Z]+)\$?(\d+)$", a1)
    if not m:
        return None
    col = 0
    for c in m.group(1):
        col = col * 26 + (ord(c.upper()) - ord("A") + 1)
    return (int(m.group(2)) - 1, col - 1)


def col_to_a1(col: int) -> str:
    """Convert 0-based column index to A1 column letter."""
    result = ""
    col += 1
    while col > 0:
        col -= 1
        result = chr(col % 26 + ord("A")) + result
        col //= 26
    return result


def row_col_to_a1(row: int, col: int) -> str:
    """Convert 0-based (row, col) to A1 notation."""
    return f"{col_to_a1(col)}{row + 1}"


# ─── Sheet Analysis ──────────────────────────────────────────────────────────

def analyze_sheet(service, spreadsheet_id: str, sheet_name: str | None = None) -> dict:
    """Get full analysis of a spreadsheet — all data, formulas, structure."""
    
    # Get spreadsheet metadata
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties,sheets.protectedRanges,namedRanges",
        includeGridData=False,
    ).execute()

    result = {
        "spreadsheet_id": spreadsheet_id,
        "named_ranges": [],
        "sheets": [],
    }

    # Get named ranges
    for nr in meta.get("namedRanges", []):
        result["named_ranges"].append({
            "name": nr["name"],
            "range": nr["range"],
        })

    sheets_meta = meta.get("sheets", [])

    for sheet_meta in sheets_meta:
        props = sheet_meta["properties"]
        name = props["title"]
        if sheet_name and name != sheet_name:
            continue

        sheet_info = {
            "name": name,
            "sheet_id": props.get("sheetId"),
            "row_count": props.get("gridProperties", {}).get("rowCount", 0),
            "col_count": props.get("gridProperties", {}).get("columnCount", 0),
            "is_protected": bool(sheet_meta.get("protectedRanges")),
            "columns": [],
            "data": [],
            "formulas": {},
            "dependencies": {},
            "summary": {},
        }

        # Get the actual data with formulas
        range_str = f"'{name}'!1:{sheet_info['row_count']}"
        try:
            resp = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_str,
                valueRenderOption="FORMULA",
                majorDimension="ROWS",
            ).execute()
        except Exception as e:
            # Try without the whole-column range
            try:
                resp = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{name}'!A1:{col_to_a1(min(sheet_info['col_count'], 26) - 1)}{sheet_info['row_count']}",
                    valueRenderOption="FORMULA",
                    majorDimension="ROWS",
                ).execute()
            except Exception as e2:
                sheet_info["error"] = str(e2)
                result["sheets"].append(sheet_info)
                continue

        rows = resp.get("values", [])
        row_count = len(rows)
        max_cols = max(len(r) for r in rows) if rows else 0

        # Convert to structured grid
        grid = []
        for r in range(row_count):
            row_data = {}
            for c in range(max(max_cols, 1)):
                val = rows[r][c] if r < len(rows) and c < len(rows[r]) else ""
                cell_ref = row_col_to_a1(r, c)
                is_formula = isinstance(val, str) and val.startswith("=")
                row_data[col_to_a1(c)] = {
                    "value": val,
                    "is_formula": is_formula,
                    "references": parse_cell_refs(val) if is_formula else [],
                    "functions": parse_functions(val) if is_formula else [],
                }
                if is_formula:
                    sheet_info["formulas"][f"{name}!{cell_ref}"] = val
                    deps = [d["ref"] for d in parse_cell_refs(val)]
                    sheet_info["dependencies"][f"{name}!{cell_ref}"] = deps
            grid.append(row_data)

        sheet_info["data"] = grid

        # Identify headers (first row)
        headers = []
        if rows:
            for c in range(max_cols):
                headers.append({
                    "col": col_to_a1(c),
                    "value": str(rows[0][c]) if c < len(rows[0]) else "",
                })
        sheet_info["columns"] = headers

        # Compute summary
        formula_count = len(sheet_info["formulas"])
        # Detect column data types from non-header, non-empty rows
        col_types = {}
        if len(rows) > 1:
            for c in range(max_cols):
                non_empty = []
                for r in range(1, len(rows)):
                    if c < len(rows[r]) and rows[r][c] != "":
                        non_empty.append(rows[r][c])
                if non_empty:
                    numeric_count = 0
                    formula_count_col = 0
                    for v in non_empty:
                        if isinstance(v, str) and v.startswith("="):
                            formula_count_col += 1
                        else:
                            try:
                                float(v)
                                numeric_count += 1
                            except (ValueError, TypeError):
                                pass
                    if formula_count_col > len(non_empty) * 0.5:
                        col_types[col_to_a1(c)] = "formula"
                    elif numeric_count > len(non_empty) * 0.5:
                        col_types[col_to_a1(c)] = "number"
                    else:
                        col_types[col_to_a1(c)] = "text"

        sheet_info["summary"] = {
            "row_count": len(rows),
            "col_count": max_cols,
            "formula_count": formula_count,
            "cell_count": len(rows) * max_cols,
            "col_types": col_types,
        }

        # Build dependency graph
        dep_graph = defaultdict(list)
        for formula_cell, deps in sheet_info["dependencies"].items():
            for dep in deps:
                dep_graph[dep].append(formula_cell)
        sheet_info["reverse_dependencies"] = dict(dep_graph)

        result["sheets"].append(sheet_info)
        if sheet_name:
            break

    return result


def format_structure(result: dict) -> dict:
    """Produce a concise structural overview for the agent to understand."""
    summary = {
        "spreadsheet_id": result["spreadsheet_id"],
        "named_ranges": [nr["name"] for nr in result["named_ranges"]],
        "sheets": [],
    }

    for sheet in result["sheets"]:
        headers = [h["value"] for h in sheet.get("columns", [])]
        col_types = sheet.get("summary", {}).get("col_types", {})
        col_info = {}
        for h in sheet.get("columns", []):
            c = h["col"]
            col_info[c] = {
                "header": h["value"],
                "type": col_types.get(c, "unknown"),
            }

        # Key formulas (first few to give a sense of the logic)
        key_formulas = {}
        formula_shown = 0
        for cell, formula in list(sheet.get("formulas", {}).items())[:15]:
            short_cell = cell.split("!")[-1] if "!" in cell else cell
            key_formulas[short_cell] = formula
            formula_shown += 1

        sheet_summary = {
            "name": sheet["name"],
            "dimensions": f"{sheet['summary']['row_count']} rows x {sheet['summary']['col_count']} cols",
            "headers": headers,
            "columns": col_info,
            "formula_count": sheet["summary"]["formula_count"],
            "key_formulas": key_formulas,
            "has_more_formulas": sheet["summary"]["formula_count"] > formula_shown,
            "is_protected": sheet.get("is_protected", False),
        }

        # Add cell dependency summary
        deps = sheet.get("dependencies", {})
        if deps:
            # Cells that are most referenced by others
            rev = sheet.get("reverse_dependencies", {})
            most_referenced = sorted(rev.items(), key=lambda x: -len(x[1]))[:5]
            sheet_summary["most_referenced_cells"] = [
                {"cell": cell, "referenced_by_count": len(refs), "referenced_by": refs}
                for cell, refs in most_referenced
            ]

        summary["sheets"].append(sheet_summary)

    return summary


def get_dependencies(service, spreadsheet_id: str, sheet_name: str | None = None) -> dict:
    """Build a dependency map showing which cells reference which."""
    result = analyze_sheet(service, spreadsheet_id, sheet_name)
    deps = {}

    for sheet in result["sheets"]:
        name = sheet["name"]
        if sheet_name and name != sheet_name:
            continue
        for cell, refs in sheet.get("dependencies", {}).items():
            short_cell = cell.split("!")[-1] if "!" in cell else cell
            deps[f"{name}!{short_cell}"] = {
                "formula": sheet["formulas"].get(cell, ""),
                "references": refs,
            }

    # Also compute reverse deps (what cells are referenced BY what)
    reverse = defaultdict(list)
    for cell, info in deps.items():
        for ref in info["references"]:
            reverse[ref].append(cell)

    return {"dependencies": deps, "reverse_dependencies": dict(reverse)}


# ─── Updates ─────────────────────────────────────────────────────────────────

def update_cell(service, spreadsheet_id: str, cell_range: str, value: str) -> dict:
    """Update a single cell. If value starts with '=', it's stored as a formula."""
    body = {"values": [[value]]}
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=cell_range,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    return {
        "updated_range": result.get("updatedRange", ""),
        "updated_cells": result.get("updatedCells", 0),
        "value": value,
    }


def update_range(service, spreadsheet_id: str, range_str: str, values: list[list]) -> dict:
    """Update a range of cells."""
    body = {"values": values}
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueInputOption="USER_ENTERED",
        body=body,
    ).execute()
    return {
        "updated_range": result.get("updatedRange", ""),
        "updated_cells": result.get("updatedCells", 0),
    }


def append_rows(service, spreadsheet_id: str, range_str: str, values: list[list]) -> dict:
    """Append rows to a sheet."""
    body = {"values": values}
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()
    updates = result.get("updates", {})
    return {
        "updated_range": updates.get("updatedRange", ""),
        "updated_cells": updates.get("updatedCells", 0),
    }


def batch_update(service, spreadsheet_id: str, changes: dict[str, str]) -> list[dict]:
    """Apply multiple cell updates in one batch."""
    results = []
    for cell_range, value in changes.items():
        r = update_cell(service, spreadsheet_id, cell_range, value)
        results.append(r)
    return results


# ─── Preview ─────────────────────────────────────────────────────────────────

def format_preview(result: dict, max_rows: int = 10) -> str:
    """Format a readable table preview."""
    lines = []
    for sheet in result["sheets"][:1]:  # Preview first matching sheet
        name = sheet["name"]
        dims = f"{sheet['summary']['row_count']}x{sheet['summary']['col_count']}"
        lines.append(f"Sheet: {name}  ({dims}, {sheet['summary']['formula_count']} formulas)")
        lines.append("")

        # Build ascii table
        max_cols = sheet["summary"]["col_count"]
        if max_cols == 0:
            lines.append("  (empty)")
            continue

        headers = [h["value"] for h in sheet.get("columns", [])]
        col_widths = [max(len(str(h)), 4) for h in headers]

        # Read rows for preview
        grid = sheet.get("data", [])
        display_rows = min(max_rows + 1, len(grid))  # +1 for header

        # Measure content widths
        for r in range(min(display_rows, len(grid))):
            row = grid[r]
            for c_idx in range(min(max_cols, len(col_widths))):
                col_letter = col_to_a1(c_idx)
                val = str(row.get(col_letter, {}).get("value", ""))
                col_widths[c_idx] = max(col_widths[c_idx], min(len(val), 30))

        # Truncate widths to reasonable max
        col_widths = [min(w, 25) for w in col_widths]

        # Build separator
        sep = "+-" + "-+-".join("-" * w for w in col_widths) + "-+"

        # Header row
        header_cells = []
        for c_idx, h in enumerate(headers[:len(col_widths)]):
            header_cells.append(h.ljust(col_widths[c_idx]))
        lines.append(sep)
        lines.append("| " + " | ".join(header_cells) + " |")
        lines.append(sep)

        # Data rows
        for r in range(1, min(display_rows, len(grid))):
            row = grid[r]
            cells = []
            for c_idx in range(min(max_cols, len(col_widths))):
                col_letter = col_to_a1(c_idx)
                cell = row.get(col_letter, {})
                val = str(cell.get("value", ""))
                is_formula = cell.get("is_formula", False)
                # Truncate long values
                if len(val) > 25:
                    val = val[:22] + "..."
                # Mark formulas with =
                if is_formula and not val.startswith("="):
                    val = "=" + val
                cells.append(val.ljust(col_widths[c_idx]))
            lines.append("| " + " | ".join(cells) + " |")
            lines.append(sep)

        remaining = len(grid) - display_rows
        if remaining > 0:
            lines.append(f"  ... and {remaining} more rows")

        # Show formulas summary
        formulas = sheet.get("formulas", {})
        if formulas:
            lines.append("")
            lines.append("  Key formulas:")
            for cell, formula in list(formulas.items())[:8]:
                short = cell.split("!")[-1] if "!" in cell else cell
                lines.append(f"    {short}: {formula}")
            if len(formulas) > 8:
                lines.append(f"    ... and {len(formulas) - 8} more formulas")

    return "\n".join(lines)


# ─── CLI ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Google Sheets Intelligence")
    sub = parser.add_subparsers(dest="command", required=True)

    # analyze
    p_analyze = sub.add_parser("analyze", help="Full spreadsheet analysis")
    p_analyze.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_analyze.add_argument("--sheet", "-s", help="Sheet name (default: all)")
    p_analyze.add_argument("--pretty", action="store_true", help="Pretty-print JSON")

    # structure
    p_struct = sub.add_parser("structure", help="Structural overview (lightweight)")
    p_struct.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_struct.add_argument("--sheet", "-s", help="Sheet name")

    # update
    p_upd = sub.add_parser("update", help="Update a single cell")
    p_upd.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_upd.add_argument("range", help="Cell range (e.g. 'Sheet1!A1')")
    p_upd.add_argument("value", help="New value (prefix with '=' for formula)")

    # update-range
    p_ur = sub.add_parser("update-range", help="Update a range of cells")
    p_ur.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_ur.add_argument("range", help="Range (e.g. 'Sheet1!A1:C3')")
    p_ur.add_argument("values", help="JSON array of arrays, e.g. '[[1,2],[3,4]]'")

    # append
    p_app = sub.add_parser("append", help="Append rows")
    p_app.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_app.add_argument("range", help="Range for columns (e.g. 'Sheet1!A:C')")
    p_app.add_argument("values", help="JSON array of arrays for new rows")

    # batch
    p_batch = sub.add_parser("batch", help="Batch update multiple cells")
    p_batch.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_batch.add_argument("changes", help='JSON dict: {"Sheet1!A1": "value", ...}')

    # preview
    p_prev = sub.add_parser("preview", help="Terminal-friendly table preview")
    p_prev.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_prev.add_argument("--sheet", "-s", help="Sheet name")
    p_prev.add_argument("--rows", type=int, default=10, help="Max rows to show")

    # named-ranges
    p_nr = sub.add_parser("named-ranges", help="List named ranges")
    p_nr.add_argument("spreadsheet_id", help="Google Sheets ID")

    # dependencies
    p_dep = sub.add_parser("dependencies", help="Show cell formula dependencies")
    p_dep.add_argument("spreadsheet_id", help="Google Sheets ID")
    p_dep.add_argument("--sheet", "-s", help="Sheet name")

    args = parser.parse_args()

    # Install deps if missing
    try:
        get_service()
    except ImportError:
        print("Installing required packages...", file=sys.stderr)
        import subprocess
        subprocess.check_call([
            sys.executable, "-m", "pip", "install",
            "google-api-python-client",
            "google-auth-oauthlib",
            "google-auth-httplib2",
        ])
        print("Done. Retrying...", file=sys.stderr)

    service = get_service()

    if args.command == "analyze":
        result = analyze_sheet(service, args.spreadsheet_id, args.sheet)
        indent = 2 if args.pretty else None
        print(json.dumps(result, indent=indent, default=str, ensure_ascii=False))

    elif args.command == "structure":
        result = analyze_sheet(service, args.spreadsheet_id, args.sheet)
        summary = format_structure(result)
        print(json.dumps(summary, indent=2, default=str, ensure_ascii=False))

    elif args.command == "update":
        result = update_cell(service, args.spreadsheet_id, args.range, args.value)
        print(json.dumps(result, indent=2, ensure_ascii=False))

    elif args.command == "update-range":
        values = json.loads(args.values)
        result = update_range(service, args.spreadsheet_id, args.range, values)
        print(json.dumps(result, indent=2, ensure_ascii=False))

    elif args.command == "append":
        values = json.loads(args.values)
        result = append_rows(service, args.spreadsheet_id, args.range, values)
        print(json.dumps(result, indent=2, ensure_ascii=False))

    elif args.command == "batch":
        changes = json.loads(args.changes)
        results = batch_update(service, args.spreadsheet_id, changes)
        print(json.dumps(results, indent=2, ensure_ascii=False))

    elif args.command == "preview":
        result = analyze_sheet(service, args.spreadsheet_id, args.sheet)
        print(format_preview(result, max_rows=args.rows))

    elif args.command == "named-ranges":
        meta = service.spreadsheets().get(
            spreadsheetId=args.spreadsheet_id,
            fields="namedRanges",
        ).execute()
        print(json.dumps(meta.get("namedRanges", []), indent=2, default=str, ensure_ascii=False))

    elif args.command == "dependencies":
        deps = get_dependencies(service, args.spreadsheet_id, args.sheet)
        print(json.dumps(deps, indent=2, default=str, ensure_ascii=False))


if __name__ == "__main__":
    main()
