---

## Verifier Implementation Pitfalls

These are lessons learned from implementing the verifier as a Python script against live Google Sheets data.

### ⚠️ S2-3: Header Row Is Not Row 1

The ADG ADO tab in `[Reg CNLS copy]` has:
- **Row 1:** Title/description
- **Row 2:** Section labels (e.g. `SHP ADG`)
- **Row 3:** Actual column headers with market names (`Period`, `SG`, `MY`, `TH`, ...)

When checking market column alignment, read **row 3** — not row 1 or row 2. Checking row 1 will always fail since it's a title.

### ⚠️ Tab Ordering: Newer = Lower Index

Tabs are ordered chronologically with the **newest tab having the lowest index**. Use the `index` property from `sheets.properties` for `newest_tab()` logic. Do NOT sort by sheetId — sheetIds are arbitrary and not monotonic with position.

### ⚠️ Date: Use Orchestrator-Provided Date, Never Infer from Sheet

The verifier receives `today_date` from the orchestrator. Do NOT infer today's date from the sheet (e.g. by looking at the most recent ADG ADO tab). They may differ — the most recent tab could be from a previous run, while the verifier runs on a different day. Also use `Asia/Shanghai` timezone.

### ⚠️ Cascade Logic: SKIP vs FAIL vs PASS

When a prerequisite step's data doesn't exist (e.g. today's ADG ADO tab wasn't created), downstream steps should report `SKIP` — not `FAIL` or `PASS`. A step that never ran is not a failure of the executor's work.

### ⚠️ Report Existence, Not Counts

For Step 6, check whether a properly-named `YY/MM/DD By cluster` tab exists. Do not report the total count of cluster tabs — it's misleading.

### ⚠️ Step 3 Cross-Check: Match by (Market, Month), Not Row

The cross-check section in `[Reg CNLS copy]` (rows 172:196) and its source in `[Reg Commercial Team]` have different row layouts. Compare by building a `(market, month) → value` lookup from both sources, then assert equality — never compare by absolute row number.
