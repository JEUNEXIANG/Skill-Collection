---
trigger: weekly, Thursday, 13:00
timezone: Asia/Shanghai
fallback_if_missed: run on next available slot, notify user via WeChat
---

# Platform SOB Update Agent

## Work Definition
Update Platform SoB & PC2 from various documents and archive them every Thursday at 1pm (Asia/Shanghai).

---

## Tools

- `google-sheets-intelligence` (SKILL.md at `/Users/apple/.hermes/skills/productivity/google-sheets-intelligence/`)
  - Used for: reading sheet structure, detecting stacked table sections, reading cell values/formulas, reading cell background colors (via raw Sheets API pattern), writing values (paste-as-value), duplicating tabs, and validation checks.
  - For background color detection (e.g. the "red section"): use the raw Sheets API `effectiveFormat.backgroundColor` pattern documented in SKILL.md §"Reading cell background colors".
  - For all row/column range resolution: agent must read the table structure dynamically using `$GSI scan-sections` and `$GSI structure`, confirm the reading with the user before proceeding.

---

## Registries

> Sheet IDs and tab inventory are maintained in `references.md` § Sheet Registry and § Tab Registry.
> Load `references.md` before executing any step to resolve workbook aliases (e.g. `[Weekly Live View]`) to their Google Sheet IDs.

---

## Permission Model

**ALL steps require human approval before execution.**

For every step, the agent must:
1. Read the relevant data and prepare the intended action.
2. Present a clear summary of what it is about to do (cells to be written, ranges, tab names, values).
3. Wait for the user to visually verify and confirm before executing the write.

Approval channel: **WeChat message** — send a summary and wait for reply.
Uncertainty escalation: if the agent cannot resolve an ambiguity after 1 retry, halt, log the issue, and notify via WeChat.

> Harness implementation details (audit trail, permission gate, sandbox mode, outcome evaluation) are in `harness.md`.

---

## Workflow

### Pre-flight: Validate Source Tabs

Before executing any step:
1. Load `references.md` to resolve all sheet aliases and tab names.
2. Confirm all required tabs exist and are accessible.
   - Command: `$GSI structure SPREADSHEET_ID --sheet "TAB_NAME"` for each tab.
   - If any tab is missing: **halt** and notify via WeChat: `"Pre-flight failed: tab [TAB_NAME] not found in [WORKBOOK_ALIAS]"`.
3. Run `$GSI scan-sections` on each source tab to detect stacked table sections. Present the section map to the user for visual confirmation before proceeding.

---

### Platform SOB

#### Step 1 — Read & Validate Source Data Structure

Read the table structure of the following tabs using the google-sheets-intelligence skill:

1. Tab `SHP/TTS ADG ADO` in `[Weekly Live View]`
2. Tab `SHP/TTS Clusters` in `[Weekly Live View]`
3. Tab `(Final) Data from CF excel` in `[Reg Commercial Team]`

- Use `$GSI scan-sections` to detect all stacked table sections (multiple site/region tables stacked vertically).
- Present the detected structure to the user and ask: *"Does this look correct? Shall I proceed?"*
- **Wait for user confirmation before proceeding to Step 2.**

---

#### Step 2 — Copy SOB Values (ADG ADO) to New Duplicated Tab in [Reg CNLS copy]

1. Read data from tab `SHP/TTS ADG ADO` of `[Weekly Live View]` (or fallback `Sheet6`/`Sheet5` if IMPORTRANGE is broken).
2. In `[Reg CNLS copy]`, find the tab with the most recent date (format: `YY/MM/DD ADG ADO`).
3. Duplicate that tab. Name the new tab: `{YY}/{MM}/{DD} ADG ADO` using **today's date** (Thursday, Asia/Shanghai).
   - Date format: `YY/MM/DD` — e.g. `26/05/14` = year 2026, month 05, day 14.
   - Place the new tab in **chronological date order** among existing tabs.
   - If a tab with today's date already exists: **do NOT overwrite**. Log warning, notify via WeChat, and halt.
4. Present the full intended paste plan (source range → destination range; layout preserved: ADO↔ADO, ADG↔ADG, SHP↔SHP, TTS↔TTS).
5. **Wait for user confirmation**, then paste the source values **as values only** (no formulas).

---

#### Step 3 — Paste SOB Values Within [Reg CNLS copy] Tab

In the newly duplicated tab (`{YY}/{MM}/{DD} ADG ADO`) of `[Reg CNLS copy]`:

**Goal:** Copy the SOB calculation section into the results section, then bring in the [Reg Commercial Team] reference values for cross-checking.

1. Use `$GSI scan-sections` to read the full table structure and identify:
   - The **SOB calculation section** (expected around rows `98:130` — verify by checking section headers/labels).
   - The **SOB results section** (expected around rows `137:169` — verify similarly).
   - Present both detected sections to the user: *"I found the calculation section at [ACTUAL RANGE] and the results section at [ACTUAL RANGE]. Does this match your expectation? Shall I copy calculation → results as values?"*
2. **Wait for user confirmation**, then paste the calculation section → results section **as values only**.
3. Read tab `(Final) Data from CF excel` in `[Reg Commercial Team]` using `$GSI scan-sections` to identify the reference values section (expected around rows `172:196` in the destination — verify by section headers).
   - Present to user: *"I will copy [DETECTED RANGE] from [Reg Commercial Team] SHP & TTS PC2 and paste as values into the cross-check section at [ACTUAL RANGE]. Does this look correct?"*
4. **Wait for user confirmation**, then paste **as values only**.

---

#### Step 4 — Validate Zero Check (Difference Table)

In the newly duplicated tab (`{YY}/{MM}/{DD} ADG ADO`) of `[Reg CNLS copy]`:

**Goal:** Confirm that the SOB values pasted in Step 3 match the reference values, by checking the difference table is entirely zero.

1. Use `$GSI scan-sections` to identify the **difference/check table** (expected around rows `199:223` — verify by looking for a section that subtracts or compares the two value sets).
   - Report the actual detected range to the audit log.
2. Check: are ALL values in the detected section equal to `0`?
   - ✅ **If yes:** notify user via WeChat — *"Step 4 zero-check passed. All difference values are 0. Proceeding to Step 5."* Wait for confirmation.
   - ❌ **If any value ≠ 0:** **halt**. Do NOT proceed to Step 5. Send WeChat alert:
     ```
     ⚠️ Step 4 Validation Failed
     Non-zero differences found — archive step will be skipped:
     - [SITE] [MONTH] [METRIC] (e.g. ID 2026-May ADG SOB): [CELL_REF] = [VALUE]
     Please check the sheet and advise.
     ```

---

#### Step 5 — Archive SOB to [Archive] (Conditional)

**Goal:** Archive the finalised SOB values only when both the difference check and the cross-source value match confirm the data is correct.

**Conditions (both must be true before proceeding):**
- Step 4 zero-check passed.
- The SOB results section in `[Reg CNLS copy]` (detected in Step 3, expected around rows `137:169`) matches **exactly** (whole number, no tolerance) with the corresponding values from `[Reg Commercial Team]`.
  - Agent reads both ranges dynamically and presents a side-by-side comparison table to the user for visual confirmation.

**Action if both conditions met:**
1. In `[Archive]`, identify all tabs and filter to **SOB tabs only** (format: `SOB-YYMMDD`). Find the one with the most recent date.
   - Note: `[Archive]` also contains PC2 tabs (`PC2-YYMMDD`) — ignore those here.
2. Duplicate the most recent SOB tab. Name the new tab: `SOB-{YYMMDD}` using today's date (e.g. `SOB-260514`).
   - Place in **chronological date order**. If today's tab already exists: **do NOT overwrite** — halt and notify via WeChat.
3. Read the destination tab structure to identify the correct paste target.
   - Present paste plan to user: *"I will paste the SOB results section [ACTUAL RANGE] from [Reg CNLS copy] into [ACTUAL RANGE] of SOB-260514. Does this look correct?"*
4. **Wait for user confirmation**, then paste **as values only**.

---

#### Step 6 — Copy Clusters Data to [Reg CNLS copy]

1. From tab `SHP/TTS Clusters` in `[Weekly Live View]`:
   - Copy columns `E` to `S`, all data rows **dynamically** (detect last non-empty row automatically).
2. In `[Reg CNLS copy]`, find the most recent `By cluster` tab (format: `YY/MM/DD By cluster`).
3. Duplicate that tab. Name the new tab: `{YY}/{MM}/{DD} By cluster` using today's date.
   - Place in **chronological date order**. If today's tab already exists: **do NOT overwrite**, halt and notify.
4. Read the destination tab structure dynamically to identify the correct paste location.
   - Present paste plan to user for visual confirmation.
5. **Wait for user confirmation**, then paste **as values only** (dynamic last row and column).

---

### PC2

#### PC2 Step 1 — Archive PC2 to [Archive]

1. From tab `(Final) Data from CF excel` in `[Reg Commercial Team]`:
   - Read all values **dynamically** (detect last non-empty row and column, starting from `A1`).
2. In `[Archive]`, find the **PC2 tab with the most recent date** (format: `PC2-YYMMDD`).
3. Duplicate that tab. Name the new tab: `PC2-{YYMMDD}` using today's date (e.g. `PC2-260514`).
   - Place in **chronological date order**. If today's tab already exists: **do NOT overwrite**, halt and notify.
4. Read the destination tab structure dynamically to identify the correct paste location.
   - Present paste plan to user for visual confirmation.
5. **Wait for user confirmation**, then paste all values **as values only**.
