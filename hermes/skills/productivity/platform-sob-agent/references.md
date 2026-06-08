# References — Platform SOB Agent

## Business Concepts

### SOB — Share of Business
SOB measures Shopee's (SHP) share relative to TikTok Shop (TTS) for a given metric.

**Formula:** `SOB = SHP metric value / TTS metric value` for the same metric and market/site.

There are two types of SOB:
| Type | Full Name | Definition |
|---|---|---|
| **ADG SOB** | Average Daily GMV SOB | Shopee average daily GMV ÷ TikTok Shop average daily GMV |
| **ADO SOB** | Average Daily Orders SOB | Shopee average daily number of orders ÷ TikTok Shop average daily number of orders |

### Platform SOB
The SOB calculated across **all sellers** on a given platform for a specific site.
- Distinct from seller-level or category-level SOB.

---

### Sites

| Code | Market |
|---|---|
| ID | Indonesia |
| SG | Singapore |
| TH | Thailand |
| VN | Vietnam |
| PH | Philippines |

---

### Financial Metrics

| Metric | Full Name | Definition |
|---|---|---|
| **GMV** | Gross Merchandise Value | Total value of goods sold through the platform |
| **TR%** | Take Rate % | Platform revenue as a % of total GMV |
| **CIR%** | Cost-to-Income Ratio % | Platform investment into sellers (rebates, discounts, coupons) as a % of total GMV |
| **Transaction Fee** | — | Platform's operational cost per transaction |
| **PC2%** | — | `TR% − CIR% − Transaction Fee` — the remaining revenue after seller investment and transaction costs |

---

### Data Freshness
- Source data in `[Weekly Live View]` is updated weekly (typically by Wednesday).
- The agent runs on **Thursday at 1pm Asia/Shanghai** to capture the most recent weekly data.
- If source tabs appear stale (no change since last week), the agent should flag this to the user before proceeding.

---

## Sheet Structure Reference

### [Weekly Live View] — Tab: `SHP/TTS ADG ADO`

**Invariant identity** — the **raw ingredient table** for Platform SOB calculation. Contains platform-level (all-seller) ADG (Avg Daily GMV) and ADO (Avg Daily Orders) **absolute values** for both SHP (Shopee) and TTS (TikTok Shop), **not precomputed SOB ratios**.

**How to recognize it (detect dynamically on every run):**
- **Two stacked vertical sections** — an ADG section (always above) and an ADO section (always below), separated by a blank row
- **Section headers:** Look for cells containing keywords in this order: `"SHP ADG"`, `"TTS ADG"`, then later `"SHP ADO"`, `"TTS ADO"`. These are typically in a header row spanning 2 rows (a category label row + a column header row)
- **Column header patterns:** The column header row contains `"Period"` + market codes (`SG`, `MY`, `TH`, `ID`, `VN`, `PH`, `SEA excl TW`, `SEA excl TW ID`) — this same market sequence repeats for each sub-table (SHP values, TTS values, SOB Multiple ratio)
- **Data rows:** Monthly periods from `Jan'24` onward, with period labels in column B, dates in column C, and later months optionally marked `(Proj.)` or `(Target)`
- **Within each section (ADG or ADO),** there are 3 side-by-side sub-tables with identical column layouts:
  1. **SHP absolute values** (typically leftmost)
  2. **TTS absolute values** (typically center)
  3. **TTS ADG/ADO Multiple** (SOB ratio = SHP/TTS, typically rightmost)

**Selection instruction:** Match when you need raw, platform-level ADG or ADO absolute values of SHP or TTS — the ingredients from which Platform SOB is computed. Do NOT match if the data is pre-computed SOB ratios, cluster-level breakdowns, or PC2 metrics.

**Structure detected dynamically** using `$GSI scan-sections` on each run — never assume fixed row/column numbers.

### [Weekly Live View] — Tab: `SHP/TTS Clusters`
- Contains cluster-level ADG & ADO data for each site.
- Table description for matching: cluster/category-level SHP/TTS ADG & ADO absolutes table (not SOB) split by market and commercial cluster, such as Overall, EL, LS, FMCG excl HB, Fashion, and HB.
- Selection instruction: when the user asks for cluster, category, segment, or SOB by cluster data, match to this table if the detected section headers/labels include cluster/category names and ADG/ADO values.
- Columns E to S are copied by the agent (dynamic last row).

### [Reg Commercial Team] — Tab: `(Final) Data from CF excel`

**Invariant identity** — the **cross-check reference table** for SHP/TTS ADG SOB & ADO SOB values. Contains pre-computed SOB ratios (called "ADG Multiple" / "ADO Multiple") from the Regional Commercial Team, used to validate that the agent's SOB calculations match the commercial team's numbers.

**How to recognize it (detect dynamically on every run):**
- **Section header:** Look for a cell containing `"SHP/TTS ADG Multiple"` — this marks the start of the only section relevant to this workflow
- **Column header pattern:** After the section header row, the next row contains market codes (`SG`, `MY`, `TH`, `ID`, `VN`, `PH`, `SEA excl TW`, `SEA excl TW ID`) repeated twice — first for ADG Multiple, then for ADO Multiple
- **Data rows:** Monthly periods starting from `2024-12-01` with approximately 25 rows of data (through projected/target months). Values display with `"x"` suffix (e.g. `"6.26x"`) and must be stripped to numeric when read
- **Only the SHP/TTS section** is needed for the workflow — other sections (SHP/LZD, transposed views) are irrelevant

**Selection instruction:** Match only when you need the commercial team's pre-computed SHP/TTS ADG SOB & ADO SOB ratios for cross-check validation. This is the reference that gets pasted into rows ~172:196 of the `[Reg CNLS copy]` working tab.

**Structure detected dynamically** using `$GSI scan-sections` on each run — never assume fixed row/column numbers.

### [Platform PC2] — Tab: `SHP & TTS PC2`

**Invariant identity** — the **SHP PC2 economics table**. Contains SHP (Shopee) PC2% values and related revenue/cost metrics organized in a wide horizontal layout (93 columns), with monthly data for each market.

**How to recognize it (detect dynamically on every run):**
- **Section label:** Look for `"SHP Bottomline"` or `"SHP Data"` near the top — the entire tab is SHP-side data
- **Horizontally stacked metrics:** Multiple metric groups side by side, each with 9 market columns (SG, MY, TH, VN, PH, ID, Regional excl ID and SG, Regional excl SG, Regional incl SG). Key metric groups include:
  1. **SHP PC2%** — the primary metric (TR% − CIR% − Transaction Fee)
  2. **SHP MP-only PC2%** — marketplace-only PC2
  3. Additional PC2 and revenue metrics further to the right
- **Data rows:** Monthly periods from approximately 2026-12 backwards to 2024-01
- **Additional vertical sections** below the main PC2 section (e.g. MP Revenue%, SHP MP-only TR%) — separated by blank rows

**Selection instruction:** Match when you need SHP PC2% values, take rate, cost-to-income, or platform economics data. Note: TTS PC2 data is NOT in this tab — it is in a separate tab or workbook.

**Structure detected dynamically** using `$GSI scan-sections` on each run — never assume fixed row/column numbers.

### Table Matching Guidance

During Pre-flight validation, use `$GSI scan-sections` output as the evidence for table selection. Compare the user's description against:
- workbook alias and tab name
- section title or nearby labels
- header names
- market/site labels
- metric labels such as ADG, ADO, SOB, PC2, TR%, CIR%, and Transaction Fee
- expected row/column shape from this reference

If more than one detected table could match the user's description, present the candidate tables with the matching evidence and ask the user to confirm before reading, copying, pasting, or archiving data.

### [Reg CNLS copy] — Working Tabs
| Tab Format | Content |
|---|---|
| `YY/MM/DD ADG ADO` | Weekly SOB working copy (e.g. `26/05/14 ADG ADO`) |
| `YY/MM/DD By cluster` | Weekly cluster data copy (e.g. `26/05/14 By cluster`) |

### [Archive] — Archive Tabs
| Tab Format | Content |
|---|---|
| `SOB-YYMMDD` | Archived SOB values (e.g. `SOB-260514`) |
| `PC2-YYMMDD` | Archived PC2 values (e.g. `PC2-260514`) |

All tabs within each workbook are ordered **chronologically by date**.

---

## Sheet Registry

> Fill in the Google Sheet IDs once available. Find the ID in the sheet URL:
> `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`

| Alias | Full Workbook Name | Google Sheet ID |
|---|---|---|---|
| `[Weekly Live View]` | Copy of (Weekly Live View) SHP vs TTS ADG ADO Absolutes | `1m4bZ11zDEHpVzQP1I4zxtFKs-vzXEYBk7NgMNJK5dpo` (⚠️ Updated 2026-05-28 — old ID `1VjeWOSX6nX_oQiU8QnB98suG2siKovuJhoBzxTK-zeI` kept as fallback) |
| `[Reg Commercial Team]` | Copy of [Reg Commercial Team] SOB ADG and ADO | `1BmRS6VjIP5_RRfQs22pgm9Ap49G_EBZJIjJDCAhmAcg` (⚠️ Updated 2026-05-28 — old ID `10kH9Welrxx7KJEOrtfWFshOrglsG-P6vTv9otwxsPho` kept as fallback) |
| `[Reg CNLS copy]` | Copy of (Reg CNLS copy) SHP vs TTS ADG ADO Absolutes | `1cN29heWI-7trzBLMvEpznXslDmlCEmLHUcKDjT9uTqg` |
| `[Archive]` | Copy of [Reg CNLS] Archive Platform SOB & PC2 | `1F99kNADGaRxiuxkxG2Gq3A7Xvoh_bG0EM2C0xYi8Ocs` |
| `[Platform PC2]` | [Reg Commercial Team] Platform PC 2 Data | `11Qqg42jx_jAhfmkjr8JghVa4zwiVXdAsuo3aOGxvVKo` (⚠️ Updated 2026-05-28 — was `TODO`) |

---

## Workflow Row Sections (in `[Reg CNLS copy]` ADG ADO Tab)

| Row Range | Purpose |
|---|---|
| `98:130` | Source SOB calculated values (to be copied down) |
| `137:169` | Destination for SOB paste (must match `[Reg Commercial Team]` values for archiving) |
| `172:196` | Values pasted from `[Reg Commercial Team]` for cross-check |
| `199:223` | Difference check table — must all equal `0` for archiving to proceed |

---

## Skill Reference

- **google-sheets-intelligence** at `/Users/apple/.hermes/skills/productivity/google-sheets-intelligence/`
  - `$GSI structure SPREADSHEET_ID --sheet "TAB"` — get table structure
  - `$GSI scan-sections SPREADSHEET_ID --sheet "TAB"` — detect stacked table sections
  - `$GSI preview SPREADSHEET_ID --rows N` — preview data
  - `$GSI update-range SPREADSHEET_ID "TAB!A1:Z100" '[[...]]'` — write values
  - Raw Sheets API (Python) — used for color detection and tab duplication
