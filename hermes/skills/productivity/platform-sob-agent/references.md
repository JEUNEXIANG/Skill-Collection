# References — Platform SOB Agent

## Business Concepts

### SOB — Share of Business
SOB measures Shopee's (SHP) share relative to TikTok Shop (TTS) for a given metric.

**Formula:** `SOB = SHP value / TTS value`

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
- Contains ADG and ADO SOB values for all sites, stacked vertically by site/region.
- The "red section" = cells highlighted in red background color → these are the SOB values to be copied.
- Structure detected dynamically by the agent using `$GSI scan-sections` on each run.

### [Weekly Live View] — Tab: `SHP/TTS Clusters`
- Contains cluster-level data for each site.
- Columns E to S are copied by the agent (dynamic last row).

### [Reg Commercial Team] — Tab: `(Final) Data from CF excel`
- Contains PC2% values (TR% − CIR% − Transaction Fee) for all sites.
- Full tab is archived dynamically (last row and column detected automatically).

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
| `[Weekly Live View]` | Copy of (Weekly Live View) SHP vs TTS ADG ADO Absolutes | `1VjeWOSX6nX_oQiU8QnB98suG2siKovuJhoBzxTK-zeI` |
| `[Reg Commercial Team]` | Copy of [Reg Commercial Team] SOB ADG and ADO | `10kH9Welrxx7KJEOrtfWFshOrglsG-P6vTv9otwxsPho` |
| `[Reg CNLS copy]` | Copy of (Reg CNLS copy) SHP vs TTS ADG ADO Absolutes | `1cN29heWI-7trzBLMvEpznXslDmlCEmLHUcKDjT9uTqg` |
| `[Archive]` | Copy of [Reg CNLS] Archive Platform SOB & PC2 | `1F99kNADGaRxiuxkxG2Gq3A7Xvoh_bG0EM2C0xYi8Ocs` |

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
