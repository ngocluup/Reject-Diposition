---
name: phi-manage
description: "Pull, filter, and email Intel TCI tcitools.intel.com forecast reports (SDA weekly, SDA monthly, MOR, MOR Spread, ENG) for any product CommonName. Use when: user asks for TCI reports, PHI parameters, PHI tracking, SDA forecast, MOR forecast, MORSpread, ENG forecast, PORSDAForecast, PORMonthlyForecast, MORForecasts, ENGForecasts, or mentions any Intel product codename (e.g., 'ARL Refresh S', 'PTL U404', 'Twin Lake', 'Catlow', 'ADL P282', 'RPL P282') paired with a report keyword (SDA, MOR, ENG). Supports per-class/operation filtering (Class=test, +PPV, yield) and column relabeling (WW->MMM YYYY for monthly, ATRev+N->YYYYQQ for MOR)."
argument-hint: "Specify product codename + report + optional filter. Examples: 'ARL Refresh S SDA', 'PTL U404 SDA Class=test', 'TWL SDA Monthly yield', 'Catlow MOR'"
---

# TCI PHI Tracking Skill -- Intel tcitools.intel.com automation

End-to-end PowerShell automation for fetching Intel **TCI forecast reports** (PORSDAForecast / PORMonthlyForecast / MORForecasts / MORSpread / ENGForecasts), filtering them, building formatted xlsx, and emailing to the user.

---

## Getting Started

### Prerequisites
- **Windows** with PowerShell 5.1+
- **TCI Access**: AGS entitlement "TCI - Test" (see Access Prerequisite below)
- **Microsoft Office** (Excel + Outlook) installed locally (COM automation)
- **Network**: Intel corpnet or VPN (Windows Integrated Auth to tcitools.intel.com)

### Package Contents

| File | Purpose |
|---|---|
| `SKILL.md` | Agent instructions (this file) |
| `tci_lib.ps1` | Shared library -- all helpers, single source of truth |
| `tci_single.ps1` | Universal parameterized runner |
| `tci_wif_template.ps1` | WIF (Work In Flight) submission template |
| `tci_run.ps1` | Batch runner wrapper |
| `tci_batch_driver.ps1` | Multi-product batch driver with parallel prefetch |
| `tci_commonname_lookup.json` | Product-to-group/subgroup mapping (286 products) |
| `tci_verified_log.md` | Verified row counts and regression examples |
| `tci_run_history.csv` | Run history log (auto-appended) |

### Quick Start
```
powershell -ExecutionPolicy Bypass -File "tci_single.ps1" -Product 'ARL Refresh S 8C+16A+GT1' -Report SDA -Filter Class
powershell -ExecutionPolicy Bypass -File "tci_run.ps1" -Product 'PTL U404' -Report MOR -Filter PPV
```

---

## When to Use

- User asks for a TCI report on any Intel CPU/GPU product codename
- Shorthand requests like `<product> + <report>` (e.g., `ARL Refresh S SDA`, `Catlow SDA Monthly Class=test`)
- "Give me PHI parameter values for X", "yield report for Y", "MOR for Z"
- Mentions: PHI, PHI tracking, SDA, MOR, ENG, MORSpread, PORSDAForecast, PORMonthlyForecast

---

## Access Prerequisite

If you detect HTTP 401/403 or empty response, **stop** and display:

> **You do not have TCI access.** Apply for **"TCI - Test"** in AGS:
> https://ags.intel.com/identityiq/lcm/requestAccess.jsf

---

## 0. Quick Reference

**Reports** (S1): `SDA`->`PORSDAForecast` . `SDA Monthly`->`PORMonthlyForecast` (relabel WW->MMM YYYY) . `MOR`->`MORSpread` (default) . `ENG`->`ENGForecasts`. Bare product (no report) -> **ASK** which report(s).

**Filters** (S2) -- all client-side post-filters on the data table:

| Token (+ synonyms) | Predicate |
|---|---|
| `Class` / `CH` / `CLASSHOT` | `op startswith TEST` **OR** (blank op AND metric IN MPS/EQA/CS MONITOR) |
| `PPV` | `op startswith PPV` **OR** (blank op AND metric = `PPV-M SAMPLE SIZE`) |
| `PPVs` / `PPVm` | `op = PPV_SPS` / `op = PPV_SPM` (single op only) |
| `BI` / `Burn-In` / `LCBI` / `HDBI` | `op startswith BURNIN` (no monitors) |
| `yield` | blank op AND metric IN (`U/D`,`R/D`,`FINISH YIELD`) |
| `+ <op name>` | `op startswith <name>` (case-insensitive) |
| (no filter given) | **ASK** (Class/PPV/BI/Yield/All/None) |

**WIF is a SEPARATE 4th question** (not a filter option). After variant/report/operation, ask **"Need WIF? Yes / No"**. No -> view current POR. Yes -> collect spec/range/value.

**Site/DLCP codes**: `VNAT`=SS, `PGAT`=PG8, `CD`/`CDAT`=CD6. DLCP miss -> fall back to blank row.

**Deliver**: multi-tab xlsx (Excel COM bulk array) + HTML email card (`Build-PhiCard`/`Build-WifCard` from `tci_lib.ps1`) to current user. Subject: `PHI of <product> - <report> <filter>`.

**Fast path**: `tci_run.ps1 -Product '<CommonName>' -Report '<report>' [-Filter <token>] [-Email <addr>]`. Accepts multiple products. Every run appends to `tci_run_history.csv` via `Add-PhiRunHistory` (warns on >20% row drop).

---

## 1. Report URL Catalog

All reports: `https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=<ID>`

| Shorthand | Report ID | Time axis | First `<th>` | Notes |
|---|---|---|---|---|
| `SDA` (weekly) | `PORSDAForecast` | ~79 weekly WW | `CommonName` | Default for "SDA" |
| `SDA Monthly` | `PORMonthlyForecast` | ~19 monthly anchors | `CommonName` | **Relabel WW -> MMM YYYY** |
| `MOR` | `MORSpread` | 21 absolute YYYYQQ | `MetricName` | **Default MOR shorthand** |
| `MOR Forecasts` | `MORForecasts` | ATRev+1..+16 (capped) | `MetricName` | **Relabel ATRev+N -> YYYYQQ** |
| `ENG` | `ENGForecasts` | ES0/ES1/ES2/QS/PO | `ATGroup` | Engineering milestones |

**Decision rule**: ambiguous "MOR" -> `MORSpread`. Use `MORForecasts` only if user explicitly says "MOR Forecasts" or wants PRQ-relative view.

---

## 2. Shorthand Vocabulary (user-confirmed)

### Report shorthand
- `SDA <product>` -> PORSDAForecast
- `SDA Monthly <product>` -> PORMonthlyForecast
- `MOR <product>` -> MORSpread (NOT MORForecasts)
- `ENG <product>` -> ENGForecasts
- Bare `<product>` -> **ASK** via `vscode_askQuestions` multi-select

### Filter shorthand
- **No filter given -> ASK** via chooser (Class/PPV/BI/Yield/All). Don't assume "all" silently.
- `+ <op>`: client-side startswith match. `TEST` matches `TEST_MPS`, `TEST_AHMT`, etc.
- `+ Class`: union of (op startswith `TEST`) OR (blank op AND met IN `MPS MONITOR`, `EQA MONITOR`, `CS MONITOR`). Always include BOTH parts.
- `+ PPV`: union of (op startswith `PPV`) OR (blank op AND met = `PPV-M SAMPLE SIZE` exact). Always include BOTH parts.
- `+ PPVs`/`PPVm`: single-op only (`PPV_SPS`/`PPV_SPM`).
- `+ BI`/`Burn-In`: op startswith `BURNIN`. No monitors (unlike Class).
- `+ yield`: blank op AND met IN (`U/D`, `R/D`, `FINISH YIELD`). SDA only.

### PPV common rules
- Units are MINUTES. `ETT` -> MetricName = `TEST TIME - MIN` (exact). `RCS` -> `RETEST RATE`. `SS` -> `PPV-M SAMPLE SIZE` (blank-op rows, value in kU).
- Site filter: `PGAT=PG8`, `VNAT=SS`, `CDAT=CD6`.
- DLCP filter: if requested DLCP not found, use blank row (aggregate) without asking.
- `from <month> to EOL`: set all time columns from start through last.

### Burn-In WIF shorthands
- `K LOT%` -> `Killable Lot Pct` (%). `Stress Time` -> `Stress_Time` (mins). `REBI` -> `REBI` (%).

### WIF (Work In Flight)

WIF = clone matched POR row(s), set `Proposed/POR` to `WIF`, set time-range columns to new value.

**WIF is orthogonal to operation filter.** Always ask as a SEPARATE 4th question after variant/report/filter.
- **No** -> view current POR values (build+email as normal)
- **Yes** -> collect via structured pop-up selections:
  1. **Spec** (from lookup scoped to chosen filter):
     - Class: `Classhot ETT`, `Classhot RCS`, `ShortEQA ETT/RCS`, `EQAm ETT/RCS`, `CVVm ETT/RCS`, `MPS Monitor`
     - BI: `K LOT%`, `Stress Time`, `REBI`
     - PPV: `PPVs ETT`, `PPVm ETT`, `SS`
  2. **Start WW & End WW** (free text, accept `EOL`)
  3. **Value** (free text with units)
  4. Optional **Site**/**DLCP** selects

**WIF Spec Lookup Table**:

| Shorthand | Operation | SubObject | MetricName | BOM filter |
|---|---|---|---|---|
| `Classhot ETT` | TEST_MPS | PBIC1 | TEST_TIME-SEC | blank only |
| `Classhot RCS` | TEST_MPS | PBIC1 | RETEST RATE | blank only |
| `ShortEQA ETT` | TEST_QA | EQA | TEST_TIME-SEC | blank only |
| `ShortEQA RCS` | TEST_QA | EQA | RETEST RATE | blank only |
| `EQAm ETT` | TEST_MPS_MONITOR | EQAM | TEST_TIME-SEC | blank only |
| `EQAm RCS` | TEST_MPS_MONITOR | EQAM | RETEST RATE | blank only |
| `CVVm ETT` | TEST_MPS_MONITOR | CVVM | TEST_TIME-SEC | blank only |
| `CVVm RCS` | TEST_MPS_MONITOR | CVVM | RETEST RATE | blank only |
| `MPS Monitor <site>` | blank | blank | MPS MONITOR | blank only |

**WIF xlsx layout (2-tab)**:
- Tab 1 (`<filter> Baseline`): full filtered baseline (same as if WIF=No)
- Tab 2 (`WIF`): ONLY matched POR rows + their WIF clones

**Color coding**: yellow bg on all WIF rows (`Interior.Color = 0x66FFFF`). Red font (`Font.Color = 0x0000FF` BGR) ONLY on cells changed for THAT spec. Use `$rowRedCols` hashtable keyed by Excel row.

**Time-range syntax**: `Jul'26` (monthly) or `ww40'26` (weekly -> `202640`). Range: `ww40'26-ww51'26`. EOL: `Jul'26-EOL`. SDA Monthly headers are relabeled -- resolve AFTER relabeling.

**Multiple specs**: `WIF: Classhot ETT ww40-43 1000, Classhot RCS ww42-ww01'27 7%`. Each gets own range/value/color.

**When WIF=Yes, send ONE email only** -- the 2-tab xlsx. Do NOT send a separate baseline first.

**Reference script**: `tci_wif_template.ps1`. Set `$ProductCache`, `$ShortName`, `$WifSpecs`, then invoke.

### Product disambiguation
- **Short family name** (ARL, PTL, RPL, NVL, TTL, GNR, TWL, CML): ALWAYS fetch CommonName panel and show ALL matching segments via `vscode_askQuestions`. Never assume only 1-2 variants.
- **ALWAYS re-prompt variant chooser** even if same product was picked last time.
- E-Temp variants: bare name -> non-E-Temp. Only pick E-Temp if user says so.
- ADL: abbreviated format (`ADL P 2+8+2`). "Catlow" = platform suffix, may match multiple.

---

## 3. Defaults

- **PHI Parameters**: select ALL checkboxes (~65-66)
- **AT_Group / AT_SubGroup** -- resolve in order:
  1. **Lookup file** (definitive): `tci_commonname_lookup.json`
  2. **Suffix heuristic** (fallback): S/H/HX -> Client/Desktop (BGA -> Mobile); U/P/M/Y/N -> Client/Mobile; SP/AP/XCC -> Server
  3. **Lookup refresh**: `tci_build_commonname_lookup_chunked.ps1` (6 chunked POSTs)
- **Output**: `<safe_commonname>_<report>[_<filter>].xlsx` in `$env:TEMP`
- **Email**: auto-resolved via `$ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress`

---

## 4. Auth & Transport

- Windows Integrated Auth. `Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing` with persistent session.
- ASP.NET WebForms: POST back `__VIEWSTATE`, `__VIEWSTATEGENERATOR`, `__EVENTVALIDATION`.
- Checkboxes: `ctl00$ContentPlaceHolder1$ctlNN` = `on`. Button: `btn_RunReport=Run Report`.
- **Server bug**: `btn_Export`/`btn_MailReport` throw SQL error. Scrape from "Run Report" HTML instead.
- **URL fallback**: try `http://tcitools.intel.com/TestReport.aspx?R=<id>` first; fall back to `https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=<id>`. `Get-TciBase` in lib caches result 12h.
- **Never use `Get-TciBase` directly** as URL -- always use `Tci-Get -ReportId <id>` / `Tci-Post -ReportId <id>` which append `?R=` automatically.

---

## 5. Standard Recipe (end-to-end)

```mermaid
flowchart LR
  A[User request] --> B[GET init page]
  B --> C[Parse hidden fields + panels]
  C --> D[Resolve CommonName -> ctlNN]
  D --> E[POST with checkboxes]
  E --> F[Cache HTML to %TEMP%]
  F --> G[Parse data table by first th]
  G --> H[Client-side filter]
  H --> I[Relabel time cols]
  I --> J[CSV -> Excel COM -> xlsx]
  J --> K[Outlook COM email]
```

**Key steps** (all code is in `tci_lib.ps1`):

1. **Init GET**: `Tci-GetInit -ReportId $rid` (cached 30min; ~7.6s first call, ~20ms cached)
2. **Parse panels**: `Get-CB` / `Get-AllCB` resolve label -> ctlNN. Filter panel IDs: `ContentPlaceHolder1_Filters_<AT_Group|AT_SubGroup|CommonName|PHIParameters|...>`
3. **POST**: `Tci-Post -ReportId $rid -Form $body -TimeoutSec 600`. Always re-enumerate ctlNN per report (MOR shifts indexes vs SDA).
4. **Parse**: `Tci-ParseTables` + `Tci-PickDataTable` (depth-tracked, match by first `<th>`)
5. **Filter**: `Tci-ApplyFilter` or manual predicate loop
6. **Relabel**: SDA Monthly -> `WwToFiscalMonth` (Intel 4-4-5 anchors). MOR Forecasts -> `AddQ` from ATREVStartDate.
7. **Build xlsx**: `Tci-ExportXlsx` (bulk 2D array write, bold/grey header, freeze row 1, autofit)
8. **Email**: `Tci-SendMail` with `Build-PhiCard` / `Build-WifCard` HTML

### SDA Monthly relabeling (Intel 4-4-5 fiscal anchors)

| Month | Jan | Feb | Mar | Apr | May | Jun | Jul | Aug | Sep | Oct | Nov | Dec |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Anchor WW | 01 | 05 | 09 | 14 | 18 | 22 | 27 | 31 | 35 | 40 | 44 | 48 |

Rule: given YYYYWW, pick the largest anchor <= WW; year stays YYYY. Do NOT use Sunday-of-WW calendar math.

**SDA Weekly: KEEP raw WW codes (no relabel).**

### MOR Forecasts relabeling

If all rows share one `ATREVStartDate`, rename: `ATRevStart` -> base YYYYQQ, `ATRev+N` -> `AddQ(base, N)`. Drop ATRev prefix entirely. If mixed dates, leave as-is.

---

## 6. Multi-Tab Workbook

When user asks for multiple reports, build one xlsx with one tab per report. Pattern: blank wb with `_placeholder_` sheet -> Move each section sheet in -> delete placeholder -> SaveAs. Use `Worksheet.Move([System.Reflection.Missing]::Value, $afterSheet)` (not Copy -- PS chokes on $null).

---

## 7. Cache Reuse

If `tci_after_<product>_<report>.html` is fresh (12h default), skip POST. Check: `(Test-Path $f) -and (Get-Item $f).Length -gt 0 -and (Get-Item $f).LastWriteTime -gt (Get-Date).AddHours(-12)`.

Same cache serves multiple filter variations on the same product+report.

---

## 8. Empty-Result Protocol

When filter returns 0 rows: (1) do NOT email empty file; (2) report finding + available metrics; (3) offer alternatives: different filter, re-POST fresh, or abort.

---

## 9. Gotchas

### PowerShell 5.1
- ASCII-only in .ps1 source. Use entities in HTMLBody (`&mdash;`, `&times;`).
- `$pid` is read-only. Multi-line paste breaks -- use .ps1 files.
- `${var}:` escaping for colons. Comments in `;`-chains break terminal.
- Function return pollution: bare strings + return hashtable = array to caller. Use `Write-Host` for diagnostics.
- `return $out` unrolls arrays. Use `return ,$out` to prevent.
- `$uniq = $bases | Sort-Object -Unique` returns scalar when 1 element. Force: `@(... | Sort-Object -Unique)`.

### TCI / HTML parsing
- Depth-track `<table>` -- naive regex grabs descendants.
- MOR sidebar contains "CommonName"/"ATRev" substrings -- match by first-`<th>`-equals only.
- MOR ctlNN != SDA ctlNN. Always re-enumerate from init.
- `Tci-Post` returns ARRAY -- `.Length` is wrong. Parse from saved file via `Get-Content -Raw`.

### Excel COM
- Use `Move` not `Copy`. After Move source wb auto-disposes -- try/catch on Close.
- FreezePanes: `$ws.Activate()` first, then set.
- PS 5.1: pre-compute `$exRow = $ri + 1` before indexing `[,]` array.
- `UsedRange.Rows.Count` off-by-one from trailing newline. Use known row count.
- `$rowRedCols` key = Excel row number (header=1, data=2..N).

### Long POSTs
- Some products take 5-15 min. Use `-TimeoutSec 600`. TCI caps at ~1000 rows per response (chunk heavy requests).
- POST can hang server-side. `Tci-Get`/`Tci-Post` have bounded-timeout + auto-retry.

---

## 10. Shared Library: tci_lib.ps1

Source at top of EVERY script: `. "$PSScriptRoot\tci_lib.ps1"`

Key exports: `Get-TciBase`, `Tci-Get`, `Tci-Post`, `Tci-GetInit`, `Tci-ParseTables`, `Tci-PickDataTable`, `Tci-RowsFromTable`, `Tci-WriteCsv`, `Tci-ExportXlsx` (with `-Validate -MinRows -MinCols`), `Tci-AssertXlsx` (legacy), `Tci-SendMail`, `Build-PhiCard`, `Build-WifCard`, `WwToFiscalMonth`, `Test-WwToFiscalMonth`, `AddQ`, `Tci-ApplyFilter`, `Test-PpvRow`, `Test-ClassRow`, `Test-ClassPpvRow`, `Add-PhiRunHistory`.

Bug fixes land in `tci_lib.ps1` FIRST -- do not copy-paste into scripts.

---

## 10b. Process Lessons (condensed)

1. Verify time-axis labels against real data BEFORE emailing.
2. Update SKILL.md FIRST when a bug is caught.
3. Validate before send: prefer `Tci-ExportXlsx -Validate`.
4. Cache reuse by default. Only re-fetch with explicit flag.
5. 53-week fiscal year: run `Test-WwToFiscalMonth` at start of new year.
6. Class+PPV = union (rows matching either). Sheet suffix: `ClassPPV`.
7. MOR OperationName is NOT col 0. Always `[Array]::IndexOf`.
8. Multi-tab: activate sheet before FreezePanes.
9. Ambiguous codename -> ask. Show all variants.
10. Legacy products have no ENG data (expected).
11. Not all products have all ops (Nova Lake H: 0 TEST rows).
12. All card text is 11pt Calibri. Single source of truth in lib.
13. Parallelize ONLY network, never COM. Sliding window max 3 concurrent POSTs.
14. TCI concurrency limit: >3 simultaneous POSTs causes timeouts (SDA worst at 7.7% success under load).
15. WIF pitfalls: probe cache first; never iterate while appending; snapshot `$origCount`.
16. Long jobs: run ASYNC, tail the Tee-Object log file.
17. Reuse cached HTML across filter variations (12h window).
18. WIF applies to ALL report types including ENG milestones.
19. ENG has NO `Proposed/POR` column -- use position-based WIF identification.

---

## 11. Verified Examples

See [tci_verified_log.md](tci_verified_log.md) for full regression reference counts and WIF examples.

---

## 12. Probe One-Liners

Use `tci_lib.ps1` functions for diagnostics:
- **List CommonNames**: parse cached init HTML, regex for checkbox labels matching keyword.
- **Inspect cached POST**: `Tci-ParseTables` + `Tci-PickDataTable` on `tci_after_*.html`, then group by OperationName.

See `test_*.ps1` scripts in workspace for ready-made probes.

---

## 13. Future Extensions

- `Class=offline` -> union of OFFLINE_BINNING* + 3 blank-op monitors
- Periodic refresh of `tci_commonname_lookup.json` (monthly via chunked builder)
- Other report types if TCI adds them (probe via `R=` regex on landing page)

---

## 14. Batch Driver

**Purpose**: drive multiple TCI runs from code or Excel input.

### Direct invocation (preferred)
```powershell
.\tci_run.ps1 -Product 'ARL Refresh S 8C+16A+GT1' -Report MOR -Filter PPV
.\tci_run.ps1 -Product 'CML PCH','Nova Lake U' -Report 'SDA Weekly' -NoEmail
.\tci_run.ps1 -Product $familyList -NoEmail -CombineXlsx -MaxCacheHours 0
```

### Parameters
| Param | Default | Notes |
|---|---|---|
| `-Product` | (required) | `[string[]]`, splits comma/semicolons |
| `-Report` | all 4 | SDA/SDA Monthly/MOR/ENG |
| `-Filter` | none | class/ppv/bi/yield/+ prefix |
| `-MaxCacheHours` | 12 | 0 = always fresh |
| `-NoParallel` | false | Force sequential POSTs |
| `-CombineXlsx` | false | All products in one workbook |
| `-NoEmail` | false | Skip Outlook COM |

### Parallel prefetch (sliding window)
- `Start-BulkPrefetch` warms POST caches via runspace pool (max 3 concurrent).
- Sorted by weight: ENG=1, MOR=2, SDAMO=3, SDA=4. SDA timeout 600s (others 480s).
- Init pages pre-fetched once (4 unique GETs). Failure is non-fatal -- sequential pass re-fetches.
- ~4.7 min/product. 15-product family: ~1h44m, 22,872 rows.

### Excel-driven input
Optional `TCI_Batch_Input.xlsx` (sheet `Batch`): columns `Product` (required), `Report`, `Filter`, `Email`.

### Scheduled weekly
```powershell
.\tci_weekly.ps1           # run tracking list now
.\tci_weekly.ps1 -Register # install Monday 07:00 scheduled task
```

### Known product quirks
| Product | Notes |
|---|---|
| Wildcat Lake (4 variants) | Client/Mobile (NOT NPI -- lookup was wrong) |
| ARL S (N3B) | Client/Desktop, `(N3B)` = process stepping |
| Titan Lake (B/BX/HL/HM/HPX) | Client/Mobile, MOR only (no SDA) |
| ALL Nova Lake (22 segments) | Client/Mobile only |
| Nova Lake H | NO TEST* ops. Only PPV/BURNIN/blank-op |
| BGA variants of S/H/HX | Client/Mobile (BGA = soldered laptop) |
| Extended Temp variants | Not in ENG panel -- ENG tab skipped gracefully |

---

## 15. Implementation Hard Rules

These exist because each was learned from a production failure.

### Frozen formats (rule 0 -- NEVER change)

Three things NEVER change, day to day, run to run, cached product or brand-new product. Do NOT reword, restyle, reorder, or "improve" them. EVER. If a request seems to need a different format, ASK first -- never change unilaterally.

- **(A) Question / picker input format.** The `vscode_askQuestions` batch is ALWAYS the same 4 questions with these EXACT `header`s and EXACT option `label`s (frozen tokens -- no parentheticals, no notes, no re-ordering):
  1. `header:"Product"` (`multiSelect:true`) -- one option per full TCI CommonName label, verbatim from the panel (e.g. `Panther Lake U 4P+0E+4LP_E`). Nothing appended.
  2. `header:"Report"` (`multiSelect:true`) -- options EXACTLY: `SDA Weekly` / `SDA Monthly` / `MOR` / `ENG` / `All`.
  3. `header:"Filter"` (single-select) -- options EXACTLY: `Class` / `PPV` / `BI` / `Yield` / `All`.
  4. `header:"WIF"` (single-select) -- options EXACTLY: `No` / `Yes`.
  - Only the Product option list changes (per the live panel); the structure, headers, and the other three option sets are immutable. If WIF=Yes, the follow-up is ALWAYS: `header:"Spec"` (multiSelect, S2 specs scoped to the chosen operation) + `header:"Range"` (free text `<startWW> - <endWW>`, accepts `EOL`) + `header:"Value"` (free text number).
- **(B) Excel output format (Excel COM).** Header row: bold, fill `Interior.Color = 14277081` (grey 0xD9D9D9), frozen top row (`SplitRow=1; FreezePanes`) plus `SplitColumn=4` freeze for the left key columns; `Columns.AutoFit()`. Filter tab named by suffix (`Class`/`PPV`/`Burnin`/`Yield`). WIF = **2-tab** workbook: tab 1 `<Filter> Baseline`, tab 2 `WIF` (POR row + WIF clone; WIF rows yellow fill `0x66FFFF`, changed cells red font `0x0000FF`, `Proposed/POR` col = `WIF`). SaveAs format `51` (.xlsx). Filename `<safe_commonname>_<report>_<Filter>[_WIF].xlsx`.
- **(C) Email format.** ALWAYS `Build-PhiCard` (no WIF) / `Build-WifCard` (WIF) from `tci_lib.ps1` -- never raw/inline HTML. Amber/orange theme, Calibri 11pt, sent via Outlook COM to the user's `PrimarySmtpAddress`. Subject EXACTLY: `PHI of <product> - <report> <filter>` (WIF appends ` + WIF <spec>`). xlsx attached.

### Data extraction
1. **Concatenate ALL data tables.** TCI splits results across 3-5 tables (~200 rows each). First table only = 60-90% data loss.
2. **Full-page checkbox parsing.** Panel-scoped parsing truncates labels. Regex entire init HTML.

### Excel build
3. **Bulk 2D array write.** Cell-by-cell is ~100x slower (277 rows x 100 cols = 15+ min vs <2s bulk).
4. **Kill stale Excel before build.** `Get-Process Excel -EA SilentlyContinue | Stop-Process -Force`
5. **PS 5.1 indexer.** Pre-compute `$exRow = $ri + 1` then `$arr[$exRow, $ci]`.

### Email
6. **Source `Build-PhiCard`/`Build-WifCard` from tci_lib.ps1.** Never inline HTML -- it drifts between runs.
7. **WIF=Yes -> ONE email** (2-tab xlsx). Tab 1 already has baseline.

### WIF
8. **Tab 1 = full baseline. Tab 2 = matched POR + WIF clones only.**
9. **Color per-spec.** Yellow row, red font ONLY on changed cells for THAT spec.
10. **Never iterate while appending.** Snapshot `$origCount` before clone loop.

### Relabeling
11. **SDA Monthly: relabel BEFORE resolving WIF column ranges.** Weekly keeps raw WW codes.

### Interactive flow
12. **ALWAYS ask before running.** 4-step chooser via `vscode_askQuestions`: Variant -> Report -> Filter -> WIF. Batch in ONE call. Skip steps user already provided.
13. **Never run on bare product name.** If no report+filter+WIF stated, pop the chooser.
14. **Consistent output.** Same format regardless of who runs it -- xlsx/card/WIF locked.

### General
15. **POST timeout = 600s.** Mature products take 4-6 min.
16. **Cache reuse (12h).** Check freshness before POST.
17. **Source `tci_lib.ps1` at top of EVERY script.**
