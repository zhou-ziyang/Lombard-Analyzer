# Lombard Analyzer

VBA tooling for analysing a Lombard loan portfolio out of daily Sophis CSV
extracts. The code lives in an Excel workbook; this repository holds the
exported standard modules so they can be diffed and versioned.

Everything is driven from the workbook's **Home** sheet, which carries the
configuration parameters (as defined names) and one button per entry point.

## Entry points

| Home section | Defined names | Button | Entry point |
| --- | --- | --- | --- |
| 01 Configuration | `path`, `report_path`, `Certificate_Path` | Text-to-Column Settings | `TextToCol.text_to_col1` |
| 01 Configuration | — | Clear Sheets | `Clean.Clean` |
| 02 Date Range Analysis | `AnalysisStartDate`, `AnalysisEndDate` | Calculate Delta | `DeltaCalculation.BuildPositionMovements` |
| 02 Date Range Analysis | `AnalysisEndDate` | Revenue Estimate | `RevenueTools.BuildRevenueSummary` |
| 03 Weekly Analysis | `WeeklyEndDate` | Weekly Analysis | `WeeklyAnalysisGenerate.GenerateWeeklyAnalysis` |
| 04 Client Dashboard | `JourneyNDG`, `journey_start` | Launch Dashboard | `Journey.ExtractNDGHistory` |

`WeeklyAnalysisEmail.CreateWeeklyEmail` is reached from a button that
`GenerateWeeklyAnalysis` draws onto the generated *Weekly Analysis* sheet, and
`JourneyPositionAnalysis.AnalyzePositionChanges` from the per-row *Analyze*
buttons that `AddPositionAnalysisButtons` draws onto *NDG Journey* and the
dashboard's history table.

## Source data

`path` points at a folder of daily extracts named by snapshot date:

```
<yyyymmdd>_Lombard_Loans_ITA_Positions.csv
<yyyymmdd>_Lombard_Loans_ITA_Accounts.csv
```

Both are semicolon-delimited. `report_path` points at an external workbook
whose *Report* sheet supplies margin-call / shortfall reasons and comments;
`Certificate_Path` points at the folder of certificate basket extracts used to
decompose structured products into their underlyings.

## Reference sheets

These are maintained by hand (with help from Bloomberg formulas) and are read,
not generated:

| Sheet | Table(s) | Purpose |
| --- | --- | --- |
| Companies | `Companies` | Master entity table: canonical name, name variants, exposure types, reference ISIN and its relationship, country of risk and sector with fallbacks |
| Bond Issuers | `BondIssuers` | Issuer ticker → issuer name and Corporate/Sovereign type |
| Fund Parent Companies | `FundParentCompanies`, `Funds` | Fund name prefix → parent company, plus per-fund overrides |
| Equity Names | `UnmappedEquities` | Queue of equity ISINs that resolved to no company; filled in by hand |
| Countries | — | Country code → country name |
| Name Variants | — | Manual entity-name variant overrides |

Generated sheets (*Weekly Analysis*, *Asset Type Mapping*, *New Geo-Sec
Lookup*, *Risk Exposure*, *NDG Journey*, *NDG Dashboard*, *Position Change
Analysis*, *Revenue Summary*, `Delta_<yyyymmdd>`, `Closed_<yyyymmdd>`) are
rebuilt from source and are not committed here.

## Layout

```
src/core/       Utils, DataTools, GlobalVariables, Cache, ImportTools, Clean, TextToCol
src/weekly/     WeeklyAnalysisGenerate, WeeklyAnalysisLayout, WeeklyAnalysisEmail
src/journey/    Journey, JourneyFormatting, JourneyDashboardTable, JourneyPositionAnalysis
src/delta/      DeltaCalculation, RevenueTools, PortfolioDataTools
src/reference/  AssetMapping
archive/        JourneyVisualization (superseded)
```

`archive/JourneyVisualization.bas` is commented out in full. Its charting
procedures were revived inside `JourneyDashboardTable`, which now carries the
same constants and procedure names alongside the Customer Overview and
Historical Events tables. It is kept only for reference.

### Pipelines

**Weekly** — `GenerateWeeklyAnalysis` loads the current, comparison and
year-end snapshots, then builds the report sections in place on one sheet
using the coordinates in `WeeklyAnalysisLayout.Layout`: portfolio overview,
collateral breakdown, new/ended loans, entered collateral, the pie chart, and
the exposure concentration block. The concentration block is the bulk of the
module: certificate baskets are expanded recursively into their underlyings,
entity names are normalised and merged (diacritics, legal suffixes, share
class suffixes, prefix matching, manual variants), resolved against the
reference sheets, staged into the `RiskExposure` table with an account scope
flag, then aggregated twice — full portfolio and excluding aggregated accounts
— into top-10 tables by name, geography and sector.
`CreateWeeklyEmail` re-exports the finished ranges as HTML and assembles the
Outlook message.

**Journey** — `ExtractNDGHistory` walks every Accounts snapshot for one NDG,
synthesises `Loan Ended` / `Loan Restarted` rows when the account disappears
and returns, derives the delta/LTV/event columns, backfills reasons and
comments from the external report workbook, formats the margin-call and
shortfall bands, then builds the dashboard. `AnalyzePositionChanges` compares
two position snapshots for one NDG and attributes the change in market value
and haircut collateral value to price, quantity, composition and residual
effects.

**Delta** — `BuildPositionMovements` diffs the start and end position
snapshots into `Delta_<date>` (new positions, new NDGs) and `Closed_<date>`,
using first-seen/last-seen dictionaries built by scanning every snapshot in
the range. `BuildRevenueSummary` prices the result by asset class.

## Encoding

The `.bas` files are exported by the VBE as Windows-1252 with CRLF line
endings, and `.gitattributes` marks them `-text` so Git does not normalise
either. Editing them with a UTF-8 tool will corrupt single-byte characters —
the euro signs in `RevenueTools` are the ones that bite first.

## Known issues

Recorded here rather than fixed, so the exported modules still match the
workbook they came from.

**Leaves Excel in a bad state**

- `GenerateWeeklyAnalysis` calls a bare `Reset` in both `ExitRoutine` and
  `ErrorHandler` (`src/weekly/WeeklyAnalysisGenerate.bas:1110,1122`). No
  procedure by that name exists — `Utils` defines `ResetExcel` — so this is
  VBA's built-in `Reset` statement, which closes open files and does nothing
  else. `Application.Calculation` is left on `xlCalculationManual` and
  `EnableEvents` on `False` after every run. `DeltaCalculation` calls
  `ResetExcel` correctly.
- `NoteHandler` is restored only in `ErrorHandler`, not on the success path,
  so after one successful weekly run the global stays pointed at
  `WriteNoteWeekly` and later `Note` calls from other modules are swallowed.

**Wrong numbers**

- `BuildRevenueSummary` excludes `"Non Eligible Asset"` but not `"UNKNOWN"`,
  which is what `GetAssetClass` returns for an unmapped asset type. Those rows
  reach the summary with a zero revenue rate but still add their position
  value to the Eligible Assets column, and the Asset Management Uplift row
  sums that whole column — so unmapped collateral inflates the uplift base.
  The `Calculate Delta` path never calls `RegisterUnknownAsset`, so nothing
  warns about it.
- `DictAnnualAMRevenue` is only `.Add`ed in the first-seen branch of
  `BuildRevenueSummary` and never accumulated afterwards, so its values are
  wrong for any asset class appearing more than once. Currently latent: the
  dictionary is not read for output.
- `GetComparisonDate` computes `DateSerial(y, Month - MonthsBack, Day)`, which
  rolls 31 March back to 3 March rather than to the end of February.

**Fragile**

- `ReadAllLines` and `ImportCsvByDate` both hard-code file number `#1` instead
  of `FreeFile`. Safe while every call is serial; error 55 the moment one
  nests inside another.
- `ReadAllLines` splits on `vbCrLf` only. An LF-terminated extract would parse
  as a single line and yield zero rows without raising anything.
- `BuildPositionMovements` validates only `idxAsset` and `idxValue`;
  `BuildPositionDateDictionaries` validates no header index at all. A missing
  column makes `FindHeaderIndex` return -1 and the position key silently loses
  a segment, rather than failing the way `RequiredPositionHeader`,
  `RequiredAccountHeader` and `RequiredColumn` do elsewhere.
- `LoadPositionCache` sets `LineCount = UBound(Lines)` and then
  `ReDim Cache.Data(1 To LineCount)`; a header-only file gives `1 To 0` and
  raises error 9. `BuildPositionMovements` sizes its output arrays the same
  way.
- `GetAssetClass` matches with `Like` in a module without `Option Compare
  Text`, so the patterns are case-sensitive. A casing change in the source
  extract sends whole asset classes to `"UNKNOWN"`.
- `Clean`'s keep-list names twelve sheets that no longer exist in the workbook
  (`Code`, `DateRange`, `PEC List`, `Database`, `Report`, `CLN`, …) and does
  not name `Risk Exposure` or `New Geo-Sec Lookup`. Both of those are caches
  the weekly run is written to reuse — `RiskStageTableCanBeReused` and
  `StoredRiskStageDate` exist precisely to avoid rebuilding them — and both
  get deleted.
- `BuildRevenueSummary` resolves `Delta_<AnalysisEndDate>` with no error
  handling, so pressing *Revenue Estimate* after changing the end date but
  before re-running *Calculate Delta* surfaces a raw subscript-out-of-range
  error.

**Tidying**

- `ImportCsvByDate` recomputes `LastRow = ws.Range("A2").End(xlUp).Row` inside
  its write loop; column A is always empty so the value is always 1. The
  result happens to be correct but the lookup is dead weight. Its
  `TextToColumns` call also splits on `Chr(10)` where the rest of the codebase
  splits these files on `";"`.
- `FormatJourneyTable` calls `FormatReportTable` twice in a row on the same
  range.
- `WriteHeader` and `WriteRow` in `DeltaCalculation` have no callers, and its
  `ErrHandler` / `CleanExit` labels are unreachable because
  `On Error GoTo ErrHandler` is commented out.
- `PortfolioDataTools` has no callers anywhere in the codebase;
  `WeeklyAnalysisGenerate` computes the same figures with
  `CalculateWeeklyPortfolioStats`. It may still be bound to a button.
- `RegisterUnknownAsset` takes an optional `NoteHandler` argument that shadows
  the global of the same name and is never used.
- `Journey.bas`'s header comment refers to `Code!journey_start` and
  `Code!report_path`; both parameters live on *Home*, which is what the code
  reads.
- `WeeklyAnalysisEmail` hard-codes the recipient addresses in `.To` and `.CC`.
