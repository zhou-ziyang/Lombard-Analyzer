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
| 03 Weekly Analysis | `WeeklyEndDate`, `EmailTo`, `EmailCc` | Weekly Analysis | `WeeklyAnalysisGenerate.GenerateWeeklyAnalysis` |
| 04 Client Dashboard | `JourneyNDG`, `journey_start` | Launch Dashboard | `Journey.ExtractNDGHistory` |

`WeeklyAnalysisEmail.CreateWeeklyEmail` is reached from a button that
`GenerateWeeklyAnalysis` draws onto the generated *Weekly Analysis* sheet, and
`JourneyPositionAnalysis.AnalyzePositionChanges` from the per-row *Analyze*
buttons that `AddPositionAnalysisButtons` draws onto *NDG Journey* and the
dashboard's history table.

`EmailTo` and `EmailCc` hold the draft's recipients, semicolon-separated, and
are read like any other Home parameter. A name that has not been created yet
reads as empty: the draft still opens, and a message says which name to add.
The addresses deliberately live in the workbook rather than in source.

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

`Clean` keeps only the sheets on its own list and deletes everything else,
*Risk Exposure* and *New Geo-Sec Lookup* included — those are caches, and
rebuilding them is the intended behaviour. The list also names sheets that
are not in this workbook (`Code`, `DateRange`, `PEC List`, `Database`,
`Report`, `CLN`, …) on purpose, so the same module can be dropped into the
other Lombard workbooks without editing it.

## Layout

```
src/core/       Utils, DataTools, GlobalVariables, Cache, ImportTools, Clean, TextToCol
src/weekly/     WeeklyAnalysisGenerate, WeeklyAnalysisLayout, WeeklyAnalysisEmail
src/journey/    Journey, JourneyFormatting, JourneyDashboardTable, JourneyPositionAnalysis
src/delta/      DeltaCalculation, RevenueTools
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
Outlook message. `docs/weekly-analysis-generate.md` walks through that module
in detail — the staging table's schema, the certificate recursion, the entity
name normalisation, and the three separate asset classifications.

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

### Asset classification and UNKNOWN

`GetAssetClass` maps the raw Sophis Asset Type string onto the reporting asset
classes. `"UNKNOWN"` means the string matched none of its patterns — it does
not mean the position is ineligible. A live position is eligible unless it is
explicitly marked Non Eligible, so UNKNOWN collateral is counted as eligible
and the unmatched string is surfaced for the mapping rules to be extended.

How it is surfaced differs by design. Revenue Summary has a row per asset
class, so an unmatched string appears there as a row literally labelled
UNKNOWN — visible, and the prompt to go fix the rules. The weekly Collateral
Breakdown has no such column, so an unmatched string would vanish silently;
that path calls `RegisterUnknownAsset` instead, which writes the string into
the report's Notes box.

## Encoding

The `.bas` files are exported by the VBE as Windows-1252 with CRLF line
endings, and `.gitattributes` marks them `-text` so Git does not normalise
either. Editing them with a UTF-8 tool will corrupt single-byte characters —
the euro signs in `RevenueTools` are the ones that bite first.

## Known issues

Open items. Anything fixed has been removed from this list; see the commit
history for what changed.

**Kept on purpose**

`ImportAccountsByDate` and `ImportPositionsByDate` have had no caller since
`PortfolioDataTools` was removed, and `ImportCsvByDate`, `SourceFileExists` and
`ResolveAvailableDate` sit behind them. They are retained for future use — a
dead-code sweep should skip them rather than take half of `ImportTools` with
them. Because nothing exercises that path, the write loop's simplification
keeps the statements it replaced as a comment.

**Tidying**

- `ImportCsvByDate`'s `TextToColumns` call splits on `Chr(10)` where the rest
  of the codebase splits these files on `";"`. Left alone until the intent is
  known.
- `Journey.bas`'s header comment refers to `Code!journey_start` and
  `Code!report_path`; both parameters live on *Home*, which is what the code
  reads.
