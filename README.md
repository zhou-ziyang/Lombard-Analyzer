# Lombard Analyzer

VBA tooling for analysing a Lombard loan portfolio out of daily Sophis CSV
extracts. The code lives in an Excel workbook; this repository holds the
exported standard modules so they can be diffed and versioned.

Everything is driven from the workbook's **Home** sheet, which carries the
configuration parameters (as defined names) and one button per entry point.

## Entry points

| Home section | Defined names | Button | Entry point |
| --- | --- | --- | --- |
| 01 Configuration | `path`, `report_path`, `Certificate_Path` | Text-to-Column Settings | `CoreTextToCol.text_to_col1` |
| 01 Configuration | — | Clear Sheets | `CoreClean.Clean` |
| 02 Date Range Analysis | `AnalysisStartDate`, `AnalysisEndDate` | Calculate Delta | `DeltaCalculation.BuildPositionMovements` |
| 02 Date Range Analysis | `AnalysisEndDate` | Revenue Estimate | `DeltaRevenue.BuildRevenueSummary` |
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

`CoreClean` keeps only the sheets on its own list and deletes everything else,
*Risk Exposure* and *New Geo-Sec Lookup* included — those are caches, and
rebuilding them is the intended behaviour. The list also names sheets that
are not in this workbook (`Code`, `DateRange`, `PEC List`, `Database`,
`Report`, `CLN`, …) on purpose, so the same module can be dropped into the
other Lombard workbooks without editing it.

## Layout

```
src/core/       CoreUtils, CoreReportFormat, CoreImport, CoreCache,
                CoreClean, CoreTextToCol, CoreGlobals
src/reference/  RefAssetMapping
src/weekly/     WeeklyAnalysisGenerate, WeeklyAnalysisLayout, WeeklyAnalysisEmail
src/journey/    Journey, JourneyFormatting, JourneyDashboardTable, JourneyPositionAnalysis
src/delta/      DeltaCalculation, DeltaRevenue
tools/          ToolsInstall (loads a folder of modules), ToolsCountryProbe
                (formula-vs-VBA experiment) — neither is part of the workbook
archive/        JourneyVisualization (superseded)
```

### Module names carry the folder

Git has folders; the VBE does not — every module imported into the
workbook lands in one flat list. So each module name starts with the folder
it comes from, and the file name is that module name, which is also what the
VBE writes when it exports. `Journey*`, `Weekly*` and `DeltaCalculation`
already read that way and were left alone.

The layers are strict, and the call graph has no edge going the other way:

```
weekly, journey, delta   →   reference   →   core
```

`core` knows nothing about Lombard loans: sheets, values, headers, files,
table formatting. `reference` knows the asset taxonomy. Each pipeline knows
its own report and nothing about the other two — the shared helpers they used
to reach across for (`ReadAllLines`, `FindHeaderIndex`, `FormatReportTable`,
`SafeCellText`) now live in `core`.

`Public` is the whole namespace in VBA: any Public procedure in any standard
module is callable from every other, and two of the same name stop the project
compiling. So Public means "something outside this module calls this", and the
only Public procedures with no caller in the source are the eight zero-argument
entry points the Home buttons name. Two exceptions carry a comment saying why
they must stay Public: `WriteNoteWeekly`, which `Application.Run` reaches by
name, and `WriteAssetTypeMapping`, whose zero arguments make it bindable to a
button that would not be visible from the source.

### Why WeeklyAnalysisGenerate stays one module

It is 12,000 lines and 188 procedures, and it does not get split, because in
VBA splitting it would cost more than it buys. 175 of those procedures are
Private, along with six Enums and forty-odd Consts. The module is the only
encapsulation boundary the language has — there are no namespaces, and
`Private` means "private to this module", not "private to this concern". Cut
it into five, and every helper the pieces share has to become Public, which
means global: several hundred new names in the one namespace the whole project
shares, including enum members like `RiskStageNDG`.

So the boundary earns its size. What splitting would have bought — being able
to find things — the file order already gives, and the sections run in the
order the report is built: source loading and CSV parsing, the report
sections, risk reference data, certificate basket expansion, entity-name
normalisation, the staging table and its aggregation, then the chart and the
notes. `docs/weekly-analysis-generate.md` walks through them.

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

## Importing into the workbook

`tools/ToolsInstall.bas` does it in one pass. Import that module once
(Ctrl+M), tick **File → Options → Trust Center → Trust Center Settings →
Macro Settings → Trust access to the VBA project object model**, then run
`InstallModules` and point it at a folder of `.bas` files: it backs up every
standard module to a timestamped folder, removes them, and imports the folder.
It skips itself in both passes, and the confirmation names separately anything
about to be removed that the folder does not replace. Remove `ToolsInstall`
by hand when the modules are in — it loads the workbook, it is not part of it.

By hand instead: File → Import File… (Ctrl+M), one `.bas` per module, and
**remove the existing module first, every time**. Import does not overwrite —
the VBE keeps the old module and names the new one `DeltaCalculation1`, and
two modules holding the same procedures fail to compile with "Ambiguous name
detected". That applies to a module whose contents changed just as much as to
one that was renamed.

Pasting the text into a new module instead leaves `Attribute VB_Name` in the
body, where it is not valid VBA and shows as a syntax error — it is a
file-format directive the importer reads and strips.

## Experiment: the concentration arithmetic as formulas

`tools/ToolsCountryProbe.bas` builds the Country of Risk concentration, Full
scope, twice over the same `RiskExposure` staging table — once with worksheet
formulas and once with a VBA pass — and puts the difference between them in a
column. Import it, run `BuildCountryConcentrationProbe`, and read the three
numbers at the top of the *Country Probe* sheet: they should all be zero.

The staging table is already a fact table (one row per position × allocated
exposure, with every dimension beside the measure), so the 432 lines of
`AggregateUnifiedRiskStageData` are a hand-written `GROUP BY` that `SUMIFS`
and `UNIQUE` do natively. What the probe tests is not whether that is possible
but whether the *semantics* survive the translation — the three that are easy
to lose being the share denominator, the union-not-sum `#NDG` on the total
row, and the tie-break on equal values.

The whole ranked table — country, value and distinct NDG count, ordered and
cut to ten — is one `GROUPBY` formula per asset class, with the distinct count
written as `LAMBDA(x, COUNTA(UNIQUE(x)))` where the aggregate goes.

**It answers yes.** Against the VBA pass the values agree to the cent, the
per-country and union `#NDG` counts agree, and the category totals are equal.
One wrinkle worth keeping in mind for a real implementation: `GROUPBY` names
its value columns in a row of its own, that row is text, and sorting by value
descending therefore carries it to the top and pushes the tenth country out of
the table — `DROP(…, 1)` takes it off before the sort.

## Encoding

The `.bas` files are exported by the VBE as Windows-1252 with CRLF line
endings, and `.gitattributes` marks them `-text` so Git does not normalise
either. Editing them with a UTF-8 tool will corrupt single-byte characters —
the euro signs in `DeltaRevenue` are the ones that bite first.

## Known issues

Open items. Anything fixed has been removed from this list; see the commit
history for what changed.

**Kept on purpose**

`ImportAccountsByDate` and `ImportPositionsByDate` have had no caller since
`PortfolioDataTools` was removed, and `ImportCsvByDate`, `SourceFileExists` and
`ResolveAvailableDate` sit behind them. They are retained for future use — a
dead-code sweep should skip them rather than take half of `CoreImport` with
them. Because nothing exercises that path, the write loop's simplification
keeps the statements it replaced as a comment.

**Tidying**

- `ImportCsvByDate`'s `TextToColumns` call splits on `Chr(10)` where the rest
  of the codebase splits these files on `";"`. Left alone until the intent is
  known.
- `Journey.bas`'s header comment refers to `Code!journey_start` and
  `Code!report_path`; both parameters live on *Home*, which is what the code
  reads.
