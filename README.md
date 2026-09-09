# Lombard Analyzer

VBA tooling for analysing a Lombard loan portfolio out of daily Sophis CSV
extracts. The code lives in an Excel workbook; this repository holds the
exported standard modules so they can be diffed and versioned.

Everything is driven from the workbook's **Home** sheet, which carries the
configuration parameters (as defined names) and one button per entry point.

## Entry points

| Home section | Defined names | Button | Entry point |
| --- | --- | --- | --- |
| 01 Configuration | `path`, `report_path` | Text-to-Column Settings | `CoreTextToCol.text_to_col1` |
| 01 Configuration | — | Clear Sheets | `CoreClean.Clean` |
| 02 Date Range Analysis | `AnalysisStartDate`, `AnalysisEndDate` | Calculate Delta | `DeltaCalculation.BuildPositionMovements` |
| 02 Date Range Analysis | `AnalysisEndDate` | Revenue Estimate | `DeltaRevenue.BuildRevenueSummary` |
| 03 Weekly Analysis | `WeeklyEndDate`, `WeeklyCompareDate`, `EmailTo`, `EmailCc` | Weekly Analysis | `WeeklyAnalysisGenerate.GenerateWeeklyAnalysis` |
| 04 Client Dashboard | `JourneyNDG`, `journey_start` | Launch Dashboard | `Journey.ExtractNDGHistory` |

`WeeklyAnalysisEmail.CreateWeeklyEmail` is reached from a button that
`GenerateWeeklyAnalysis` draws onto the generated *Weekly Analysis* sheet,
`WeeklyAnalysisGenerate.InsertRenamedCompanies` from the *Insert Renamed*
button drawn onto *New Geo-Sec Lookup* when a company has changed its name, and
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
whose *Report* sheet supplies margin-call / shortfall reasons and comments.
The certificate reference used to decompose structured products into their
underlyings is two worksheets in this workbook, listed below.

## Reference sheets

These are maintained by hand (with help from Bloomberg formulas) and are read,
not generated:

| Sheet | Table(s) | Purpose |
| --- | --- | --- |
| Companies | `Companies` | Master entity table: canonical name, name variants, exposure types, reference ISIN and its relationship, country of risk and sector with fallbacks; *Renamed From* records the name a company had before — the row it was copied from, when Companies had one |
| Bond Issuers | `BondIssuers` | Issuer ticker → issuer name and Corporate/Sovereign type; *Previous Names* keeps every name Sophis has since corrected |
| Fund Parent Companies | `FundParentCompanies`, `Funds` | Fund name prefix → parent company, plus per-fund overrides |
| Equity Names | `UnmappedEquities` | Queue of equity ISINs that resolved to no company; filled in by hand |
| Countries | — | Country code → country name |
| Name Variants | — | Manual entity-name variant overrides |
| Certificates | — | Certificate ISIN → its underlying RIC(s) |
| Certificate Underlyings | — | RIC → underlying name, ISIN, asset class, basket component RICs |

Generated sheets (*Weekly Analysis*, *Asset Type
Mapping*, *New Geo-Sec Lookup*, *Risk Exposure yyyymmdd* — one per staged date, *NDG Journey*, *NDG Dashboard*, *Position Change
Analysis*, *Revenue Summary*, `Delta_<yyyymmdd>`, `Closed_<yyyymmdd>`) are
rebuilt from source and are not committed here.

Both certificate sheets are on `CoreClean`'s keep list, like every other
reference sheet: they hold maintained data, not a rebuilt cache.

*Companies* is fed from *New Geo-Sec Lookup*, which every staging rebuild
writes: one row per entity Companies does not know, or knows under fewer
names, exposure types or a different reference ISIN than the run saw. Columns
A–E come filled. *Country of Risk* (F) and *Sector* (H) come as `BDP`
formulas over the reference ISIN — `CNTRY_OF_RISK` and `INDUSTRY_SECTOR` —
so they resolve on a machine with a Bloomberg terminal; *Fallback Geography*
(G) and *Fallback Sector* (I) are left for a manual lookup. None of F–I is
ever copied from Companies. An entity missing from Companies still ranks by
name; its Country of Risk and Sector fall to *Others* until the row is added.

*Name Variants* on a Companies row are the spellings and legal forms that
count as one company; the *Name Variants* sheet is the fuzzy merge's own log,
with a *Manual Override* column for the cases the merge gets wrong.

### Renamed companies

A renamed company arrives as a name Companies has never heard of, and would
be looked up afresh and shown under *Others* until someone added it. Two
things vouch for it: for a bond issuer, the name *Bond Issuers* held before
Sophis corrected it — `UpdateRiskReferenceDatabases` now keeps that in
*Previous Names* instead of discarding it — and, for anything, the ISIN of
what it issued against Companies' *Reference ISIN* (issued and underlying
securities; a fund's ISIN names the fund, not its parent).

A renamed company is **a company of its own**. `DetectRenamedCompanies` runs
the two bridges over every entity Companies does not know; a match whose name
is not merely another spelling of the row's own name is registered for this
run as a copy of that row under the new name, so the report shows the new
name with the old row's geography, sector and reference ISIN rather than
*Others*. The old
row is never touched: positions that still carry the old name keep resolving
to it, and a report for an earlier date reads as it always did. The lookup
sheet lists the new name with *Renamed From* filled — in the same column
Companies has (or will get) for it, so rows paste across whole — with the
old row's geography and sector shown in place of the Bloomberg formula, and
draws an *Insert Renamed* button. Pressing it is the one way the code writes
to Companies, and it only adds: a new row copied whole from the old one —
reference ISIN included, since that is what identified it — with the new
name, the variants and exposure types the run saw, and *Renamed From*
recording where it came from. The column is created the
first time it is needed. The weekly Notes list the renames.

A name Companies has never had in any form, but that *Bond Issuers*
remembers under an earlier one, is a new company with a history: the lookup
sheet lists it as any new company — Bloomberg formulas in place, nothing to
copy from — with *Renamed From* filled from that memory, so the rename is on
record when the row is pasted in. Previous names never become variants: the
old name belongs to the old company.

## Layout

```
src/core/       CoreUtils, CoreReportFormat, CoreImport, CoreCache,
                CoreClean, CoreTextToCol, CoreGlobals
src/reference/  RefAssetMapping
src/weekly/     WeeklyAnalysisGenerate, WeeklyAnalysisLayout, WeeklyAnalysisEmail
src/journey/    Journey, JourneyFormatting, JourneyDashboardTable, JourneyPositionAnalysis
src/delta/      DeltaCalculation, DeltaRevenue
tools/          ToolsInstall (loads a folder of modules), ToolsExposureProbe
                (checks the report's formulas against a VBA pass) — neither
                is part of the workbook
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
only Public procedures with no caller in the source are the zero-argument
entry points a button names — the eight on Home, plus
`InsertRenamedCompanies` behind a button the code itself draws. Two
exceptions carry a comment saying
why they must stay Public: `WriteNoteWeekly`, which `Application.Run` reaches
by name, and `WriteAssetTypeMapping`, whose zero arguments make it bindable to
a button that would not be visible from the source.

### Why WeeklyAnalysisGenerate stays one module

It is 13,900 lines and 233 procedures, and it does not get split, because in
VBA splitting it would cost more than it buys. 229 of those procedures are
Private, along with five Enums and forty-odd Consts. The module is the only
encapsulation boundary the language has — there are no namespaces, and
`Private` means "private to this module", not "private to this concern". Cut
it into five, and every helper the pieces share has to become Public, which
means global: several hundred new names in the one namespace the whole project
shares, including enum members like `RiskStageNDG`.

So the boundary earns its size. What splitting would have bought — being able
to find things — the file order already gives, and the sections run in the
order the report is built: source loading and CSV parsing, the report
sections, risk reference data, certificate basket expansion, entity-name
normalisation, the ranked formulas and the staging table they read, then the
pie, the loan-flow diagram and the notes. `docs/weekly-analysis-generate.md` walks through them.

`archive/JourneyVisualization.bas` is commented out in full. Its charting
procedures were revived inside `JourneyDashboardTable`, which now carries the
same constants and procedure names alongside the Customer Overview and
Historical Events tables. It is kept only for reference.

### Pipelines

**Weekly** — `GenerateWeeklyAnalysis` reads two Home dates, the report date
(`WeeklyEndDate`) and the date the report is compared to (`WeeklyCompareDate`,
normally the previous report's; missing, not a date or not earlier, and the
run stops with a message), loads each date's snapshot with its month-earlier
one and the year-end positions, then builds the report sections in place on
one sheet using the coordinates in `WeeklyAnalysisLayout.Layout`. Every table
carries the compared date's figures beside this report's, under an *As of*
header over bare dates: *Active Lombard Loans* shows year-end, three months
back, the compared date and the current date; *Collateral Breakdown* shows
year-end, the compared date and the current date with their shares, then
`% Change WoW` against the compared date and `% Change YTD` against year-end;
*New Lombard Loans in the Past Month*, *Lombard Loans Ended in the Past
Month* and *Collateral Entered with New NDGs in the Past Month* each show the
compared date's row (amounts and shares, for the entered table) over this
report's, both over the past month, and close with a `% Change WoW` row
between the two. Every ratio is a formula, blank on a base that is zero or
under half a cent — the residue an allocation can leave. The current
snapshot's rows in the overview and the breakdown are highlighted, dark red
(#943634) under white. The overview and the two movement tables stack in the
left column with the same five columns — loans, approved loan, drawn amount,
collateral value — so the three read as one; the notes box sits under them at
the same width, the breakdown column starts one spacer column to their right
with the entered table, the pie and the loan-flow diagram under it, and the
concentration block one spacer column after the breakdown. The loan-flow
diagram is a Sankey drawn from shapes, since Excel has no chart of that
kind: what the collateral lost over the month on the left — the loans
ended, and the positions of the NDGs that stayed that fell — the collateral
categories in the middle, what it gained on the right — the new loans, and
the positions that rose — each band as wide as the collateral that movement
carried out of or into its category, each category bar green, red or grey
by the way its net move went, the pieces grouped as one shape so the email
copies it as one picture.
The concentration block is the bulk of the module: certificate baskets are
expanded recursively into their underlyings, entity names are normalised and
merged (diacritics, legal suffixes, share class suffixes, prefix matching,
manual variants), resolved against the reference sheets, and staged into the
`RiskExposure` table with an account scope flag. The top-10 tables by name,
geography and sector — full portfolio and excluding segregated accounts — are
then worksheet formulas over that table, one per subtable, left live in the
sheet. Beside each rank, a move: where the name stood on the compared date,
as a green or red arrow with the places moved, `=` for none, `new` for a
name that date did not rank. The compared date's ranking is read from its
staged exposure: staging sheets are one per date, *Risk Exposure yyyymmdd*,
so the last run's is on file, and a date that was never staged is staged on
the spot by the same pass the report date gets — alone and quietly: the
reference sheets brought up to that snapshot but no issuer name corrected,
no lookup rows, no notes — so the moves are always there and both dates
resolve names the same way. The reuse question names
the two dates the run needs and which are on file: the answer rebuilds or
reuses those, a date not on file is staged either way, and no other date's
table is ever taken in its place. The run's own table carries the
`RiskExposure` name the formulas use; every other date's is suffixed with
its date. The undated *Risk Exposure* sheet earlier builds wrote is adopted
as a dated one on the first run. `CreateWeeklyEmail` re-exports the finished ranges as HTML — active
loans, breakdown, new loans, loans ended, entered collateral, the pie, the
loan-flow diagram, the concentration tables — and assembles the Outlook message, its intro naming
the date compared to. `docs/weekly-analysis-generate.md` walks
through that module in detail — the staging table's schema, the certificate
recursion, the entity name normalisation, the ranked formula, and the three
separate asset classifications.

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

## The concentration arithmetic is formulas

The exposure concentration section — all twenty-two subtables: three
dimensions, the asset classes `BuildRiskSubtableVisibility` leaves visible in
each, and both account scopes — is worksheet formulas over the `RiskExposure`
staging table. `WriteTopExposureGroup` writes one `LET` per subtable,
calculates it, reads how many rows it spilled, and places the total row
underneath. The formulas stay live: the numbers follow the staging table
without a rerun.

The staging table is a fact table (one row per position × allocated exposure,
with every dimension beside the measure), so the ranking is a `GROUP BY` that
`GROUPBY` does natively: name, value and distinct NDG count, ordered and cut to
ten, with the distinct count written as `LAMBDA(x, COUNTA(UNIQUE(x)))` where
the aggregate goes. What needed care was not whether that is possible but
whether the *semantics* survive the translation. Four are easy to lose:

- **Issuer's denominator is not its numerator.** Issuer receives every row of
  a class. A certificate whose underlying could not be identified at all is
  not dropped but marked — `__UNKNOWN_CERTIFICATE_UNDERLYING__|` while the
  basket is being expanded, then the `Unknown certificate underlying` exposure
  type once staged — and it carries the certificate's whole value at weight 1.
  The ranked list skips it; the share denominator counts it. So part of an
  Issuer table's denominator can never appear in the table, which is why the
  formula carries two masks, one for what ranks and one for what the share
  divides by. Country of Risk and Sector never see those rows, and their two
  masks match.

  That condition should not happen and is not benign: it means the reference
  data could not say what is inside a certificate the portfolio is lending
  against. The analysis still completes — the money is real exposure either
  way — but the report names the certificates in its Notes rather than
  absorbing them silently.
- **The `#NDG` on a total row is a union, not a sum** of the ten counts above
  it.
- **Ties break on name ascending**, case-insensitively, after value descending.
- **The class name in the table is not the label on the report.** Certificates
  is stored as `Certificates (Excl. Protected)`, so the formula matches a list
  of class values rather than one string.

One wrinkle: `GROUPBY` names its value columns in a row of its own, that row
is text, and sorting by value descending therefore carries it to the top and
pushes the tenth name out of the table — `DROP(…, 1)` takes it off before the
sort.

`tools/ToolsExposureProbe.bas` is how this was proved before it replaced the
VBA aggregation, and it is the check to run after touching the formula:
import it, run `BuildExposureProbe`, and the *Exposure Probe* sheet writes all
twenty-two subtables twice — once with the formulas, once with a VBA pass —
with the difference in a column and three totals at the top that should all
be zero.

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
