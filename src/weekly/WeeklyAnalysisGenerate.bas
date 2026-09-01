Attribute VB_Name = "WeeklyAnalysisGenerate"
Option Explicit

' v82: production version of v81. Removes performance timers, diagnostic
' counts, Immediate Window output and the timing-completion message box.
' v81: keeps the single v80 staging pipeline but routes each row to only one
' aggregate set. After the scan, the small Excl. DPM dictionaries are merged
' into Full, avoiding duplicate dictionary updates for every Non-DPM row.
' v80: creates, finalizes, writes and aggregates one unified Risk Exposure
' data set. Account Scope is assigned when each audit row is created; the
' same aggregation scan validates scope and builds Full and Excl. DPM.
' v79: stores Non-DPM and DPM audit rows in one Risk Exposure table with an
' Account Scope field. Notes use one fixed-width merged, wrapped box so their
' text cannot widen the report columns during the final AutoFit.
' v76: caches repeated top-level Certificate RIC expansions. Recursive cycle
' detection and the component-weighting rules remain unchanged.

' v71: replaces per-call temporary dictionaries in Geo-Sec completeness
' checks with direct scans of the short delimited lists. Matching semantics
' remain based on the same exact-name normalization used by v70.

' v70: skips exact duplicate geography observations before they repeat the
' same entry, variant and candidate-ISIN work. It also profiles the remaining
' reference-table and Geo-Sec worksheet phases in more detail.

' v69: carries geography entries forward directly when their canonical key is
' unique and merges candidate ISIN dictionaries only for genuine collisions.
' This preserves the v68 business rules and detailed performance diagnostics.

' v68: adds low-overhead hierarchical timings and operation counts to the
' current v65 optimization branch. Business logic and report output are
' unchanged; detailed diagnostics are printed to the Immediate Window.

' v65: preserves the v64 calculation flow while caching repeated stage
' resolutions, avoiding repeated geography-candidate merges and caching the
' expensive comparison properties used by the existing entity sort.
' Risk Exposure staging sheets are no longer reformatted on each rebuild.

' v64: reduces worksheet-object traffic without changing report behaviour.

' v63: standardizes "Excl." capitalization in user-facing table labels.

' v61: corrects the final risk-table count header to "#NDG".
' v60: restores the concise "Excl. DPM" report titles and shortens the final
' risk-table count header from "NDG Count" to "#Count".
' v59: restores the user-facing "Excl. Segregated Accounts" wording and
' replaces repeated right-hand Certificate detail with one "Same as left"
' row whenever the complete Certificate result is identical to the left table.
' v58: restores the v55 calculation flow for performance. Sovereign Bonds and
' Funds are hidden from Sector Concentration, and Companies is read directly
' from the maintained Geography Final and Sector Final table columns.
' v55: temporarily hides all Overall concentration subsections, hides
' Sovereign Bonds and Funds from Geographic Concentration, places Funds before
' Certificates in all three concentration dimensions, and reads Bond Type.
' v54: reads Companies by header name, uses Fallback Geography when Country of
' Risk is unusable, and uses the new Fallback Sector when Sector is unusable.
' New Geo-Sec Lookup now mirrors the maintained nine-column Companies schema.
' v53: aligns each Overall concentration with its visible asset subsections.
' Hiding an asset class from Issuer, Country or Sector now also removes that
' class from the corresponding Overall denominator and Top 10 calculation.
' v52: removes the final obsolete ISINRelationship assignment left after the
' lean staging refactor, restoring Option Explicit compile compatibility.
' v51: reduces the reusable risk detail tables to audit-relevant columns only.
' Reference lookup fields and constant implementation flags are no longer
' persisted; entity eligibility is derived from Exposure Type during summary.
' v50: safely rebuilds an existing Non-DPM / DPM ListObject without clearing
' the worksheet underneath the live table object. The tables can still be
' reused by header name, and their ListObject identities remain unchanged.
' v49: can rebuild the Non-DPM / DPM risk staging tables or reuse their
' existing data by header name. Reuse skips reference-table updates,
' certificate-file loading, name resolution and staging-table writes; every
' concentration table still aggregates from the same validated stage schema.

Public AssetTypeMapping As Object
Public ReportNotes As Collection

Private Const CERTIFICATE_PATH_NAME As String = "Certificate_Path"
Private Const CERTIFICATE_SEARCH_PREFIX As String = "SearchResults"
Private Const CERTIFICATE_INSTRUMENT_PREFIX As String = "InstrumentList"
Private Const MISSING_CERTIFICATE_RIC_PREFIX As String = "__MISSING_RIC__|"
'
' A certificate whose underlying could not be named at all.  The prefix marks
' the component while the basket is being expanded; it is stripped before the
' row reaches the staging table, where UNKNOWN_UNDERLYING_TYPE identifies it
' instead.  Neither is a category of exposure - both say the parse failed.
'
Private Const UNKNOWN_UNDERLYING_PREFIX As String = _
    "__UNKNOWN_CERTIFICATE_UNDERLYING__|"
Private Const CERTIFICATE_MAX_BASKET_DEPTH As Long = 12
Private Const EQUITY_NAMES_SHEET As String = "Equity Names"
Private Const BOND_ISSUERS_SHEET As String = "Bond Issuers"
Private Const FUND_PARENT_COMPANIES_SHEET As String = _
    "Fund Parent Companies"
Private Const UNMAPPED_EQUITIES_TABLE As String = "UnmappedEquities"
Private Const BOND_ISSUERS_TABLE As String = "BondIssuers"
Private Const FUND_PARENT_COMPANIES_TABLE As String = _
    "FundParentCompanies"
Private Const FUNDS_TABLE As String = "Funds"
Private Const CORPORATE_BONDS_CLASS As String = "Corporate Bonds"
Private Const CORPORATE_BONDS_DISPLAY_NAME As String = _
    "Corporate Bonds (Excl. UniCredit SpA)"
Private Const SOVEREIGN_BONDS_CLASS As String = "Sovereign Bonds"
Private Const BOND_ISSUER_TYPE_CORPORATE As String = "Corporate"
Private Const BOND_ISSUER_TYPE_SOVEREIGN As String = "Sovereign"
Private Const COMPANIES_SHEET As String = "Companies"
Private Const COMPANIES_TABLE As String = "Companies"
Private Const COUNTRIES_SHEET As String = "Countries"
' Excel worksheet names cannot contain "/". The hyphen preserves the
' requested wording while remaining a valid sheet name.
Private Const GEO_SEC_LOOKUP_SHEET As String = "New Geo-Sec Lookup"
Private Const LEGACY_GEOGRAPHY_LOOKUP_SHEET As String = _
    "Geography Lookup"
Private Const NAME_VARIANTS_SHEET As String = "Name Variants"
Private Const ENTITY_PREFIX_MIN_LENGTH As Long = 12
Private Const ENTITY_PREFIX_MIN_RATIO As Double = 0.65
Private Const ENTITY_LEGAL_SUFFIX_ROOT_MIN_LENGTH As Long = 3
Private Const GEO_ISIN_PRIORITY_ISSUED_SECURITY As Long = 1
Private Const GEO_ISIN_PRIORITY_UNDERLYING As Long = 2
Private Const GEO_ISIN_PRIORITY_MANAGED_FUND As Long = 3
Private Const TOP_NAME_COUNT As Long = 10
Private Const OTHER_RISK_DIMENSION As String = "Others"
Private Const UNKNOWN_UNDERLYING_TYPE As String = _
    "Unknown certificate underlying"
Private Const NON_DPM_SCOPE As String = "Non-DPM"

Private Const UNKNOWN_UNDERLYING_NOTE_LIMIT As Long = 10

Private Const RISK_FORMULA_INDENT As String = "    "
Private Const RISK_BIND_WIDTH As Long = 6
Private Const POSITION_FILE_SUFFIX As String = _
    "_Lombard_Loans_ITA_Positions.csv"
Private Const ACCOUNT_FILE_SUFFIX As String = _
    "_Lombard_Loans_ITA_Accounts.csv"
Private Const RISK_STAGE_SHEET As String = "Risk Exposure"
Private Const RISK_STAGE_TABLE As String = "RiskExposure"
Private Const LEGACY_RISK_STAGE_NON_DPM_SHEET As String = _
    "Risk Exposure - Non-DPM"
Private Const LEGACY_RISK_STAGE_DPM_SHEET As String = _
    "Risk Exposure - DPM"

Private Enum RiskStageField

    RiskStageSourceRow = 1
    RiskStageNDG = 2
    RiskStageProductISIN = 3
    RiskStageSecurityName = 4
    RiskStageOriginalAssetType = 5
    RiskStageAdditionalComment = 6
    RiskStageAssetClass = 7
    RiskStageExposureName = 8
    RiskStageExposureType = 9
    RiskStageAllocationWeight = 10
    RiskStagePositionValue = 11
    RiskStageAllocatedValue = 12
    RiskStageGeography = 13
    RiskStageSector = 14
    RiskStageResolutionSource = 15
    RiskStageAccountScope = 16

End Enum

Private Const RISK_STAGE_FIELD_COUNT As Long = 16

Private Enum WeeklyPositionField

    WeeklyPosNDG = 1
    WeeklyPosISIN = 2
    WeeklyPosSecurityName = 3
    WeeklyPosAssetType = 4
    WeeklyPosPositionValue = 5
    WeeklyPosIssuer = 6
    WeeklyPosAdditionalComment = 7

End Enum

Private Enum RiskPositionCacheField

    RiskPositionCacheReportingAssetClass = 1
    RiskPositionCacheIsDPM = 2

End Enum

Private Const RISK_POSITION_CACHE_FIELD_COUNT As Long = 2

Private Enum WeeklyAccountField

    WeeklyAccountNDG = 1
    WeeklyAccountApproved = 2
    WeeklyAccountDrawn = 3

End Enum

Private WeeklyPositionCache As Object
Private WeeklyAccountCache As Object

'
' Counted while the staging rows are built, so a run that stages nothing can
' say whether the positions were missing or their exposure names were.
'
Private RiskStagePositionsScanned As Long
Private RiskStageRowsDropped As Long

Private Enum IssuerPlaceholderMode

    PlaceholderFromISIN = 1
    PlaceholderFromSecurityName = 2

End Enum

Private Function BuildRiskSubtableVisibility() As Object

    Dim Visibility As Object

    Set Visibility = CreateObject("Scripting.Dictionary")
    Visibility.CompareMode = vbTextCompare

    ' Comment out one line below to hide that subsection from both the full
    ' and Excl. DPM view. The same asset class is then also
    ' excluded from the corresponding Overall calculation. The Overall line
    ' controls whether
    ' the resulting Overall subsection itself is displayed.

    Visibility.Add "Issuer|Equity", True
    Visibility.Add "Issuer|Corporate Bonds", True
    Visibility.Add "Issuer|Sovereign Bonds", True
    Visibility.Add "Issuer|Funds", True
    Visibility.Add "Issuer|Certificates", True
    ' Visibility.Add "Issuer|Overall", True

    Visibility.Add "Country|Equity", True
    Visibility.Add "Country|Corporate Bonds", True
    ' Visibility.Add "Country|Sovereign Bonds", True
    ' Visibility.Add "Country|Funds", True
    Visibility.Add "Country|Certificates", True
    ' Visibility.Add "Country|Overall", True

    Visibility.Add "Sector|Equity", True
    Visibility.Add "Sector|Corporate Bonds", True
    ' Visibility.Add "Sector|Sovereign Bonds", True
    ' Visibility.Add "Sector|Funds", True
    Visibility.Add "Sector|Certificates", True
    ' Visibility.Add "Sector|Overall", True

    Set BuildRiskSubtableVisibility = Visibility

End Function

Private Function RiskSubtableIsVisible( _
    ByVal Visibility As Object, _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String) As Boolean

    Dim VisibilityKey As String

    If Visibility Is Nothing Then Exit Function

    VisibilityKey = DimensionKey & "|" & AssetClassKey

    If Not Visibility.Exists(VisibilityKey) Then Exit Function

    RiskSubtableIsVisible = CBool(Visibility(VisibilityKey))

End Function

Private Function NextRiskSectionRow( _
    ByVal LeftNextRow As Long, _
    ByVal RightNextRow As Long) As Long

    If LeftNextRow >= RightNextRow Then

        NextRiskSectionRow = _
            LeftNextRow + Layout.RiskSectionGapRows

    Else

        NextRiskSectionRow = _
            RightNextRow + Layout.RiskSectionGapRows

    End If

End Function

Private Function EuroNumberFormat() As String

    EuroNumberFormat = ChrW(&H20AC) & "#,##0.00"

End Function

Private Sub InitialiseWeeklySourceCache()

    Set WeeklyPositionCache = _
        CreateObject("Scripting.Dictionary")

    WeeklyPositionCache.CompareMode = vbTextCompare

    Set WeeklyAccountCache = _
        CreateObject("Scripting.Dictionary")

    WeeklyAccountCache.CompareMode = vbTextCompare

End Sub

Private Sub ClearWeeklySourceCache()

    Set WeeklyPositionCache = Nothing
    Set WeeklyAccountCache = Nothing

End Sub

Private Function WeeklySourceFilePath( _
    ByVal SnapshotDate As Date, _
    ByVal FileSuffix As String) As String

    WeeklySourceFilePath = _
        PathSelection() & _
        GetDateCode(SnapshotDate) & _
        FileSuffix

End Function

Private Function ReadWeeklySourceText( _
    ByVal FilePath As String) As String

    Dim FileNumber As Integer
    Dim FileSize As Long
    Dim FileText As String

    FileNumber = FreeFile

    On Error GoTo ReadFailed

    Open FilePath For Binary Access Read As #FileNumber

    FileSize = LOF(FileNumber)

    If FileSize > 0 Then

        FileText = Space$(FileSize)
        Get #FileNumber, , FileText

    End If

    Close #FileNumber

    ReadWeeklySourceText = FileText

    Exit Function

ReadFailed:

    On Error Resume Next
    Close #FileNumber
    On Error GoTo 0

    Err.Raise _
        vbObjectError + 1910, _
        "ReadWeeklySourceText", _
        "Source file could not be read:" & vbCrLf & _
        FilePath

End Function

Private Function WeeklySourceLines( _
    ByVal FilePath As String) As Variant

    Dim FileText As String

    FileText = ReadWeeklySourceText(FilePath)

    FileText = Replace(FileText, vbCrLf, vbLf)
    FileText = Replace(FileText, vbCr, vbLf)

    WeeklySourceLines = Split(FileText, vbLf)

End Function

Private Function CleanWeeklyCsvField( _
    ByVal FieldValue As Variant) As String

    Dim Result As String

    If IsError(FieldValue) Then Exit Function
    If IsNull(FieldValue) Then Exit Function
    If IsEmpty(FieldValue) Then Exit Function

    Result = Trim$(CStr(FieldValue))

    If Len(Result) >= 2 Then

        If Left$(Result, 1) = Chr$(34) And _
           Right$(Result, 1) = Chr$(34) Then

            Result = Mid$(Result, 2, Len(Result) - 2)
            Result = Replace(Result, Chr$(34) & Chr$(34), Chr$(34))

        End If

    End If

    Result = Replace(Result, ChrW(&HFEFF), "")

    ' Binary reads preserve an UTF-8 BOM as three ANSI characters.
    ' Test the byte values so the VBA source itself remains ASCII-safe.
    If Len(Result) >= 3 Then

        If AscW(Mid$(Result, 1, 1)) = 239 And _
           AscW(Mid$(Result, 2, 1)) = 187 And _
           AscW(Mid$(Result, 3, 1)) = 191 Then

            Result = Mid$(Result, 4)

        End If

    End If

    CleanWeeklyCsvField = Trim$(Result)

End Function

Private Function WeeklyCsvField( _
    ByRef Fields As Variant, _
    ByVal FieldIndex As Long) As String

    If FieldIndex < LBound(Fields) Or _
       FieldIndex > UBound(Fields) Then Exit Function

    WeeklyCsvField = _
        CleanWeeklyCsvField(Fields(FieldIndex))

End Function

Private Function WeeklyCsvDouble( _
    ByVal InputValue As Variant) As Double

    Dim NumberText As String
    Dim LastComma As Long
    Dim LastDot As Long
    Dim IsNegative As Boolean

    If IsError(InputValue) Then Exit Function
    If IsNull(InputValue) Then Exit Function
    If IsEmpty(InputValue) Then Exit Function

    ' Values coming from CSV are strings. Do not apply CDbl to those before
    ' normalising their decimal/thousands separators: CDbl is locale-sensitive.
    Select Case VarType(InputValue)

        Case vbByte, vbInteger, vbLong, vbSingle, _
             vbDouble, vbCurrency, vbDecimal

            WeeklyCsvDouble = CDbl(InputValue)

            Exit Function

    End Select

    NumberText = CleanWeeklyCsvField(InputValue)

    If NumberText = "" Then Exit Function

    NumberText = Replace(NumberText, ChrW(&H20AC), "")
    NumberText = Replace(NumberText, ChrW(&HA0), "")
    NumberText = Replace(NumberText, " ", "")
    NumberText = Replace(NumberText, "%", "")

    If Left$(NumberText, 1) = "(" And _
       Right$(NumberText, 1) = ")" Then

        IsNegative = True
        NumberText = Mid$(NumberText, 2, Len(NumberText) - 2)

    ElseIf Right$(NumberText, 1) = "-" Then

        IsNegative = True
        NumberText = Left$(NumberText, Len(NumberText) - 1)

    End If

    LastComma = InStrRev(NumberText, ",")
    LastDot = InStrRev(NumberText, ".")

    If LastComma > 0 And LastDot > 0 Then

        If LastComma > LastDot Then

            NumberText = Replace(NumberText, ".", "")
            NumberText = Replace(NumberText, ",", ".")

        Else

            NumberText = Replace(NumberText, ",", "")

        End If

    ElseIf LastComma > 0 Then

        NumberText = Replace(NumberText, ",", ".")

    End If

    WeeklyCsvDouble = Val(NumberText)

    If IsNegative Then

        WeeklyCsvDouble = -Abs(WeeklyCsvDouble)

    End If

End Function

Private Function FindWeeklyHeaderIndex( _
    ByRef HeaderFields As Variant, _
    ByVal CandidateHeaders As Variant, _
    ByVal FallbackIndex As Long, _
    ByVal FilePath As String, _
    Optional ByVal Required As Boolean = True) As Long

    Dim Candidate As Variant
    Dim HeaderIndex As Long
    Dim HeaderText As String

    FindWeeklyHeaderIndex = -1

    For HeaderIndex = LBound(HeaderFields) To UBound(HeaderFields)

        HeaderText = _
            NormalizeHeader( _
                CleanWeeklyCsvField( _
                    HeaderFields(HeaderIndex)))

        For Each Candidate In CandidateHeaders

            If HeaderText = _
                NormalizeHeader(CStr(Candidate)) Then

                FindWeeklyHeaderIndex = HeaderIndex

                Exit Function

            End If

        Next Candidate

    Next HeaderIndex

    If FallbackIndex >= LBound(HeaderFields) And _
       FallbackIndex <= UBound(HeaderFields) Then

        FindWeeklyHeaderIndex = FallbackIndex

        Exit Function

    End If

    If Required Then

        Err.Raise _
            vbObjectError + 1911, _
            "FindWeeklyHeaderIndex", _
            "A required source column could not be found in:" & _
            vbCrLf & FilePath

    End If

End Function

Private Function WeeklyDataHasRows( _
    ByRef Data As Variant) As Boolean

    Dim FirstRow As Long
    Dim LastRow As Long

    If Not IsArray(Data) Then Exit Function

    On Error GoTo NoRows

    FirstRow = LBound(Data, 1)
    LastRow = UBound(Data, 1)

    WeeklyDataHasRows = (LastRow >= FirstRow)

    Exit Function

NoRows:

    Err.Clear

End Function

Private Function LoadWeeklyPositionData( _
    ByVal SnapshotDate As Date) As Variant

    Dim FilePath As String
    Dim Lines As Variant
    Dim HeaderFields As Variant
    Dim Fields As Variant
    Dim Data() As Variant

    Dim ColNDG As Long
    Dim ColISIN As Long
    Dim ColSecurityName As Long
    Dim ColAssetType As Long
    Dim ColPositionValue As Long
    Dim ColIssuer As Long
    Dim ColAdditionalComment As Long

    Dim DataRowCount As Long
    Dim OutputRow As Long
    Dim r As Long

    FilePath = _
        WeeklySourceFilePath( _
            SnapshotDate, _
            POSITION_FILE_SUFFIX)

    If WeeklyPositionCache Is Nothing Then

        InitialiseWeeklySourceCache

    End If

    If WeeklyPositionCache.Exists(FilePath) Then

        LoadWeeklyPositionData = _
            WeeklyPositionCache(FilePath)

        Exit Function

    End If

    If Dir(FilePath) = "" Then

        MissingFiles = MissingFiles & vbCrLf & FilePath

        Exit Function

    End If

    Lines = WeeklySourceLines(FilePath)

    If UBound(Lines) < 1 Then Exit Function

    HeaderFields = Split(CStr(Lines(0)), ";")

    ColNDG = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("NDG"), _
            1, _
            FilePath)

    ColISIN = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("ISIN"), _
            3, _
            FilePath)

    ColSecurityName = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Security Name"), _
            4, _
            FilePath)

    ColAssetType = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array( _
                "Asset Type / Classification", _
                "Asset Type", _
                "Classification"), _
            5, _
            FilePath)

    ColPositionValue = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Position Value"), _
            12, _
            FilePath)

    ColIssuer = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Issuer"), _
            18, _
            FilePath, _
            False)

    ColAdditionalComment = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Additional Comment"), _
            19, _
            FilePath, _
            False)

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            DataRowCount = DataRowCount + 1

        End If

    Next r

    If DataRowCount = 0 Then Exit Function

    ReDim Data( _
        1 To DataRowCount, _
        1 To WeeklyPosAdditionalComment)

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            Fields = Split(CStr(Lines(r)), ";")
            OutputRow = OutputRow + 1

            Data(OutputRow, WeeklyPosNDG) = _
                WeeklyCsvField(Fields, ColNDG)

            Data(OutputRow, WeeklyPosISIN) = _
                WeeklyCsvField(Fields, ColISIN)

            Data(OutputRow, WeeklyPosSecurityName) = _
                WeeklyCsvField(Fields, ColSecurityName)

            Data(OutputRow, WeeklyPosAssetType) = _
                WeeklyCsvField(Fields, ColAssetType)

            Data(OutputRow, WeeklyPosPositionValue) = _
                WeeklyCsvDouble( _
                    WeeklyCsvField( _
                        Fields, _
                        ColPositionValue))

            Data(OutputRow, WeeklyPosIssuer) = _
                WeeklyCsvField(Fields, ColIssuer)

            Data(OutputRow, WeeklyPosAdditionalComment) = _
                WeeklyCsvField( _
                    Fields, _
                    ColAdditionalComment)

        End If

    Next r

    WeeklyPositionCache.Add FilePath, Data

    LoadWeeklyPositionData = Data

End Function

Private Function LoadWeeklyAccountData( _
    ByVal SnapshotDate As Date) As Variant

    Dim FilePath As String
    Dim Lines As Variant
    Dim HeaderFields As Variant
    Dim Fields As Variant
    Dim Data() As Variant

    Dim ColNDG As Long
    Dim ColApproved As Long
    Dim ColDrawn As Long

    Dim DataRowCount As Long
    Dim OutputRow As Long
    Dim r As Long

    FilePath = _
        WeeklySourceFilePath( _
            SnapshotDate, _
            ACCOUNT_FILE_SUFFIX)

    If WeeklyAccountCache Is Nothing Then

        InitialiseWeeklySourceCache

    End If

    If WeeklyAccountCache.Exists(FilePath) Then

        LoadWeeklyAccountData = _
            WeeklyAccountCache(FilePath)

        Exit Function

    End If

    If Dir(FilePath) = "" Then

        MissingFiles = MissingFiles & vbCrLf & FilePath

        Exit Function

    End If

    Lines = WeeklySourceLines(FilePath)

    If UBound(Lines) < 1 Then Exit Function

    HeaderFields = Split(CStr(Lines(0)), ";")

    ColNDG = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("NDG"), _
            0, _
            FilePath)

    ColApproved = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Max Approved Loan", "Approved"), _
            3, _
            FilePath)

    ColDrawn = _
        FindWeeklyHeaderIndex( _
            HeaderFields, _
            Array("Drawn Amount", "Drawn"), _
            4, _
            FilePath)

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            DataRowCount = DataRowCount + 1

        End If

    Next r

    If DataRowCount = 0 Then Exit Function

    ReDim Data( _
        1 To DataRowCount, _
        1 To WeeklyAccountDrawn)

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            Fields = Split(CStr(Lines(r)), ";")
            OutputRow = OutputRow + 1

            Data(OutputRow, WeeklyAccountNDG) = _
                WeeklyCsvField(Fields, ColNDG)

            Data(OutputRow, WeeklyAccountApproved) = _
                WeeklyCsvDouble( _
                    WeeklyCsvField( _
                        Fields, _
                        ColApproved))

            Data(OutputRow, WeeklyAccountDrawn) = _
                WeeklyCsvDouble( _
                    WeeklyCsvField( _
                        Fields, _
                        ColDrawn))

        End If

    Next r

    WeeklyAccountCache.Add FilePath, Data

    LoadWeeklyAccountData = Data

End Function

Private Function GetAccountNDGDictionary( _
    ByRef AccountData As Variant) As Object

    Dim Dict As Object
    Dim NDG As String
    Dim r As Long

    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare

    If WeeklyDataHasRows(AccountData) Then

        For r = LBound(AccountData, 1) To UBound(AccountData, 1)

            NDG = _
                CleanWeeklyCsvField( _
                    AccountData( _
                        r, _
                        WeeklyAccountNDG))

            If NDG <> "" Then Dict(NDG) = True

        Next r

    End If

    Set GetAccountNDGDictionary = Dict

End Function

Private Sub CalculateWeeklyPortfolioStats( _
    ByVal SnapshotDate As Date, _
    ByRef LoanCount As Long, _
    ByRef DrawnAmount As Double, _
    ByRef ApprovedAmount As Double)

    Dim AccountData As Variant
    Dim NDGs As Object
    Dim NDG As String
    Dim r As Long

    AccountData = LoadWeeklyAccountData(SnapshotDate)

    If Not WeeklyDataHasRows(AccountData) Then Exit Sub

    Set NDGs = CreateObject("Scripting.Dictionary")
    NDGs.CompareMode = vbTextCompare

    For r = LBound(AccountData, 1) To UBound(AccountData, 1)

        NDG = _
            CleanWeeklyCsvField( _
                AccountData( _
                    r, _
                    WeeklyAccountNDG))

        If NDG <> "" Then NDGs(NDG) = True

        ApprovedAmount = _
            ApprovedAmount + _
            CDbl( _
                AccountData( _
                    r, _
                    WeeklyAccountApproved))

        DrawnAmount = _
            DrawnAmount + _
            CDbl( _
                AccountData( _
                    r, _
                    WeeklyAccountDrawn))

    Next r

    LoanCount = NDGs.Count

End Sub

Public Sub GenerateWeeklyAnalysis()

    Dim OldNoteHandler As String
    Dim CurrentDate As Date
    Dim ComparisonDate As Date
    Dim YTDDate As Date

    Dim ws As Worksheet

    Dim AccountsCurrent As Variant
    Dim AccountsCompare As Variant
    Dim PositionsCurrent As Variant
    Dim PositionsCompare As Variant
    Dim PositionsYTD As Variant

    Dim UnknownAssets As Object

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    OldNoteHandler = NoteHandler
    NoteHandler = "WriteNoteWeekly"
    Set ReportNotes = New Collection
    MissingFiles = ""
    ResetSheetOverwriteDecision

    InitialiseWeeklySourceCache

    Set UnknownAssets = CreateObject("Scripting.Dictionary")
    Set AssetTypeMapping = CreateObject("Scripting.Dictionary")

CurrentDate = _
    Worksheets("Home").Range("WeeklyEndDate").Value

If Not SourceFileExists(CurrentDate, "POSITIONS") _
   Or Not SourceFileExists(CurrentDate, "ACCOUNTS") Then

    Fatal _
        "Analysis end date source files not found:" & vbCrLf & _
        Format(CurrentDate, "dd/mm/yyyy")

End If

ComparisonDate = _
    ResolveComparisonDate( _
        GetComparisonDate(CurrentDate))

YTDDate = _
    ResolveComparisonDate( _
        GetYTDDate(CurrentDate))

    Set ws = CreateOrReplaceSheet("Weekly Analysis")

    If ws Is Nothing Then GoTo ExitRoutine

    ws.Cells.Clear
    ws.UsedRange.UnMerge

    AccountsCurrent = LoadWeeklyAccountData(CurrentDate)
    AccountsCompare = LoadWeeklyAccountData(ComparisonDate)
    PositionsCurrent = LoadWeeklyPositionData(CurrentDate)
    PositionsCompare = LoadWeeklyPositionData(ComparisonDate)
    PositionsYTD = LoadWeeklyPositionData(YTDDate)
    
    InitializeLayout
    BuildHeader ws, CurrentDate
    BuildPortfolioSection ws, CurrentDate, ComparisonDate

    If WeeklyDataHasRows(PositionsCurrent) And _
       WeeklyDataHasRows(PositionsYTD) Then

        BuildCollateralBreakdown _
            ws, _
            PositionsCurrent, _
            PositionsYTD, _
            UnknownAssets

        WriteAssetTypeMapping

    End If

    If WeeklyDataHasRows(PositionsCurrent) Then

        BuildRiskGranularitySection ws, PositionsCurrent

    End If
    
    If WeeklyDataHasRows(AccountsCurrent) And _
       WeeklyDataHasRows(AccountsCompare) And _
       WeeklyDataHasRows(PositionsCurrent) Then

        BuildNewLoansSection _
            ws, _
            AccountsCurrent, _
            AccountsCompare, _
            PositionsCurrent

    End If
    
    If WeeklyDataHasRows(AccountsCurrent) And _
       WeeklyDataHasRows(AccountsCompare) Then

        BuildEndedLoansSection _
            ws, _
            AccountsCurrent, _
            AccountsCompare, _
            PositionsCompare

    End If
        
    If WeeklyDataHasRows(AccountsCurrent) And _
       WeeklyDataHasRows(AccountsCompare) And _
       WeeklyDataHasRows(PositionsCurrent) Then

        BuildEnteredCollateralSection _
            ws, _
            AccountsCurrent, _
            AccountsCompare, _
            PositionsCurrent, _
            UnknownAssets

    End If

    With ws.UsedRange

        .Font.name = "Aptos Display"
        .VerticalAlignment = xlCenter
        .Columns.AutoFit
        .Rows.RowHeight = 16
    
    End With
    
    '
    ' Main title
    '
    ws.Rows(Layout.HeaderRow).AutoFit
    
    '
    ' As of row
    '
    ws.Rows(Layout.HeaderRow + 1).AutoFit

    If Not ws Is Nothing Then

        BuildNotes ws
        FormatNotesBox ws

    End If

    CreateWeeklyEmailButton ws

    If MissingFiles <> "" Then

        MsgBox "Weekly Analysis completed with warnings." & vbCrLf & vbCrLf & "Missing source files:" & vbCrLf & MissingFiles, vbExclamation

    End If

ExitRoutine:

    ClearWeeklySourceCache

    NoteHandler = OldNoteHandler

    ResetExcel

    If Not ws Is Nothing Then

        ws.Activate

    End If

    Exit Sub

ErrorHandler:

    ResetExcel

    MsgBox _
        Err.Description, _
        vbCritical, _
        "Weekly Analysis"

    GoTo ExitRoutine

End Sub

Private Sub WriteSectionTitle(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal ColNo As Long, ByVal Width As Long, ByVal Title As String)

    With ws.Range(ws.Cells(RowNo, ColNo), ws.Cells(RowNo, ColNo + Width - 1))

        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Value = Title
        
    End With
    
End Sub

Private Sub BuildHeader( _
    ByVal ws As Worksheet, _
    ByVal ReportDate As Date)

    With ws.Range( _
        ws.Cells(Layout.HeaderRow, Layout.HeaderCol), _
        ws.Cells(Layout.HeaderRow, Layout.HeaderCol + 17))

        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        .Value = "Weekly Analysis Lombard Loan"

        .Font.Bold = True
        .Font.Size = 20

    End With

    With ws.Range( _
        ws.Cells(Layout.HeaderRow + 1, Layout.HeaderCol), _
        ws.Cells(Layout.HeaderRow + 1, Layout.HeaderCol + 17))

        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        .Value = "As of " & Format(ReportDate, "dd/mm/yyyy")

        .Font.Size = 14

    End With

End Sub

Private Sub BuildPortfolioSection( _
    ByVal ws As Worksheet, _
    ByVal CurrentDate As Date, _
    ByVal ComparisonDate As Date)

    Dim Date2 As Date
    Dim Date3 As Date
    Dim YTDDate As Date

    Dim RowDates As Variant
    Dim ShowYTD As Boolean

    Dim FirstDataRow As Long
    Dim LastDataRow As Long

    Dim r As Long
    Dim c As Long
    Dim i As Long

    r = Layout.PortfolioRow
    c = Layout.PortfolioCol

    Date2 = ResolveComparisonDate(GetComparisonDate(CurrentDate, 2))
    Date3 = ResolveComparisonDate(GetComparisonDate(CurrentDate, 3))
    YTDDate = ResolveComparisonDate(GetYTDDate(CurrentDate))

    '
    ' Three months of history already reaches into the previous year, so a
    ' separate year-end row would only repeat what the table shows anyway.
    '
    ShowYTD = (Year(Date3) = Year(CurrentDate))

    If ShowYTD Then
        RowDates = Array(YTDDate, Date3, Date2, ComparisonDate, CurrentDate)
    Else
        RowDates = Array(Date3, Date2, ComparisonDate, CurrentDate)
    End If

    FirstDataRow = r + 2
    LastDataRow = FirstDataRow + UBound(RowDates)

    WriteSectionTitle ws, r, c, 4, "Overview"

    ws.Cells(r + 1, c).Value = "Date"
    ws.Cells(r + 1, c + 1).Value = "Loans"
    ws.Cells(r + 1, c + 2).Value = "Drawn"
    ws.Cells(r + 1, c + 3).Value = "Approved"

    For i = 0 To UBound(RowDates)
        WritePortfolioRow ws, FirstDataRow + i, c, CDate(RowDates(i))
    Next i

    If ShowYTD Then
        ws.Cells(FirstDataRow, c).Value = "YE " & Year(CurrentDate) - 1
    End If

    ws.Range( _
        ws.Cells(FirstDataRow, c), _
        ws.Cells(LastDataRow, c)).NumberFormat = "dd/mm/yyyy"

    ws.Range( _
        ws.Cells(FirstDataRow, c + 2), _
        ws.Cells(LastDataRow, c + 3)).NumberFormat = EuroNumberFormat()

    FormatReportTable _
        ws.Range(ws.Cells(r + 1, c), ws.Cells(LastDataRow, c + 3)), _
        1

    FormatFirstColumn ws, r + 1, LastDataRow, c

End Sub

Private Sub WritePortfolioRow( _
    ByVal ws As Worksheet, _
    ByVal TargetRow As Long, _
    ByVal StartCol As Long, _
    ByVal RefDate As Date)

    Dim LoanCount As Long
    Dim DrawnAmount As Double
    Dim ApprovedAmount As Double

    CalculateWeeklyPortfolioStats _
        RefDate, _
        LoanCount, _
        DrawnAmount, _
        ApprovedAmount

    ws.Cells(TargetRow, StartCol).Value = RefDate
    ws.Cells(TargetRow, StartCol + 1).Value = LoanCount
    ws.Cells(TargetRow, StartCol + 2).Value = DrawnAmount
    ws.Cells(TargetRow, StartCol + 3).Value = ApprovedAmount

End Sub

'
' The collateral categories in report order: dictionary key first, then the
' column header. Every collateral table, total and pie slice is driven from
' this one list, so a category is added or renamed in a single place.
'
Private Function CollateralCategories() As Variant

    CollateralCategories = Array( _
        Array("Certificates", "Certificates"), _
        Array("Cash", "Cash"), _
        Array("Funds", "Mutual Funds & ETF"), _
        Array("Insurance", "Insurance"), _
        Array("GP", "GP"), _
        Array("Bonds", "Bonds"), _
        Array("Equity", "Equity"), _
        Array("Non Eligible Asset", "Non-Eligible"))

End Function

Private Function CollateralCategoryCount() As Long

    CollateralCategoryCount = UBound(CollateralCategories()) + 1

End Function

'
' Pie slice colours, in the same order as CollateralCategories.
'
Private Function CollateralSliceColors() As Variant

    CollateralSliceColors = Array( _
        RGB(95, 125, 150), _
        RGB(130, 155, 170), _
        RGB(170, 185, 190), _
        RGB(215, 205, 185), _
        RGB(195, 170, 145), _
        RGB(170, 135, 120), _
        RGB(165, 105, 120), _
        RGB(120, 120, 120))

End Function

Private Function NewCollateralDictionary() As Object

    Dim Amounts As Object
    Dim Category As Variant

    Set Amounts = CreateObject("Scripting.Dictionary")

    For Each Category In CollateralCategories()
        Amounts(Category(0)) = 0
    Next Category

    Set NewCollateralDictionary = Amounts

End Function

Private Function CollateralTotal(ByVal Amounts As Object) As Double

    Dim Category As Variant

    For Each Category In CollateralCategories()
        CollateralTotal = CollateralTotal + Amounts(Category(0))
    Next Category

End Function

'
' The four row shapes every collateral table is built from. Each writes the
' category columns only; the caller owns the row label in LeftCol.
'
Private Sub WriteCollateralHeaders( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal LeftCol As Long)

    Dim Categories As Variant
    Dim i As Long

    Categories = CollateralCategories()

    For i = 0 To UBound(Categories)
        ws.Cells(RowNo, LeftCol + 1 + i).Value = Categories(i)(1)
    Next i

End Sub

Private Sub WriteCollateralAmounts( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal LeftCol As Long, _
    ByVal Amounts As Object)

    Dim Categories As Variant
    Dim i As Long

    Categories = CollateralCategories()

    For i = 0 To UBound(Categories)
        ws.Cells(RowNo, LeftCol + 1 + i).Value = Amounts(Categories(i)(0))
    Next i

End Sub

Private Sub WriteCollateralShares( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal LeftCol As Long, _
    ByVal Amounts As Object, _
    ByVal Total As Double)

    Dim Categories As Variant
    Dim i As Long

    If Total = 0 Then Exit Sub

    Categories = CollateralCategories()

    For i = 0 To UBound(Categories)
        ws.Cells(RowNo, LeftCol + 1 + i).Value = _
            Amounts(Categories(i)(0)) / Total
    Next i

End Sub

'
' Percentage change against a base period. A category the base period did not
' hold has no meaningful change, so its cell is left empty.
'
Private Sub WriteCollateralChange( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal LeftCol As Long, _
    ByVal Amounts As Object, _
    ByVal BaseAmounts As Object)

    Dim Categories As Variant
    Dim BaseValue As Double
    Dim i As Long

    Categories = CollateralCategories()

    For i = 0 To UBound(Categories)

        BaseValue = BaseAmounts(Categories(i)(0))

        If BaseValue <> 0 Then
            ws.Cells(RowNo, LeftCol + 1 + i).Value = _
                (Amounts(Categories(i)(0)) - BaseValue) / BaseValue
        End If

    Next i

End Sub

Private Sub BuildCollateralBreakdown( _
    ByVal ws As Worksheet, _
    ByRef CurrentPositions As Variant, _
    ByRef YTDPositions As Variant, _
    ByRef UnknownAssets As Object)

    Dim DictCurrent As Object
    Dim DictYTD As Object

    Dim TotalCurrent As Double
    Dim TotalYTD As Double

    Dim ReportDate As Date

    Dim r As Long
    Dim c As Long
    Dim LastCol As Long

    If Not WeeklyDataHasRows(CurrentPositions) Then Exit Sub
    If Not WeeklyDataHasRows(YTDPositions) Then Exit Sub

    r = Layout.BreakdownRow
    c = Layout.BreakdownCol
    LastCol = c + CollateralCategoryCount()

    ReportDate = Worksheets("Home").Range("WeeklyEndDate").Value

    Set DictCurrent = _
        BuildCollateralDictionary(CurrentPositions, UnknownAssets)
    Set DictYTD = _
        BuildCollateralDictionary(YTDPositions, UnknownAssets)

    TotalCurrent = CollateralTotal(DictCurrent)
    TotalYTD = CollateralTotal(DictYTD)

    WriteSectionTitle _
        ws, r, c, _
        CollateralCategoryCount() + 1, _
        "Collateral Breakdown"

    ws.Cells(r + 1, c).Value = "Date"
    WriteCollateralHeaders ws, r + 1, c

    ws.Cells(r + 2, c).Value = "YE " & Year(ReportDate) - 1
    WriteCollateralAmounts ws, r + 2, c, DictYTD

    ws.Cells(r + 3, c).Value = "% of Portfolio"
    WriteCollateralShares ws, r + 3, c, DictYTD, TotalYTD

    ws.Cells(r + 4, c).Value = ReportDate
    WriteCollateralAmounts ws, r + 4, c, DictCurrent

    ws.Cells(r + 5, c).Value = "% of Portfolio"
    WriteCollateralShares ws, r + 5, c, DictCurrent, TotalCurrent

    ws.Cells(r + 6, c).Value = "% Change YTD"
    WriteCollateralChange ws, r + 6, c, DictCurrent, DictYTD

    '
    ' Formatting
    '

    ws.Range(ws.Cells(r + 2, c), ws.Cells(r + 4, c)).NumberFormat = "dd/mm/yyyy"

    ws.Range( _
        ws.Cells(r + 2, c + 1), _
        ws.Cells(r + 4, LastCol)).NumberFormat = EuroNumberFormat()

    ws.Range( _
        ws.Cells(r + 3, c + 1), _
        ws.Cells(r + 3, LastCol)).NumberFormat = "0.00%"

    ws.Range( _
        ws.Cells(r + 5, c + 1), _
        ws.Cells(r + 5, LastCol)).NumberFormat = "0.00%"

    ws.Range( _
        ws.Cells(r + 6, c + 1), _
        ws.Cells(r + 6, LastCol)).NumberFormat = "0.00%"

    FormatReportTable _
        ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 6, LastCol)), _
        1

    FormatFirstColumn ws, r + 1, r + 6, c

    AddBottomBorder ws, r + 3, c, LastCol
    AddBottomBorder ws, r + 5, c, LastCol

    ws.Cells(r + 3, c).Font.Bold = False
    ws.Cells(r + 5, c).Font.Bold = False

    With ws.Range(ws.Cells(r + 6, c), ws.Cells(r + 6, LastCol))
        .Font.Bold = True
        .Interior.Color = RGB(212, 212, 212)
    End With

    CreateCollateralPieChart ws, DictCurrent, ReportDate

End Sub

Private Sub BuildNewLoansSection( _
    ByVal ws As Worksheet, _
    ByRef CurrentAccounts As Variant, _
    ByRef PreviousAccounts As Variant, _
    ByRef CurrentPositions As Variant)

    WriteLoanMovementSection _
        ws, _
        Layout.NewLoanRow, _
        Layout.NewLoanCol, _
        "New Lombard Loans in the Past Month", _
        "New Loans", _
        CurrentAccounts, _
        PreviousAccounts, _
        CurrentPositions

End Sub

Private Sub BuildEndedLoansSection( _
    ByVal ws As Worksheet, _
    ByRef CurrentAccounts As Variant, _
    ByRef ComparisonAccounts As Variant, _
    ByRef ComparisonPositions As Variant)

    WriteLoanMovementSection _
        ws, _
        Layout.EndedLoanRow, _
        Layout.EndedLoanCol, _
        "Lombard Loans Ended in the Past Month", _
        "Ended Loans", _
        ComparisonAccounts, _
        CurrentAccounts, _
        ComparisonPositions

End Sub

'
' One loan-movement table: the accounts present in SubjectAccounts but absent
' from ReferenceAccounts, with their collateral read from the snapshot the
' loans were still live in. New loans compare the current snapshot against the
' previous one; ended loans compare the previous snapshot against the current
' one. An NDG is counted once however many account rows it holds.
'
Private Sub WriteLoanMovementSection( _
    ByVal ws As Worksheet, _
    ByVal TopRow As Long, _
    ByVal LeftCol As Long, _
    ByVal Title As String, _
    ByVal CountHeader As String, _
    ByRef SubjectAccounts As Variant, _
    ByRef ReferenceAccounts As Variant, _
    ByRef SubjectPositions As Variant)

    Dim ReferenceNDGs As Object
    Dim MovedNDGs As Object

    Dim NDG As String

    Dim LoanCount As Long
    Dim TotalApproved As Double
    Dim TotalDrawn As Double
    Dim TotalCollateral As Double

    Dim r As Long

    If Not WeeklyDataHasRows(SubjectAccounts) Then Exit Sub
    If Not WeeklyDataHasRows(ReferenceAccounts) Then Exit Sub

    Set ReferenceNDGs = GetAccountNDGDictionary(ReferenceAccounts)
    Set MovedNDGs = NewNDGSet()

    For r = LBound(SubjectAccounts, 1) To UBound(SubjectAccounts, 1)

        NDG = CleanWeeklyCsvField(SubjectAccounts(r, WeeklyAccountNDG))

        If NDG <> "" Then

            If Not ReferenceNDGs.Exists(NDG) And _
               Not MovedNDGs.Exists(NDG) Then

                MovedNDGs(NDG) = True
                LoanCount = LoanCount + 1

                TotalApproved = TotalApproved + _
                    CDbl(SubjectAccounts(r, WeeklyAccountApproved))

                TotalDrawn = TotalDrawn + _
                    CDbl(SubjectAccounts(r, WeeklyAccountDrawn))

            End If

        End If

    Next r

    If WeeklyDataHasRows(SubjectPositions) Then

        For r = LBound(SubjectPositions, 1) To UBound(SubjectPositions, 1)

            NDG = CleanWeeklyCsvField(SubjectPositions(r, WeeklyPosNDG))

            If MovedNDGs.Exists(NDG) Then
                TotalCollateral = TotalCollateral + _
                    CDbl(SubjectPositions(r, WeeklyPosPositionValue))
            End If

        Next r

    End If

    WriteSectionTitle ws, TopRow, LeftCol, 4, Title

    ws.Cells(TopRow + 1, LeftCol).Value = CountHeader
    ws.Cells(TopRow + 1, LeftCol + 1).Value = "Max Approved Loan"
    ws.Cells(TopRow + 1, LeftCol + 2).Value = "Drawn Amount"
    ws.Cells(TopRow + 1, LeftCol + 3).Value = "Collateral Value"

    ws.Cells(TopRow + 2, LeftCol).Value = LoanCount
    ws.Cells(TopRow + 2, LeftCol + 1).Value = TotalApproved
    ws.Cells(TopRow + 2, LeftCol + 2).Value = TotalDrawn
    ws.Cells(TopRow + 2, LeftCol + 3).Value = TotalCollateral

    ws.Range( _
        ws.Cells(TopRow + 2, LeftCol + 1), _
        ws.Cells(TopRow + 3, LeftCol + 3)).NumberFormat = EuroNumberFormat()

    FormatReportTable _
        ws.Range( _
            ws.Cells(TopRow + 1, LeftCol), _
            ws.Cells(TopRow + 2, LeftCol + 3)), _
        1

End Sub

Private Function BuildCollateralDictionary( _
    ByRef PositionData As Variant, _
    ByRef UnknownAssets As Object) As Object

    Dim Dict As Object

    Dim AssetType As String
    Dim AssetClass As String

    Dim r As Long

    If Not WeeklyDataHasRows(PositionData) Then Exit Function

    Set Dict = NewCollateralDictionary()

    For r = LBound(PositionData, 1) To UBound(PositionData, 1)

        AssetType = CleanWeeklyCsvField(PositionData(r, WeeklyPosAssetType))
        AssetClass = GetAssetClass(AssetType)

        If AssetType <> "" Then
            If Not AssetTypeMapping.Exists(AssetType) Then
                AssetTypeMapping.Add AssetType, AssetClass
            End If
        End If

        If AssetClass = "UNKNOWN" Then
            RegisterUnknownAsset AssetType, UnknownAssets
        Else
            Dict(AssetClass) = Dict(AssetClass) + _
                CDbl(PositionData(r, WeeklyPosPositionValue))
        End If

    Next r

    Set BuildCollateralDictionary = Dict

End Function

'
' Left Public although only GenerateWeeklyAnalysis calls it.  It takes
' no arguments, which is the one shape a worksheet button can be bound
' to, and a binding in the workbook would not be visible from here.
'
Public Sub WriteAssetTypeMapping()

    Dim ws As Worksheet

    Dim Keys() As String

    Dim i As Long
    Dim j As Long

    Dim Temp As String

    Dim r As Long

    If AssetTypeMapping Is Nothing Then Exit Sub
    If AssetTypeMapping.Count = 0 Then Exit Sub

    Set ws = _
        CreateOrReplaceSheet("Asset Type Mapping")

    ws.Cells(1, 1).Value = "Asset Type"
    ws.Cells(1, 2).Value = "Category"

    ws.Rows(1).Font.Bold = True

    '
    ' Copy keys to array
    '

    ReDim Keys(0 To AssetTypeMapping.Count - 1)

    i = 0

    Dim k As Variant

    For Each k In AssetTypeMapping.Keys

        Keys(i) = CStr(k)

        i = i + 1

    Next k

    '
    ' Sort ascending
    '

    For i = LBound(Keys) To UBound(Keys) - 1

        For j = i + 1 To UBound(Keys)

            If StrComp(Keys(i), Keys(j), vbTextCompare) > 0 Then

                Temp = Keys(i)
                Keys(i) = Keys(j)
                Keys(j) = Temp

            End If

        Next j

    Next i

    '
    ' Output
    '

    r = 2

    For i = LBound(Keys) To UBound(Keys)

        ws.Cells(r, 1).Value = Keys(i)

        ws.Cells(r, 2).Value = _
            AssetTypeMapping(Keys(i))

        r = r + 1

    Next i

    ws.Columns("A:B").AutoFit

End Sub

Private Sub BuildEnteredCollateralSection( _
    ByVal ws As Worksheet, _
    ByRef CurrentAccounts As Variant, _
    ByRef PreviousAccounts As Variant, _
    ByRef CurrentPositions As Variant, _
    ByRef UnknownAssets As Object)

    Dim PreviousNDGs As Object
    Dim NewNDGs As Object
    Dim Amounts As Object

    Dim NDG As String
    Dim AssetType As String
    Dim AssetClass As String

    Dim r As Long

    If Not WeeklyDataHasRows(CurrentAccounts) Then Exit Sub
    If Not WeeklyDataHasRows(PreviousAccounts) Then Exit Sub
    If Not WeeklyDataHasRows(CurrentPositions) Then Exit Sub

    Set PreviousNDGs = GetAccountNDGDictionary(PreviousAccounts)
    Set NewNDGs = NewNDGSet()
    Set Amounts = NewCollateralDictionary()

    For r = LBound(CurrentAccounts, 1) To UBound(CurrentAccounts, 1)

        NDG = CleanWeeklyCsvField(CurrentAccounts(r, WeeklyAccountNDG))

        If NDG <> "" Then
            If Not PreviousNDGs.Exists(NDG) Then NewNDGs(NDG) = True
        End If

    Next r

    For r = LBound(CurrentPositions, 1) To UBound(CurrentPositions, 1)

        NDG = CleanWeeklyCsvField(CurrentPositions(r, WeeklyPosNDG))

        If NewNDGs.Exists(NDG) Then

            AssetType = _
                CleanWeeklyCsvField(CurrentPositions(r, WeeklyPosAssetType))
            AssetClass = GetAssetClass(AssetType)

            If AssetClass = "UNKNOWN" Then
                RegisterUnknownAsset AssetType, UnknownAssets
            Else
                Amounts(AssetClass) = Amounts(AssetClass) + _
                    CDbl(CurrentPositions(r, WeeklyPosPositionValue))
            End If

        End If

    Next r

    WriteEnteredCollateralLayout _
        ws, Amounts, Layout.EnteredRow, Layout.EnteredCol

End Sub

Private Sub WriteEnteredCollateralLayout( _
    ByVal ws As Worksheet, _
    ByVal Amounts As Object, _
    ByVal TopRow As Long, _
    ByVal LeftCol As Long)

    Dim Total As Double
    Dim LastCol As Long

    Total = CollateralTotal(Amounts)
    LastCol = LeftCol + CollateralCategoryCount()

    WriteSectionTitle _
        ws, TopRow, LeftCol, _
        CollateralCategoryCount() + 1, _
        "Collateral Entered with New NDGs in the Past Month"

    WriteCollateralHeaders ws, TopRow + 1, LeftCol

    ws.Cells(TopRow + 2, LeftCol).Value = "Collateral Value"
    WriteCollateralAmounts ws, TopRow + 2, LeftCol, Amounts

    ws.Cells(TopRow + 3, LeftCol).Value = "%"
    WriteCollateralShares ws, TopRow + 3, LeftCol, Amounts, Total

    ws.Range( _
        ws.Cells(TopRow + 2, LeftCol + 1), _
        ws.Cells(TopRow + 2, LastCol)).NumberFormat = EuroNumberFormat()

    ws.Range( _
        ws.Cells(TopRow + 3, LeftCol + 1), _
        ws.Cells(TopRow + 3, LastCol)).NumberFormat = "0.00%"

    FormatReportTable _
        ws.Range( _
            ws.Cells(TopRow + 1, LeftCol), _
            ws.Cells(TopRow + 3, LastCol)), _
        1

    FormatFirstColumn ws, TopRow + 1, TopRow + 3, LeftCol

    ws.Cells(TopRow + 3, LeftCol).Font.Bold = False

End Sub

' Top 10 analysis only.
' Do not use this function in the core collateral classification,
' collateral breakdown, pie chart or entered-collateral calculations.
Private Function ResolveTopTenAssetClass( _
    ByVal BaseAssetClass As String, _
    ByVal AdditionalComment As String) As String

    Dim NormalizedComment As String

    If StrComp(Trim(BaseAssetClass), "GP", vbTextCompare) <> 0 Then

        ResolveTopTenAssetClass = BaseAssetClass

        Exit Function

    End If

    NormalizedComment = UCase(Trim(AdditionalComment))
    NormalizedComment = Replace(NormalizedComment, " ", "")
    NormalizedComment = Replace(NormalizedComment, "-", "")
    NormalizedComment = Replace(NormalizedComment, "_", "")

    If Left(NormalizedComment, 6) = "EQUITY" Then

        ResolveTopTenAssetClass = "Equity"

    ElseIf Left(NormalizedComment, 11) = "FIXEDINCOME" Then

        ResolveTopTenAssetClass = "Bonds"

    ElseIf Left(NormalizedComment, 4) = "FUND" Then

        ResolveTopTenAssetClass = "Funds"

    Else

        ResolveTopTenAssetClass = "GP"

    End If

End Function

Private Function IsRiskRelevantCertificateAssetType( _
    ByVal RawAssetType As String) As Boolean

    ' This authoritative risk-analysis filter is intentionally based on the
    ' original Sophis Asset Type, not on GetAssetClass. It controls Top 10
    ' exposure, category totals, percentage denominators, Geography Lookup
    ' and any future geography / sector analysis based on those outputs.
    IsRiskRelevantCertificateAssetType = _
        (NormalizeHeader(RawAssetType) = "CERTIFICATESOTHER")

End Function

Private Function IsUniCreditBondIssuer( _
    ByVal IssuerName As String) As Boolean

    Dim NormalizedIssuer As String

    NormalizedIssuer = UCase(Trim(IssuerName))
    NormalizedIssuer = Replace(NormalizedIssuer, " ", "")
    NormalizedIssuer = Replace(NormalizedIssuer, ".", "")
    NormalizedIssuer = Replace(NormalizedIssuer, "[", "")
    NormalizedIssuer = Replace(NormalizedIssuer, "]", "")

    IsUniCreditBondIssuer = _
        (InStr(1, NormalizedIssuer, "UNICREDIT", vbTextCompare) > 0 Or _
         Left(NormalizedIssuer, 5) = "UCGIM")

End Function

Private Function FindPositionColumn( _
    ByVal ws As Worksheet, _
    ByVal CandidateHeaders As Variant, _
    ByVal FallbackColumn As Long) As Long

    Dim HeaderRow As Long
    Dim LastCol As Long
    Dim ColNo As Long

    Dim Candidate As Variant
    Dim CellHeader As String

    If ws Is Nothing Then

        FindPositionColumn = FallbackColumn

        Exit Function

    End If

    For HeaderRow = 1 To 2

        LastCol = _
            ws.Cells( _
                HeaderRow, _
                ws.Columns.Count).End(xlToLeft).Column

        For ColNo = 1 To LastCol

            If Not IsError(ws.Cells(HeaderRow, ColNo).Value) Then

                CellHeader = _
                    NormalizeHeader( _
                        CStr(ws.Cells(HeaderRow, ColNo).Value))

                For Each Candidate In CandidateHeaders

                    If CellHeader = _
                        NormalizeHeader(CStr(Candidate)) Then

                        FindPositionColumn = ColNo

                        Exit Function

                    End If

                Next Candidate

            End If

        Next ColNo

    Next HeaderRow

    FindPositionColumn = FallbackColumn

End Function

Private Function NormalizeHeader( _
    ByVal HeaderText As String) As String

    Dim Result As String

    Result = UCase(Trim(HeaderText))
    Result = Replace(Result, " ", "")
    Result = Replace(Result, "_", "")
    Result = Replace(Result, "-", "")
    Result = Replace(Result, "/", "")
    Result = Replace(Result, "\", "")
    Result = Replace(Result, ".", "")
    Result = Replace(Result, "(", "")
    Result = Replace(Result, ")", "")
    Result = Replace(Result, "%", "")
    Result = Replace(Result, "&", "AND")

    NormalizeHeader = Result

End Function

Private Function GetOptionalWorksheet( _
    ByVal SheetName As String) As Worksheet

    On Error Resume Next

    Set GetOptionalWorksheet = _
        ThisWorkbook.Worksheets(SheetName)

    On Error GoTo 0

End Function

Private Function GetWorksheetFromWorkbook( _
    ByVal wb As Workbook, _
    ByVal SheetName As String) As Worksheet

    If wb Is Nothing Then Exit Function

    On Error Resume Next

    Set GetWorksheetFromWorkbook = _
        wb.Worksheets(SheetName)

    On Error GoTo 0

End Function

Private Function GetReferenceDataWorksheet( _
    ByVal SheetName As String, _
    ByVal PreferredWorkbook As Workbook) As Worksheet

    Dim wb As Workbook

    Set GetReferenceDataWorksheet = _
        GetWorksheetFromWorkbook( _
            PreferredWorkbook, _
            SheetName)

    If Not GetReferenceDataWorksheet Is Nothing Then Exit Function

    If Not PreferredWorkbook Is ThisWorkbook Then

        Set GetReferenceDataWorksheet = _
            GetWorksheetFromWorkbook( _
                ThisWorkbook, _
                SheetName)

        If Not GetReferenceDataWorksheet Is Nothing Then Exit Function

    End If

    ' Also accept an already-open dedicated reference workbook.
    For Each wb In Application.Workbooks

        Set GetReferenceDataWorksheet = _
            GetWorksheetFromWorkbook( _
                wb, _
                SheetName)

        If Not GetReferenceDataWorksheet Is Nothing Then Exit Function

    Next wb

End Function

Private Function GetReferenceDataTable( _
    ByVal SheetName As String, _
    ByVal TableName As String, _
    ByVal PreferredWorkbook As Workbook) As ListObject

    Dim wsMap As Worksheet

    Set wsMap = _
        GetReferenceDataWorksheet( _
            SheetName, _
            PreferredWorkbook)

    If wsMap Is Nothing Then Exit Function

    On Error Resume Next

    Set GetReferenceDataTable = _
        wsMap.ListObjects(TableName)

    On Error GoTo 0

End Function

Private Function GetTableColumnIndex( _
    ByVal DataTable As ListObject, _
    ByVal HeaderName As String) As Long

    Dim TableColumn As ListColumn

    If DataTable Is Nothing Then Exit Function

    On Error Resume Next

    GetTableColumnIndex = _
        DataTable.ListColumns(HeaderName).Index

    On Error GoTo 0

    If GetTableColumnIndex > 0 Then Exit Function

    For Each TableColumn In DataTable.ListColumns

        If NormalizeHeader(TableColumn.name) = _
           NormalizeHeader(HeaderName) Then

            GetTableColumnIndex = TableColumn.Index

            Exit Function

        End If

    Next TableColumn

End Function

Private Function BuildTableKeyIndex( _
    ByVal DataTable As ListObject, _
    ByVal KeyHeader As String) As Object

    Dim RowIndex As Object
    Dim KeyColumn As Long
    Dim KeyData As Variant
    Dim KeyRange As Range
    Dim KeyWorksheet As Worksheet
    Dim r As Long
    Dim LookupKey As String

    Set RowIndex = NewExactNameMap()

    If DataTable Is Nothing Then

        Set BuildTableKeyIndex = RowIndex

        Exit Function

    End If

    KeyColumn = GetTableColumnIndex(DataTable, KeyHeader)

    If KeyColumn = 0 Or _
       DataTable.DataBodyRange Is Nothing Then

        Set BuildTableKeyIndex = RowIndex

        Exit Function

    End If

    Set KeyRange = DataTable.ListColumns(KeyColumn).DataBodyRange
    Set KeyWorksheet = KeyRange.Parent
    KeyData = _
        ReadWorksheetColumnValues( _
            KeyWorksheet, _
            KeyRange.Row, _
            KeyRange.Row + KeyRange.Rows.Count - 1, _
            KeyRange.Column)

    For r = 1 To UBound(KeyData, 1)

        LookupKey = _
            NormalizeExactNameKey( _
                SafeText( _
                    KeyData(r, 1)))

        If LookupKey <> "" And _
           Not RowIndex.Exists(LookupKey) Then

            RowIndex.Add LookupKey, r

        End If

    Next r

    Set BuildTableKeyIndex = RowIndex

End Function

Private Function BondIssuerTicker( _
    ByVal SecurityName As String) As String

    Dim FirstSpace As Long

    SecurityName = Trim(SecurityName)

    If SecurityName = "" Then Exit Function

    FirstSpace = InStr(1, SecurityName, " ")

    If FirstSpace > 1 Then

        BondIssuerTicker = _
            UCase(Left(SecurityName, FirstSpace - 1))

    Else

        BondIssuerTicker = UCase(SecurityName)

    End If

End Function

Private Function ClassifyBondIssuerType( _
    ByVal RawAssetType As String) As String

    Dim NormalizedType As String

    NormalizedType = UCase(Trim(RawAssetType))

    If Left(NormalizedType, 16) = "SENIOR CORPORATE" Or _
       Left(NormalizedType, 22) = "SUBORDINATED CORPORATE" Then

        ClassifyBondIssuerType = _
            BOND_ISSUER_TYPE_CORPORATE

    ElseIf Left(NormalizedType, 9) = "SOVEREIGN" Then

        ClassifyBondIssuerType = _
            BOND_ISSUER_TYPE_SOVEREIGN

    End If

End Function

Private Function BondRiskAssetClass( _
    ByVal IssuerType As String) As String

    Dim NormalizedType As String

    NormalizedType = UCase(Trim(IssuerType))

    If InStr( _
            1, _
            NormalizedType, _
            "CORPORATE", _
            vbTextCompare) > 0 Then

        BondRiskAssetClass = CORPORATE_BONDS_CLASS

    ElseIf InStr( _
                1, _
                NormalizedType, _
                "SOVEREIGN", _
                vbTextCompare) > 0 Then

        BondRiskAssetClass = SOVEREIGN_BONDS_CLASS

    End If

End Function

Private Function NewBondCandidate( _
    ByVal IssuerTicker As String) As Object

    Dim Candidate As Object
    Dim IsinNames As Object

    Set Candidate = CreateObject("Scripting.Dictionary")
    Candidate.CompareMode = vbTextCompare

    Set IsinNames = NewExactNameMap()

    Candidate.Add "IssuerTicker", IssuerTicker
    Candidate.Add "FirstISIN", ""
    Candidate.Add "FirstBondName", ""
    Candidate.Add "IssuerName", ""
    Candidate.Add "IssuerType", ""
    Candidate.Add "ISINNames", IsinNames

    Set NewBondCandidate = Candidate

End Function

Private Sub AddBondDatabaseCandidate( _
    ByVal Candidates As Object, _
    ByVal IssuerTicker As String, _
    ByVal ISIN As String, _
    ByVal BondName As String, _
    ByVal IssuerName As String, _
    ByVal IssuerType As String)

    Dim Candidate As Object
    Dim IsinNames As Object
    Dim IsinKey As String

    If Candidates Is Nothing Then Exit Sub

    IssuerTicker = BondIssuerTicker(IssuerTicker)

    If IssuerTicker = "" Then Exit Sub

    If Candidates.Exists(IssuerTicker) Then

        Set Candidate = Candidates(IssuerTicker)

    Else

        Set Candidate = NewBondCandidate(IssuerTicker)
        Candidates.Add IssuerTicker, Candidate

    End If

    ISIN = UCase(Trim(ISIN))
    BondName = Trim(BondName)
    IssuerName = Trim(IssuerName)
    IssuerType = Trim(IssuerType)

    If CStr(Candidate("FirstISIN")) = "" And _
       ISIN <> "" Then

        Candidate("FirstISIN") = ISIN

    End If

    If CStr(Candidate("FirstBondName")) = "" And _
       BondName <> "" Then

        Candidate("FirstBondName") = BondName

    End If

    If CStr(Candidate("IssuerName")) = "" And _
       IssuerName <> "" Then

        Candidate("IssuerName") = IssuerName

    End If

    If CStr(Candidate("IssuerType")) = "" And _
       IssuerType <> "" Then

        Candidate("IssuerType") = IssuerType

    End If

    If ISIN <> "" And BondName <> "" Then

        Set IsinNames = Candidate("ISINNames")
        IsinKey = NormalizeExactNameKey(ISIN)

        If Not IsinNames.Exists(IsinKey) Then

            IsinNames.Add IsinKey, BondName

        End If

    End If

End Sub

Private Sub SetFundLookupFormula( _
    ByVal TargetCell As Range, _
    ByVal LookupCell As Range, _
    ByVal FormulaKind As String)

    Dim InvariantFormula As String
    Dim LocalFormula As String
    Dim LookupAddress As String

    If TargetCell Is Nothing Then Exit Sub
    If LookupCell Is Nothing Then Exit Sub

    LookupAddress = LookupCell.Address(False, False)

    If StrComp(FormulaKind, "Prefix", vbTextCompare) = 0 Then

        InvariantFormula = _
            "=XLOOKUP(1,(" & FUND_PARENT_COMPANIES_TABLE & _
            "[Prefix]<>"""")*(LEFT(" & LookupAddress & ",LEN(" & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix]))=" & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix])," & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix],0)"

        LocalFormula = _
            "=XLOOKUP(1;(" & FUND_PARENT_COMPANIES_TABLE & _
            "[Prefix]<>"""")*(LEFT(" & LookupAddress & ";LEN(" & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix]))=" & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix]);" & _
            FUND_PARENT_COMPANIES_TABLE & "[Prefix];0)"

    Else

        InvariantFormula = _
            "=XLOOKUP(" & LookupAddress & "," & _
            FUND_PARENT_COMPANIES_TABLE & _
            "[Prefix]," & FUND_PARENT_COMPANIES_TABLE & _
            "[Company Name],0)"

        LocalFormula = _
            "=XLOOKUP(" & LookupAddress & ";" & _
            FUND_PARENT_COMPANIES_TABLE & _
            "[Prefix];" & FUND_PARENT_COMPANIES_TABLE & _
            "[Company Name];0)"

    End If

    On Error Resume Next

    TargetCell.Formula2 = InvariantFormula

    If Err.Number <> 0 Then

        Err.Clear
        TargetCell.FormulaLocal = LocalFormula

    End If

    On Error GoTo 0

End Sub

Private Function FundLookupFormulaNeedsRefresh( _
    ByVal TargetCell As Range) As Boolean

    Dim CurrentValue As String

    If TargetCell Is Nothing Then Exit Function

    If TargetCell.HasFormula Then

        FundLookupFormulaNeedsRefresh = True

        Exit Function

    End If

    CurrentValue = SafeText(TargetCell.Value)

    FundLookupFormulaNeedsRefresh = _
        (CurrentValue = "" Or CurrentValue = "0")

End Function

Private Function BuildRiskPositionCache( _
    ByRef PositionData As Variant) As Variant

    Dim CachedData() As Variant
    Dim SourceRowIndex As Long
    Dim AssetType As String
    Dim AdditionalComment As String
    Dim BaseAssetClass As String

    If Not WeeklyDataHasRows(PositionData) Then Exit Function

    ' LoadWeeklyPositionData has already cleaned every text field and
    ' converted Position Value before this risk section is called.
    ReDim CachedData( _
        LBound(PositionData, 1) To UBound(PositionData, 1), _
        1 To RISK_POSITION_CACHE_FIELD_COUNT)

    For SourceRowIndex = _
        LBound(PositionData, 1) To _
        UBound(PositionData, 1)

        AssetType = _
            CStr( _
                PositionData( _
                    SourceRowIndex, _
                    WeeklyPosAssetType))
        AdditionalComment = _
            CStr( _
                PositionData( _
                    SourceRowIndex, _
                    WeeklyPosAdditionalComment))
        BaseAssetClass = GetAssetClass(AssetType)

        CachedData( _
            SourceRowIndex, _
            RiskPositionCacheReportingAssetClass) = _
            ResolveTopTenAssetClass( _
                BaseAssetClass, _
                AdditionalComment)
        CachedData( _
            SourceRowIndex, _
            RiskPositionCacheIsDPM) = _
            (StrComp( _
                Trim(BaseAssetClass), _
                "GP", _
                vbTextCompare) = 0)

    Next SourceRowIndex

    BuildRiskPositionCache = CachedData

End Function

Private Sub UpdateRiskReferenceDatabases( _
    ByRef PositionData As Variant, _
    ByRef RiskPositionData As Variant, _
    ByVal PreferredWorkbook As Workbook)

    Dim EquityCandidates As Object
    Dim BondCandidates As Object
    Dim FundCandidates As Object
    Dim Candidate As Object
    Dim IsinNames As Object
    Dim ExistingRows As Object
    Dim DataTable As ListObject
    Dim DataRow As ListRow
    Dim Item As Variant
    Dim LookupKey As String
    Dim ReferenceISIN As String
    Dim ReferenceName As String
    Dim CurrentValue As String

    Dim EquityIsinCol As Long
    Dim EquityNameCol As Long
    Dim BondTickerCol As Long
    Dim BondIsinCol As Long
    Dim BondNameCol As Long
    Dim BondIssuerCol As Long
    Dim BondTypeCol As Long
    Dim FundNameCol As Long
    Dim FundIsinCol As Long
    Dim FundPrefixCol As Long
    Dim FundCompanyCol As Long

    Dim r As Long
    Dim AssetType As String
    Dim ReportingAssetClass As String
    Dim SecurityName As String
    Dim IssuerName As String
    Dim ISIN As String
    Dim IssuerTicker As String
    Dim IssuerType As String
    Set EquityCandidates = NewExactNameMap()
    Set BondCandidates = NewExactNameMap()
    Set FundCandidates = NewExactNameMap()

    For r = _
        LBound(RiskPositionData, 1) To _
        UBound(RiskPositionData, 1)

        ReportingAssetClass = _
            CStr( _
                RiskPositionData( _
                    r, _
                    RiskPositionCacheReportingAssetClass))

        Select Case ReportingAssetClass

            Case "Equity"

                SecurityName = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosSecurityName))
                ISIN = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosISIN))

                If Trim(SecurityName) = "" And _
                   Trim(ISIN) <> "" Then

                    LookupKey = NormalizeExactNameKey(ISIN)

                    If Not EquityCandidates.Exists(LookupKey) Then

                        EquityCandidates.Add LookupKey, UCase(Trim(ISIN))

                    End If

                End If

            Case "Bonds"

                AssetType = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosAssetType))
                SecurityName = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosSecurityName))
                IssuerName = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosIssuer))
                ISIN = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosISIN))

                IssuerTicker = BondIssuerTicker(SecurityName)

                If IssuerTicker <> "" Then

                    IssuerType = ClassifyBondIssuerType(AssetType)

                    AddBondDatabaseCandidate _
                        BondCandidates, _
                        IssuerTicker, _
                        ISIN, _
                        SecurityName, _
                        IssuerName, _
                        IssuerType

                End If

            Case "Funds"

                SecurityName = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosSecurityName))
                ISIN = _
                    CStr( _
                        PositionData( _
                            r, _
                            WeeklyPosISIN))

                If Trim(SecurityName) <> "" Then

                    LookupKey = _
                        NormalizeExactNameKey(SecurityName)

                    If Not FundCandidates.Exists(LookupKey) Then

                        Set Candidate = _
                            CreateObject("Scripting.Dictionary")

                        Candidate.CompareMode = vbTextCompare
                        Candidate.Add "FundName", Trim(SecurityName)
                        Candidate.Add "ReferenceISIN", UCase(Trim(ISIN))

                        FundCandidates.Add LookupKey, Candidate

                    Else

                        Set Candidate = FundCandidates(LookupKey)

                        If CStr(Candidate("ReferenceISIN")) = "" And _
                           Trim(ISIN) <> "" Then

                            Candidate("ReferenceISIN") = _
                                UCase(Trim(ISIN))

                        End If

                    End If

                End If

        End Select

    Next r


    ' Equity Names: append only missing-name ISINs. Existing rows are kept.
    Set DataTable = _
        GetReferenceDataTable( _
            EQUITY_NAMES_SHEET, _
            UNMAPPED_EQUITIES_TABLE, _
            PreferredWorkbook)

    If Not DataTable Is Nothing Then

        EquityIsinCol = GetTableColumnIndex(DataTable, "ISIN")
        EquityNameCol = _
            GetTableColumnIndex(DataTable, "Security Name")

        If EquityIsinCol > 0 And EquityNameCol > 0 Then

            Set ExistingRows = _
                BuildTableKeyIndex(DataTable, "ISIN")

            For Each Item In EquityCandidates.Keys

                If Not ExistingRows.Exists(CStr(Item)) Then

                    Set DataRow = DataTable.ListRows.Add

                    DataRow.Range.Cells(1, EquityIsinCol).NumberFormat = "@"
                    DataRow.Range.Cells(1, EquityIsinCol).Value = _
                        CStr(EquityCandidates(Item))

                    ExistingRows.Add CStr(Item), DataRow.Index

                End If

            Next Item

        End If

    End If


    ' Bond Issuers: one maintained row per ticker. Candidate fields are
    ' accumulated independently, so another ISIN under the same ticker may
    ' supply a missing issuer or classification without extra table scans.
    Set DataTable = _
        GetReferenceDataTable( _
            BOND_ISSUERS_SHEET, _
            BOND_ISSUERS_TABLE, _
            PreferredWorkbook)

    If Not DataTable Is Nothing Then

        BondTickerCol = _
            GetTableColumnIndex(DataTable, "Issuer Ticker")
        BondIsinCol = _
            GetTableColumnIndex(DataTable, "Reference ISIN")
        BondNameCol = _
            GetTableColumnIndex(DataTable, "Reference Bond Name")
        BondIssuerCol = _
            GetTableColumnIndex(DataTable, "Issuer Name")
        BondTypeCol = _
            GetTableColumnIndex(DataTable, "Bond Type")

        If BondTickerCol > 0 And BondIsinCol > 0 And _
           BondNameCol > 0 And BondIssuerCol > 0 And _
           BondTypeCol > 0 Then

            Set ExistingRows = _
                BuildTableKeyIndex(DataTable, "Issuer Ticker")

            For Each Item In BondCandidates.Keys

                Set Candidate = BondCandidates(Item)

                If ExistingRows.Exists(CStr(Item)) Then

                    Set DataRow = _
                        DataTable.ListRows( _
                            CLng(ExistingRows(Item)))

                Else

                    Set DataRow = DataTable.ListRows.Add
                    DataRow.Range.Cells(1, BondTickerCol).Value = _
                        CStr(Candidate("IssuerTicker"))
                    ExistingRows.Add CStr(Item), DataRow.Index

                End If

                CurrentValue = _
                    SafeText( _
                        DataRow.Range.Cells(1, BondIsinCol).Value)

                If CurrentValue = "" And _
                   CStr(Candidate("FirstISIN")) <> "" Then

                    DataRow.Range.Cells(1, BondIsinCol).NumberFormat = "@"
                    DataRow.Range.Cells(1, BondIsinCol).Value = _
                        CStr(Candidate("FirstISIN"))
                    CurrentValue = CStr(Candidate("FirstISIN"))

                End If

                If SafeText( _
                        DataRow.Range.Cells(1, BondNameCol).Value) = "" Then

                    ReferenceName = ""
                    Set IsinNames = Candidate("ISINNames")
                    LookupKey = NormalizeExactNameKey(CurrentValue)

                    If LookupKey <> "" And _
                       IsinNames.Exists(LookupKey) Then

                        ReferenceName = CStr(IsinNames(LookupKey))

                    ElseIf CurrentValue = "" Then

                        ReferenceName = _
                            CStr(Candidate("FirstBondName"))

                    End If

                    If ReferenceName <> "" Then

                        DataRow.Range.Cells(1, BondNameCol).Value = _
                            ReferenceName

                    End If

                End If

                ' Sophis may later supply a corrected issuer name; this is
                ' the only populated field intentionally allowed to update.
                If CStr(Candidate("IssuerName")) <> "" Then

                    DataRow.Range.Cells(1, BondIssuerCol).Value = _
                        CStr(Candidate("IssuerName"))

                End If

                If SafeText( _
                        DataRow.Range.Cells(1, BondTypeCol).Value) = "" And _
                   CStr(Candidate("IssuerType")) <> "" Then

                    DataRow.Range.Cells(1, BondTypeCol).Value = _
                        CStr(Candidate("IssuerType"))

                End If

            Next Item

        End If

    End If


    ' Funds: ordinary and segregated positions share one exact-name table.
    Set DataTable = _
        GetReferenceDataTable( _
            FUND_PARENT_COMPANIES_SHEET, _
            FUNDS_TABLE, _
            PreferredWorkbook)

    If Not DataTable Is Nothing Then

        FundNameCol = GetTableColumnIndex(DataTable, "Fund Name")
        FundIsinCol = _
            GetTableColumnIndex(DataTable, "Reference ISIN")
        FundPrefixCol = GetTableColumnIndex(DataTable, "Prefix")
        FundCompanyCol = _
            GetTableColumnIndex(DataTable, "Company Name")

        If FundNameCol > 0 And FundIsinCol > 0 And _
           FundPrefixCol > 0 And FundCompanyCol > 0 Then

            Set ExistingRows = _
                BuildTableKeyIndex(DataTable, "Fund Name")

            For Each Item In FundCandidates.Keys

                Set Candidate = FundCandidates(Item)

                If ExistingRows.Exists(CStr(Item)) Then

                    Set DataRow = _
                        DataTable.ListRows( _
                            CLng(ExistingRows(Item)))

                Else

                    Set DataRow = DataTable.ListRows.Add
                    DataRow.Range.Cells(1, FundNameCol).Value = _
                        CStr(Candidate("FundName"))
                    ExistingRows.Add CStr(Item), DataRow.Index

                End If

                ReferenceISIN = _
                    SafeText( _
                        DataRow.Range.Cells(1, FundIsinCol).Value)

                If ReferenceISIN = "" And _
                   CStr(Candidate("ReferenceISIN")) <> "" Then

                    DataRow.Range.Cells(1, FundIsinCol).NumberFormat = "@"
                    DataRow.Range.Cells(1, FundIsinCol).Value = _
                        CStr(Candidate("ReferenceISIN"))

                End If

                If FundLookupFormulaNeedsRefresh( _
                        DataRow.Range.Cells(1, FundPrefixCol)) Then

                    SetFundLookupFormula _
                        DataRow.Range.Cells(1, FundPrefixCol), _
                        DataRow.Range.Cells(1, FundNameCol), _
                        "Prefix"

                End If

                If FundLookupFormulaNeedsRefresh( _
                        DataRow.Range.Cells(1, FundCompanyCol)) Then

                    SetFundLookupFormula _
                        DataRow.Range.Cells(1, FundCompanyCol), _
                        DataRow.Range.Cells(1, FundPrefixCol), _
                        "Company"

                End If

            Next Item

        End If

    End If


End Sub

Private Function SafeText( _
    ByVal InputValue As Variant) As String

    If IsError(InputValue) Then Exit Function
    If IsNull(InputValue) Then Exit Function
    If IsEmpty(InputValue) Then Exit Function

    SafeText = Trim(CStr(InputValue))

End Function

Private Function BuildTemporaryIssuerName( _
    ByVal ISIN As String, _
    ByVal SecurityName As String, _
    ByVal SourceRow As Long, _
    ByVal PlaceholderMode As IssuerPlaceholderMode) As String

    Dim TemporaryName As String

    If PlaceholderMode = PlaceholderFromSecurityName Then

        TemporaryName = Trim(SecurityName)

    Else

        TemporaryName = Trim(ISIN)

    End If

    If TemporaryName = "" Then

        TemporaryName = Trim(ISIN)

    End If

    If TemporaryName = "" Then

        TemporaryName = Trim(SecurityName)

    End If

    If TemporaryName = "" Then

        TemporaryName = "Missing identifier row " & CStr(SourceRow)

    End If

    BuildTemporaryIssuerName = _
        "[" & TemporaryName & "]"

End Function


Private Function NewNDGSet() As Object

    Dim NDGs As Object

    Set NDGs = CreateObject("Scripting.Dictionary")
    NDGs.CompareMode = vbTextCompare

    Set NewNDGSet = NDGs

End Function


Private Function GetDefinedNameText( _
    ByVal wb As Workbook, _
    ByVal DefinedName As String) As String

    Dim nm As name
    Dim rng As Range
    Dim ShortName As String

    If wb Is Nothing Then Exit Function

    On Error Resume Next

    Set rng = wb.Names(DefinedName).RefersToRange

    On Error GoTo 0

    If Not rng Is Nothing Then

        GetDefinedNameText = SafeText(rng.Cells(1, 1).Value)

        Exit Function

    End If

    For Each nm In wb.Names

        ShortName = nm.name

        If InStrRev(ShortName, "!") > 0 Then

            ShortName = _
                Mid( _
                    ShortName, _
                    InStrRev(ShortName, "!") + 1)

        End If

        ShortName = Replace(ShortName, "'", "")

        If StrComp( _
                ShortName, _
                DefinedName, _
                vbTextCompare) = 0 Then

            On Error Resume Next

            Set rng = nm.RefersToRange

            On Error GoTo 0

            If Not rng Is Nothing Then

                GetDefinedNameText = _
                    SafeText(rng.Cells(1, 1).Value)

                Exit Function

            End If

        End If

    Next nm

End Function

Private Function GetCertificateFolderPath( _
    ByVal PreferredWorkbook As Workbook) As String

    Dim FolderPath As String

    FolderPath = _
        GetDefinedNameText( _
            PreferredWorkbook, _
            CERTIFICATE_PATH_NAME)

    If FolderPath = "" Then

        FolderPath = _
            GetDefinedNameText( _
                ThisWorkbook, _
                CERTIFICATE_PATH_NAME)

    End If

    If FolderPath = "" Then

        If Not ActiveWorkbook Is Nothing Then

            FolderPath = _
                GetDefinedNameText( _
                    ActiveWorkbook, _
                    CERTIFICATE_PATH_NAME)

        End If

    End If

    FolderPath = Trim(Replace(FolderPath, Chr(34), ""))

    If FolderPath <> "" Then

        If InStr(FolderPath, ":") = 0 And _
           Left(FolderPath, 2) <> "\\" And _
           Left(FolderPath, 1) <> "/" Then

            If Not PreferredWorkbook Is Nothing Then

                If PreferredWorkbook.Path <> "" Then

                    FolderPath = _
                        PreferredWorkbook.Path & _
                        Application.PathSeparator & _
                        FolderPath

                End If

            End If

        End If

    End If

    GetCertificateFolderPath = FolderPath

End Function

Private Function CombineFolderAndFile( _
    ByVal FolderPath As String, _
    ByVal FileName As String) As String

    Dim LastCharacter As String

    If FolderPath = "" Then Exit Function

    LastCharacter = Right(FolderPath, 1)

    If LastCharacter <> "\" And LastCharacter <> "/" Then

        FolderPath = FolderPath & Application.PathSeparator

    End If

    CombineFolderAndFile = FolderPath & FileName

End Function

Private Function EmbeddedFileTimestamp( _
    ByVal FileName As String) As Date

    Dim i As Long
    Dim j As Long

    Dim DateToken As String
    Dim TimeToken As String

    Dim YearNo As Long
    Dim MonthNo As Long
    Dim DayNo As Long
    Dim HourNo As Long
    Dim MinuteNo As Long

    Dim CandidateDate As Date

    For i = 1 To Len(FileName) - 7

        DateToken = Mid(FileName, i, 8)

        If DateToken Like "########" Then

            YearNo = CLng(Left(DateToken, 4))
            MonthNo = CLng(Mid(DateToken, 5, 2))
            DayNo = CLng(Right(DateToken, 2))

            On Error Resume Next

            CandidateDate = _
                DateSerial( _
                    YearNo, _
                    MonthNo, _
                    DayNo)

            If Err.Number = 0 And _
               Year(CandidateDate) = YearNo And _
               Month(CandidateDate) = MonthNo And _
               Day(CandidateDate) = DayNo Then

                On Error GoTo 0

                For j = i + 8 To Len(FileName) - 3

                    TimeToken = Mid(FileName, j, 4)

                    If TimeToken Like "####" Then

                        HourNo = CLng(Left(TimeToken, 2))
                        MinuteNo = CLng(Right(TimeToken, 2))

                        If HourNo <= 23 And MinuteNo <= 59 Then

                            CandidateDate = _
                                CandidateDate + _
                                TimeSerial( _
                                    HourNo, _
                                    MinuteNo, _
                                    0)

                        End If

                        Exit For

                    End If

                Next j

                EmbeddedFileTimestamp = CandidateDate

                Exit Function

            End If

            Err.Clear
            On Error GoTo 0

        End If

    Next i

End Function

Private Function FindLatestCertificateFile( _
    ByVal FolderPath As String, _
    ByVal FilePrefix As String) As String

    Dim FileName As String
    Dim FullPath As String

    Dim CandidateStamp As Date
    Dim CandidateModified As Date
    Dim LatestStamp As Date
    Dim LatestModified As Date

    On Error GoTo SearchFailed

    FileName = _
        Dir( _
            CombineFolderAndFile( _
                FolderPath, _
                FilePrefix & "*.xls*"), _
            vbNormal Or vbReadOnly Or vbHidden Or vbSystem)

    Do While FileName <> ""

        FullPath = CombineFolderAndFile(FolderPath, FileName)

        CandidateStamp = EmbeddedFileTimestamp(FileName)
        CandidateModified = FileDateTime(FullPath)

        If CandidateStamp = 0 Then

            CandidateStamp = CandidateModified

        End If

        If FindLatestCertificateFile = "" Or _
           CandidateStamp > LatestStamp Or _
           (CandidateStamp = LatestStamp And _
            CandidateModified > LatestModified) Then

            FindLatestCertificateFile = FullPath
            LatestStamp = CandidateStamp
            LatestModified = CandidateModified

        End If

        FileName = Dir()

    Loop

    Exit Function

SearchFailed:

    FindLatestCertificateFile = ""

End Function

Private Function GetOpenWorkbookByPath( _
    ByVal FilePath As String) As Workbook

    Dim wb As Workbook

    For Each wb In Application.Workbooks

        If StrComp( _
                wb.FullName, _
                FilePath, _
                vbTextCompare) = 0 Then

            Set GetOpenWorkbookByPath = wb

            Exit Function

        End If

    Next wb

End Function

Private Function OpenCertificateReferenceWorkbook( _
    ByVal FilePath As String, _
    ByRef OpenedByCode As Boolean) As Workbook

    Set OpenCertificateReferenceWorkbook = _
        GetOpenWorkbookByPath(FilePath)

    If Not OpenCertificateReferenceWorkbook Is Nothing Then Exit Function

    Set OpenCertificateReferenceWorkbook = _
        Workbooks.Open( _
            FileName:=FilePath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, _
            AddToMru:=False)

    OpenedByCode = True

End Function

Private Function GetReferenceWorksheet( _
    ByVal wb As Workbook, _
    ByVal PreferredName As String) As Worksheet

    If wb Is Nothing Then Exit Function

    On Error Resume Next

    Set GetReferenceWorksheet = _
        wb.Worksheets(PreferredName)

    If GetReferenceWorksheet Is Nothing Then

        Set GetReferenceWorksheet = wb.Worksheets(1)

    End If

    On Error GoTo 0

End Function

Private Function NewCertificateComponent( _
    ByVal ComponentName As String, _
    ByVal ComponentWeight As Double, _
    Optional ByVal ReferenceISIN As String = "", _
    Optional ByVal AssetClass As String = "", _
    Optional ByVal IsTemporary As Boolean = False) As Object

    Dim Component As Object

    Set Component = CreateObject("Scripting.Dictionary")
    Component.CompareMode = vbTextCompare

    Component.Add "Name", Trim(ComponentName)
    Component.Add "Weight", ComponentWeight
    Component.Add "ReferenceISIN", UCase(Trim(ReferenceISIN))
    Component.Add "AssetClass", Trim(AssetClass)
    Component.Add "Temporary", IsTemporary

    Set NewCertificateComponent = Component

End Function

Private Sub AddCertificateComponent( _
    ByVal Components As Collection, _
    ByVal ComponentName As String, _
    ByVal ComponentWeight As Double, _
    Optional ByVal ReferenceISIN As String = "", _
    Optional ByVal AssetClass As String = "", _
    Optional ByVal IsTemporary As Boolean = False)

    Dim Component As Object

    If Components Is Nothing Then Exit Sub
    If Trim(ComponentName) = "" Then Exit Sub
    If ComponentWeight = 0 Then Exit Sub

    Set Component = _
        NewCertificateComponent( _
            ComponentName, _
            ComponentWeight, _
            ReferenceISIN, _
            AssetClass, _
            IsTemporary)

    Components.Add Component

End Sub

Private Sub AppendScaledCertificateComponents( _
    ByVal Target As Collection, _
    ByVal Source As Collection, _
    ByVal ScaleFactor As Double)

    Dim Item As Variant
    Dim Component As Object

    If Target Is Nothing Then Exit Sub
    If Source Is Nothing Then Exit Sub

    For Each Item In Source

        Set Component = Item

        AddCertificateComponent _
            Target, _
            CStr(Component("Name")), _
            CDbl(Component("Weight")) * ScaleFactor, _
            CStr(Component("ReferenceISIN")), _
            CStr(Component("AssetClass")), _
            CBool(Component("Temporary"))

    Next Item

End Sub

Private Function CertificateComponentsContainResolvedEntity( _
    ByVal Components As Collection) As Boolean

    Dim Item As Variant
    Dim Component As Object
    Dim ComponentName As String

    If Components Is Nothing Then Exit Function

    For Each Item In Components

        Set Component = Item
        ComponentName = CStr(Component("Name"))

        If Left( _
                ComponentName, _
                Len(MISSING_CERTIFICATE_RIC_PREFIX)) <> _
           MISSING_CERTIFICATE_RIC_PREFIX And _
           Left( _
                ComponentName, _
                Len(UNKNOWN_UNDERLYING_PREFIX)) <> _
           UNKNOWN_UNDERLYING_PREFIX Then

            CertificateComponentsContainResolvedEntity = True

            Exit Function

        End If

    Next Item

End Function

Private Function RegexReplaceText( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    ByVal Replacement As String, _
    Optional ByVal ReplaceAll As Boolean = True) As String

    Dim Regex As Object

    On Error GoTo ReplaceFailed

    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = Pattern
    Regex.Global = ReplaceAll
    Regex.IgnoreCase = True

    RegexReplaceText = Regex.Replace(InputText, Replacement)

    Exit Function

ReplaceFailed:

    RegexReplaceText = InputText
    Err.Clear

End Function

Private Function ExtractFirstISIN( _
    ByVal InputText As String) As String

    Dim Regex As Object
    Dim Matches As Object

    On Error GoTo ExtractFailed

    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = _
        "(^|[^A-Z0-9])([A-Z]{2}[A-Z0-9]{9}[0-9])([^A-Z0-9]|$)"
    Regex.Global = False
    Regex.IgnoreCase = True

    Set Matches = Regex.Execute(UCase(InputText))

    If Matches.Count > 0 Then

        ExtractFirstISIN = _
            UCase(CStr(Matches(0).SubMatches(1)))

    End If

    Exit Function

ExtractFailed:

    ExtractFirstISIN = ""
    Err.Clear

End Function

Private Function IsBasketUnderlyingName( _
    ByVal UnderlyingName As String) As Boolean

    Dim NormalizedName As String

    NormalizedName = _
        UCase( _
            Trim( _
                Replace( _
                    UnderlyingName, _
                    Chr(160), _
                    " ")))

    If Left(NormalizedName, 6) <> "BASKET" Then Exit Function

    IsBasketUnderlyingName = _
        (Len(NormalizedName) = 6 Or _
         Mid(NormalizedName, 7, 1) = ":" Or _
         Mid(NormalizedName, 7, 1) = " ")

End Function

Private Function RemoveBasketPrefixes( _
    ByVal BasketText As String) As String

    Dim Result As String

    Result = Replace(BasketText, Chr(160), " ")
    Result = Trim(Result)

    Do While IsBasketUnderlyingName(Result)

        Result = Trim(Mid(Result, 7))

        If Left(Result, 1) = ":" Then

            Result = Trim(Mid(Result, 2))

        End If

    Loop

    RemoveBasketPrefixes = Result

End Function

Private Function IsNonEntityBasketComponent( _
    ByVal ComponentText As String) As Boolean

    Dim NormalizedText As String

    NormalizedText = _
        UCase( _
            Trim( _
                Replace( _
                    ComponentText, _
                    Chr(160), _
                    " ")))

    If NormalizedText = "" Then

        IsNonEntityBasketComponent = True

        Exit Function

    End If

    If InStr(1, NormalizedText, "CASH", vbTextCompare) > 0 Or _
       InStr(1, NormalizedText, "CONTINGENT VALUE RIGHT", vbTextCompare) > 0 Or _
       InStr(1, NormalizedText, "WARRANT", vbTextCompare) > 0 Or _
       InStr(1, NormalizedText, "CORPORATE ACTION", vbTextCompare) > 0 Then

        IsNonEntityBasketComponent = True

        Exit Function

    End If

    NormalizedText = _
        RegexReplaceText( _
            NormalizedText, _
            "^[ ]*[+-]?[ ]*[0-9]+([.,][0-9]+)?[ ]*(X|\*)?[ ]*", _
            "", _
            False)

    NormalizedText = Trim(NormalizedText)

    IsNonEntityBasketComponent = _
        (NormalizedText = "RIGHT" Or _
         NormalizedText = "RIGHTS" Or _
         NormalizedText = "CASH COMPONENT")

End Function

Private Function TruncateBasketTextAtPhrase( _
    ByVal InputText As String, _
    ByVal Phrase As String) As String

    Dim Position As Long

    Position = _
        InStr( _
            1, _
            UCase(InputText), _
            UCase(Phrase), _
            vbBinaryCompare)

    If Position > 1 Then

        TruncateBasketTextAtPhrase = _
            Trim(Left(InputText, Position - 1))

    Else

        TruncateBasketTextAtPhrase = InputText

    End If

End Function

Private Function CleanBasketComponentName( _
    ByVal ComponentText As String) As String

    Dim Result As String
    Dim PreviousResult As String
    Dim Phrase As Variant

    Result = RemoveBasketPrefixes(ComponentText)

    Do

        PreviousResult = Result

        Result = _
            RegexReplaceText( _
                Result, _
                "^[ ]*[+-]?[ ]*[0-9]+([.,][0-9]+)?[ ]*(X|\*)[ ]*", _
                "", _
                False)

        Result = _
            RegexReplaceText( _
                Result, _
                "^[ ]*[+-]?[ ]*[0-9]+([.,][0-9]+)?[ ]+SHARES?[ ]+", _
                "", _
                False)

        Result = _
            RegexReplaceText( _
                Result, _
                "^[ ]*[+-]?[ ]*[0-9]+([.,][0-9]+)?[ ]+", _
                "", _
                False)

        Result = _
            RegexReplaceText( _
                Result, _
                "^[ ]*X[ ]+", _
                "", _
                False)

        Result = Trim(Result)

    Loop While Result <> PreviousResult

    Result = _
        RegexReplaceText( _
            Result, _
            "\([ ]*(ISIN[ ]+)?[A-Z]{2}[A-Z0-9]{9}[0-9][ ]*\)", _
            " ")

    Result = _
        RegexReplaceText( _
            Result, _
            "\([ ]*[A-Z0-9._-]{2,15}[ ]*\)", _
            " ")

    Result = _
        RegexReplaceText( _
            Result, _
            "\([^)]*(VERKAUF|UMTAUSCH|TENDER)[^)]*\)", _
            " ")

    For Each Phrase In Array( _
        " ZUM VERKAUF", _
        " Z.VERKAUF", _
        " Z. VERKAUF", _
        " ZUM UMTAUSCH", _
        " Z.UMTAUSCH", _
        " Z. UMTAUSCH", _
        " TENDERED SHARE", _
        " - CLOSEOUT", _
        "_COMPO", _
        " COMPO ")

        Result = _
            TruncateBasketTextAtPhrase( _
                Result, _
                CStr(Phrase))

    Next Phrase

    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)

    Do While InStr(Result, "  ") > 0

        Result = Replace(Result, "  ", " ")

    Loop

    Do While Len(Result) > 0 And _
             InStr("+;/", Right(Result, 1)) > 0

        Result = Trim(Left(Result, Len(Result) - 1))

    Loop

    CleanBasketComponentName = Result

End Function

Private Function TrySplitConcatenatedLegalEntities( _
    ByVal InputText As String, _
    ByRef FirstName As String, _
    ByRef SecondName As String) As Boolean

    Dim Suffix As Variant
    Dim Position As Long
    Dim CandidateFirst As String
    Dim CandidateSecond As String

    For Each Suffix In Array( _
        " S.P.A. ", _
        " S.A. ", _
        " N.V. ", _
        " INC. ", _
        " CORP. ", _
        " LTD. ", _
        " PLC ", _
        " AG ", _
        " SA ", _
        " NV ", _
        " SPA ", _
        " INC ", _
        " CORP ", _
        " LTD ")

        Position = _
            InStr( _
                1, _
                UCase(InputText), _
                CStr(Suffix), _
                vbBinaryCompare)

        If Position > 1 Then

            CandidateFirst = _
                Trim( _
                    Left( _
                        InputText, _
                        Position + Len(CStr(Suffix)) - 2))

            CandidateSecond = _
                Trim( _
                    Mid( _
                        InputText, _
                        Position + Len(CStr(Suffix))))

            If Len(CandidateFirst) >= 4 And _
               Len(CandidateSecond) >= 4 And _
               Left(CandidateSecond, 1) <> "(" Then

                FirstName = CandidateFirst
                SecondName = CandidateSecond
                TrySplitConcatenatedLegalEntities = True

                Exit Function

            End If

        End If

    Next Suffix

End Function

Private Function SplitBasketTextComponents( _
    ByVal BasketText As String) As Collection

    Dim Components As Collection
    Dim RawParts As Variant
    Dim RawPart As Variant
    Dim Result As String
    Dim ComponentName As String
    Dim FirstName As String
    Dim SecondName As String
    Dim Delimiter As String
    Dim DigitNo As Long

    Set Components = New Collection

    Result = RemoveBasketPrefixes(BasketText)
    Delimiter = Chr(30)

    Result = Replace(Result, vbCrLf, Delimiter)
    Result = Replace(Result, vbCr, Delimiter)
    Result = Replace(Result, vbLf, Delimiter)
    Result = Replace(Result, "+", Delimiter)
    Result = Replace(Result, " / ", Delimiter)

    For DigitNo = 0 To 9

        Result = _
            Replace( _
                Result, _
                " - " & CStr(DigitNo), _
                Delimiter & CStr(DigitNo), _
                1, _
                -1, _
                vbTextCompare)

    Next DigitNo

    Result = _
        RegexReplaceText( _
            Result, _
            "\)[ ]+([0-9]+([.,][0-9]+)?[ ]*X[ ]+)", _
            ")" & Delimiter & "$1")

    RawParts = Split(Result, Delimiter)

    For Each RawPart In RawParts

        If Not IsNonEntityBasketComponent(CStr(RawPart)) Then

            ComponentName = _
                CleanBasketComponentName(CStr(RawPart))

            If ComponentName <> "" Then

                Components.Add Trim(CStr(RawPart))

            End If

        End If

    Next RawPart

    If Components.Count = 1 Then

        ComponentName = _
            CleanBasketComponentName(CStr(Components(1)))

        If TrySplitConcatenatedLegalEntities( _
                ComponentName, _
                FirstName, _
                SecondName) Then

            Set Components = New Collection
            AddUniqueText Components, FirstName
            AddUniqueText Components, SecondName

        End If

    End If

    Set SplitBasketTextComponents = Components

End Function

Private Function FindKnownRicInText( _
    ByVal ComponentText As String, _
    ByVal RicToName As Object) As String

    Dim TokenText As String
    Dim Tokens As Variant
    Dim Token As Variant
    Dim Candidate As String

    If RicToName Is Nothing Then Exit Function

    TokenText = Replace(ComponentText, Chr(160), " ")
    TokenText = Replace(TokenText, "(", " ")
    TokenText = Replace(TokenText, ")", " ")
    TokenText = Replace(TokenText, ";", " ")
    TokenText = Replace(TokenText, "+", " ")
    TokenText = Replace(TokenText, "/", " ")
    TokenText = Replace(TokenText, vbTab, " ")

    Tokens = Split(TokenText, " ")

    For Each Token In Tokens

        Candidate = UCase(Trim(CStr(Token)))

        Do While Len(Candidate) > 0 And _
                 InStr(",:[]{}", Right(Candidate, 1)) > 0

            Candidate = Left(Candidate, Len(Candidate) - 1)

        Loop

        If Candidate <> "" Then

            If RicToName.Exists(Candidate) Then

                FindKnownRicInText = Candidate

                Exit Function

            End If

        End If

    Next Token

End Function

Private Function ExpandCertificateRic( _
    ByVal Ric As String, _
    ByVal RicToName As Object, _
    ByVal RicToISIN As Object, _
    ByVal RicToAssetClass As Object, _
    ByVal RicToBasketRics As Object, _
    ByVal IsinToRic As Object, _
    ByVal NameToRic As Object, _
    ByVal RecursionPath As Object, _
    ByVal Depth As Long) As Collection

    Dim Result As Collection
    Dim ChildRics As Collection
    Dim TextComponents As Collection
    Dim ChildResult As Collection

    Dim UnderlyingName As String
    Dim UnderlyingISIN As String
    Dim UnderlyingAssetClass As String
    Dim RawChildRics As String
    Dim ChildRic As String
    Dim ComponentText As String
    Dim CleanedName As String
    Dim ExplicitISIN As String
    Dim LookupKey As String
    Dim PartWeight As Double

    Dim Parts As Variant
    Dim Part As Variant

    Set Result = New Collection

    Ric = UCase(Trim(Ric))

    If Ric = "" Then

        AddCertificateComponent _
            Result, _
            MISSING_CERTIFICATE_RIC_PREFIX & "(blank RIC)", _
            1, _
            "", _
            "", _
            True

        Set ExpandCertificateRic = Result

        Exit Function

    End If

    If Depth > CERTIFICATE_MAX_BASKET_DEPTH Then

        AddCertificateComponent _
            Result, _
            MISSING_CERTIFICATE_RIC_PREFIX & Ric, _
            1, _
            "", _
            "", _
            True

        Set ExpandCertificateRic = Result

        Exit Function

    End If

    If RecursionPath.Exists(Ric) Then

        AddCertificateComponent _
            Result, _
            MISSING_CERTIFICATE_RIC_PREFIX & Ric, _
            1, _
            "", _
            "", _
            True

        Set ExpandCertificateRic = Result

        Exit Function

    End If

    If Not RicToName.Exists(Ric) Then

        AddCertificateComponent _
            Result, _
            MISSING_CERTIFICATE_RIC_PREFIX & Ric, _
            1, _
            "", _
            "", _
            True

        Set ExpandCertificateRic = Result

        Exit Function

    End If

    RecursionPath.Add Ric, True

    UnderlyingName = CStr(RicToName(Ric))

    If RicToISIN.Exists(Ric) Then

        UnderlyingISIN = CStr(RicToISIN(Ric))

    End If

    If RicToAssetClass.Exists(Ric) Then

        UnderlyingAssetClass = CStr(RicToAssetClass(Ric))

    End If

    If Not IsBasketUnderlyingName(UnderlyingName) Then

        AddCertificateComponent _
            Result, _
            UnderlyingName, _
            1, _
            UnderlyingISIN, _
            UnderlyingAssetClass

        GoTo ReturnResult

    End If

    Set ChildRics = New Collection

    If RicToBasketRics.Exists(Ric) Then

        RawChildRics = CStr(RicToBasketRics(Ric))
        RawChildRics = Replace(RawChildRics, vbCrLf, ";")
        RawChildRics = Replace(RawChildRics, vbCr, ";")
        RawChildRics = Replace(RawChildRics, vbLf, ";")
        RawChildRics = Replace(RawChildRics, ",", ";")
        RawChildRics = Replace(RawChildRics, "|", ";")

        Parts = Split(RawChildRics, ";")

        For Each Part In Parts

            AddUniqueText _
                ChildRics, _
                UCase(Trim(CStr(Part)))

        Next Part

    End If

    If ChildRics.Count > 0 Then

        PartWeight = 1 / ChildRics.Count

        For Each Part In ChildRics

            ChildRic = CStr(Part)

            Set ChildResult = _
                ExpandCertificateRic( _
                    ChildRic, _
                    RicToName, _
                    RicToISIN, _
                    RicToAssetClass, _
                    RicToBasketRics, _
                    IsinToRic, _
                    NameToRic, _
                    RecursionPath, _
                    Depth + 1)

            AppendScaledCertificateComponents _
                Result, _
                ChildResult, _
                PartWeight

        Next Part

        If CertificateComponentsContainResolvedEntity(Result) Then

            GoTo ReturnResult

        End If

        ' An obsolete component RIC can disappear from a newer InstrumentList.
        ' In that case the descriptive Basket text is the more useful fallback.
        Set Result = New Collection

    End If

    Set TextComponents = _
        SplitBasketTextComponents(UnderlyingName)

    If TextComponents.Count = 0 Then

        AddCertificateComponent _
            Result, _
            UNKNOWN_UNDERLYING_PREFIX & UnderlyingName, _
            1, _
            "", _
            UnderlyingAssetClass

        GoTo ReturnResult

    End If

    PartWeight = 1 / TextComponents.Count

    For Each Part In TextComponents

        ComponentText = CStr(Part)
        ExplicitISIN = ""
        ChildRic = _
            FindKnownRicInText( _
                ComponentText, _
                RicToName)

        If ChildRic = "" Then

            ExplicitISIN = ExtractFirstISIN(ComponentText)

            If ExplicitISIN <> "" Then

                If IsinToRic.Exists(ExplicitISIN) Then

                    ChildRic = CStr(IsinToRic(ExplicitISIN))

                End If

            End If

        End If

        CleanedName = CleanBasketComponentName(ComponentText)

        If ChildRic = "" And CleanedName <> "" Then

            LookupKey = NormalizeExactNameKey(CleanedName)

            If NameToRic.Exists(LookupKey) Then

                ChildRic = CStr(NameToRic(LookupKey))

            End If

        End If

        If ChildRic <> "" Then

            Set ChildResult = _
                ExpandCertificateRic( _
                    ChildRic, _
                    RicToName, _
                    RicToISIN, _
                    RicToAssetClass, _
                    RicToBasketRics, _
                    IsinToRic, _
                    NameToRic, _
                    RecursionPath, _
                    Depth + 1)

            AppendScaledCertificateComponents _
                Result, _
                ChildResult, _
                PartWeight

        ElseIf CleanedName <> "" Then

            AddCertificateComponent _
                Result, _
                CleanedName, _
                PartWeight, _
                ExplicitISIN, _
                UnderlyingAssetClass

        Else

            AddCertificateComponent _
                Result, _
                MISSING_CERTIFICATE_RIC_PREFIX & ComponentText, _
                PartWeight, _
                "", _
                UnderlyingAssetClass, _
                True

        End If

    Next Part

ReturnResult:

    If RecursionPath.Exists(Ric) Then

        RecursionPath.Remove Ric

    End If

    Set ExpandCertificateRic = Result

End Function

Private Function LoadCertificateUnderlyingMap( _
    ByRef MappingReady As Boolean, _
    ByVal PreferredWorkbook As Workbook, _
    ByRef UnderlyingReferenceIsinMap As Object, _
    ByRef UnderlyingAssetClassMap As Object, _
    ByRef MappingIssue As String) As Object

    Dim Mapping As Object
    Dim IsinToRics As Object
    Dim RicToName As Object
    Dim RicToISIN As Object
    Dim RicToAssetClass As Object
    Dim RicToBasketRics As Object
    Dim IsinToRic As Object
    Dim NameToRic As Object
    Dim ExpandedRicCache As Object

    Dim wbSearch As Workbook
    Dim wbInstrument As Workbook
    Dim wsSearch As Worksheet
    Dim wsInstrument As Worksheet

    Dim SearchOpenedByCode As Boolean
    Dim InstrumentOpenedByCode As Boolean

    Dim FolderPath As String
    Dim SearchPath As String
    Dim InstrumentPath As String

    Dim SearchIsinCol As Long
    Dim SearchRicCol As Long
    Dim InstrumentRicCol As Long
    Dim InstrumentNameCol As Long
    Dim InstrumentIsinCol As Long
    Dim InstrumentAssetClassCol As Long
    Dim InstrumentBasketRicsCol As Long

    Dim LastRow As Long
    Dim r As Long
    Dim InstrumentRicData As Variant
    Dim InstrumentNameData As Variant
    Dim InstrumentIsinData As Variant
    Dim InstrumentAssetClassData As Variant
    Dim InstrumentBasketRicsData As Variant
    Dim SearchIsinData As Variant
    Dim SearchRicData As Variant

    Dim CertificateISIN As String
    Dim RawRics As String
    Dim Ric As String
    Dim UnderlyingName As String
    Dim UnderlyingISIN As String
    Dim UnderlyingAssetClass As String
    Dim RawBasketRics As String
    Dim UnderlyingKey As String

    Dim Parts As Variant
    Dim Part As Variant
    Dim IsinKey As Variant

    Dim Rics As Collection
    Dim UnderlyingComponents As Collection
    Dim RicComponents As Collection
    Dim RecursionPath As Object
    Dim Component As Object

    Dim TopLevelWeight As Double

    Dim MissingRicCount As Long

    Set Mapping = CreateObject("Scripting.Dictionary")
    Mapping.CompareMode = vbTextCompare

    Set IsinToRics = CreateObject("Scripting.Dictionary")
    IsinToRics.CompareMode = vbTextCompare

    Set RicToName = CreateObject("Scripting.Dictionary")
    RicToName.CompareMode = vbTextCompare

    Set RicToISIN = CreateObject("Scripting.Dictionary")
    RicToISIN.CompareMode = vbTextCompare

    Set RicToAssetClass = CreateObject("Scripting.Dictionary")
    RicToAssetClass.CompareMode = vbTextCompare

    Set RicToBasketRics = CreateObject("Scripting.Dictionary")
    RicToBasketRics.CompareMode = vbTextCompare

    Set IsinToRic = CreateObject("Scripting.Dictionary")
    IsinToRic.CompareMode = vbTextCompare

    Set NameToRic = CreateObject("Scripting.Dictionary")
    NameToRic.CompareMode = vbTextCompare

    Set ExpandedRicCache = CreateObject("Scripting.Dictionary")
    ExpandedRicCache.CompareMode = vbTextCompare

    Set UnderlyingReferenceIsinMap = _
        CreateObject("Scripting.Dictionary")
    UnderlyingReferenceIsinMap.CompareMode = vbTextCompare

    Set UnderlyingAssetClassMap = _
        CreateObject("Scripting.Dictionary")
    UnderlyingAssetClassMap.CompareMode = vbTextCompare

    FolderPath = GetCertificateFolderPath(PreferredWorkbook)

    If FolderPath = "" Then

        MappingIssue = _
            "The named cell '" & _
            CERTIFICATE_PATH_NAME & _
            "' is missing or empty."

        GoTo ReturnMapping

    End If

    SearchPath = _
        FindLatestCertificateFile( _
            FolderPath, _
            CERTIFICATE_SEARCH_PREFIX)

    InstrumentPath = _
        FindLatestCertificateFile( _
            FolderPath, _
            CERTIFICATE_INSTRUMENT_PREFIX)

    If SearchPath = "" Then

        MappingIssue = _
            "No SearchResults*.xls* file was found in '" & _
            FolderPath & _
            "'."

        GoTo ReturnMapping

    End If

    If InstrumentPath = "" Then

        MappingIssue = _
            "No InstrumentList*.xls* file was found in '" & _
            FolderPath & _
            "'."

        GoTo ReturnMapping

    End If


    On Error GoTo LoadFailed

    Set wbSearch = _
        OpenCertificateReferenceWorkbook( _
            SearchPath, _
            SearchOpenedByCode)

    Set wbInstrument = _
        OpenCertificateReferenceWorkbook( _
            InstrumentPath, _
            InstrumentOpenedByCode)


    Set wsSearch = _
        GetReferenceWorksheet( _
            wbSearch, _
            "Search Results")

    Set wsInstrument = _
        GetReferenceWorksheet( _
            wbInstrument, _
            "Instruments")

    If wsSearch Is Nothing Or wsInstrument Is Nothing Then

        Err.Raise _
            vbObjectError + 9130, _
            "LoadCertificateUnderlyingMap", _
            "A required worksheet could not be opened."

    End If


    SearchIsinCol = _
        FindPositionColumn( _
            wsSearch, _
            Array("ISIN"), _
            1)

    SearchRicCol = _
        FindPositionColumn( _
            wsSearch, _
            Array("Underlying RIC", "Underlying RICs", "RIC"), _
            6)

    InstrumentRicCol = _
        FindPositionColumn( _
            wsInstrument, _
            Array("RIC"), _
            1)

    InstrumentNameCol = _
        FindPositionColumn( _
            wsInstrument, _
            Array("Underlying Name", "Company Name", "Name"), _
            4)

    InstrumentIsinCol = _
        FindPositionColumn( _
            wsInstrument, _
            Array("ISIN"), _
            7)

    InstrumentAssetClassCol = _
        FindPositionColumn( _
            wsInstrument, _
            Array("Asset Class"), _
            3)

    InstrumentBasketRicsCol = _
        FindPositionColumn( _
            wsInstrument, _
            Array( _
                "Basket Component RICs", _
                "Basket Components", _
                "Component RICs"), _
            51)

    LastRow = _
        wsInstrument.Cells( _
            wsInstrument.Rows.Count, _
            InstrumentRicCol).End(xlUp).Row

    If LastRow >= 2 Then

        InstrumentRicData = _
            ReadWorksheetColumnValues( _
                wsInstrument, 2, LastRow, InstrumentRicCol)
        InstrumentNameData = _
            ReadWorksheetColumnValues( _
                wsInstrument, 2, LastRow, InstrumentNameCol)
        InstrumentIsinData = _
            ReadWorksheetColumnValues( _
                wsInstrument, 2, LastRow, InstrumentIsinCol)
        InstrumentAssetClassData = _
            ReadWorksheetColumnValues( _
                wsInstrument, 2, LastRow, InstrumentAssetClassCol)
        InstrumentBasketRicsData = _
            ReadWorksheetColumnValues( _
                wsInstrument, 2, LastRow, InstrumentBasketRicsCol)

    End If


    For r = 1 To LastRow - 1

        Ric = _
            UCase( _
                SafeText( _
                    InstrumentRicData(r, 1)))

        UnderlyingName = _
            SafeText( _
                InstrumentNameData(r, 1))

        UnderlyingISIN = _
            UCase( _
                SafeText( _
                    InstrumentIsinData(r, 1)))

        UnderlyingAssetClass = _
            SafeText( _
                InstrumentAssetClassData(r, 1))

        RawBasketRics = _
            SafeText( _
                InstrumentBasketRicsData(r, 1))

        If Ric <> "" And UnderlyingName <> "" Then

            If Not RicToName.Exists(Ric) Then

                RicToName.Add Ric, UnderlyingName

            End If

            UnderlyingKey = _
                NormalizeExactNameKey(UnderlyingName)

            If UnderlyingKey <> "" Then

                If Not NameToRic.Exists(UnderlyingKey) Then

                    NameToRic.Add UnderlyingKey, Ric

                End If

            End If

        End If

        If Ric <> "" And UnderlyingISIN <> "" Then

            If Not RicToISIN.Exists(Ric) Then

                RicToISIN.Add Ric, UnderlyingISIN

            End If

            If Not IsinToRic.Exists(UnderlyingISIN) Then

                IsinToRic.Add UnderlyingISIN, Ric

            End If

        End If

        If Ric <> "" And UnderlyingAssetClass <> "" Then

            If Not RicToAssetClass.Exists(Ric) Then

                RicToAssetClass.Add Ric, UnderlyingAssetClass

            End If

        End If

        If Ric <> "" And RawBasketRics <> "" Then

            If Not RicToBasketRics.Exists(Ric) Then

                RicToBasketRics.Add Ric, RawBasketRics

            End If

        End If

        If UnderlyingName <> "" Then

            UnderlyingKey = _
                NormalizeExactNameKey(UnderlyingName)

            If UnderlyingISIN <> "" Then

                If Not UnderlyingReferenceIsinMap.Exists(UnderlyingKey) Then

                    UnderlyingReferenceIsinMap.Add _
                        UnderlyingKey, _
                        UnderlyingISIN

                ElseIf SafeText( _
                        UnderlyingReferenceIsinMap(UnderlyingKey)) = "" Then

                    UnderlyingReferenceIsinMap(UnderlyingKey) = _
                        UnderlyingISIN

                End If

            End If

            If UnderlyingAssetClass <> "" Then

                If Not UnderlyingAssetClassMap.Exists(UnderlyingKey) Then

                    UnderlyingAssetClassMap.Add _
                        UnderlyingKey, _
                        UnderlyingAssetClass

                End If

            End If

        End If

    Next r


    LastRow = _
        wsSearch.Cells( _
            wsSearch.Rows.Count, _
            SearchIsinCol).End(xlUp).Row

    If LastRow >= 2 Then

        SearchIsinData = _
            ReadWorksheetColumnValues( _
                wsSearch, 2, LastRow, SearchIsinCol)
        SearchRicData = _
            ReadWorksheetColumnValues( _
                wsSearch, 2, LastRow, SearchRicCol)

    End If


    For r = 1 To LastRow - 1

        CertificateISIN = _
            UCase( _
                SafeText( _
                    SearchIsinData(r, 1)))

        RawRics = _
            SafeText( _
                SearchRicData(r, 1))

        If CertificateISIN <> "" And RawRics <> "" Then

            If IsinToRics.Exists(CertificateISIN) Then

                Set Rics = IsinToRics(CertificateISIN)

            Else

                Set Rics = New Collection
                IsinToRics.Add CertificateISIN, Rics

            End If

            RawRics = Replace(RawRics, vbCrLf, ",")
            RawRics = Replace(RawRics, vbCr, ",")
            RawRics = Replace(RawRics, vbLf, ",")
            RawRics = Replace(RawRics, ";", ",")
            RawRics = Replace(RawRics, "|", ",")

            Parts = Split(RawRics, ",")

            For Each Part In Parts

                Ric = UCase(Trim(CStr(Part)))

                AddUniqueText Rics, Ric

            Next Part

        End If

    Next r


    For Each IsinKey In IsinToRics.Keys

        Set Rics = IsinToRics(IsinKey)
        Set UnderlyingComponents = New Collection

        TopLevelWeight = 1 / Rics.Count

        For Each Part In Rics

            Ric = CStr(Part)

            If ExpandedRicCache.Exists(Ric) Then

                Set RicComponents = ExpandedRicCache(Ric)

            Else

                Set RecursionPath = _
                    CreateObject("Scripting.Dictionary")
                RecursionPath.CompareMode = vbTextCompare

                Set RicComponents = _
                    ExpandCertificateRic( _
                        Ric, _
                        RicToName, _
                        RicToISIN, _
                        RicToAssetClass, _
                        RicToBasketRics, _
                        IsinToRic, _
                        NameToRic, _
                        RecursionPath, _
                        0)

                ExpandedRicCache.Add Ric, RicComponents

            End If

            AppendScaledCertificateComponents _
                UnderlyingComponents, _
                RicComponents, _
                TopLevelWeight

        Next Part

        For Each Part In UnderlyingComponents

            Set Component = Part

            If CBool(Component("Temporary")) Then

                MissingRicCount = MissingRicCount + 1

            End If

            UnderlyingKey = _
                NormalizeExactNameKey( _
                    CStr(Component("Name")))

            If Left( _
                    CStr(Component("Name")), _
                    Len(MISSING_CERTIFICATE_RIC_PREFIX)) <> _
               MISSING_CERTIFICATE_RIC_PREFIX And _
               Left( _
                    CStr(Component("Name")), _
                    Len(UNKNOWN_UNDERLYING_PREFIX)) <> _
               UNKNOWN_UNDERLYING_PREFIX Then

                If CStr(Component("ReferenceISIN")) <> "" Then

                    If Not UnderlyingReferenceIsinMap.Exists( _
                            UnderlyingKey) Then

                        UnderlyingReferenceIsinMap.Add _
                            UnderlyingKey, _
                            CStr(Component("ReferenceISIN"))

                    End If

                End If

                If CStr(Component("AssetClass")) <> "" Then

                    If Not UnderlyingAssetClassMap.Exists( _
                            UnderlyingKey) Then

                        UnderlyingAssetClassMap.Add _
                            UnderlyingKey, _
                            CStr(Component("AssetClass"))

                    End If

                End If

            End If

        Next Part

        If UnderlyingComponents.Count > 0 Then

            Mapping.Add CStr(IsinKey), UnderlyingComponents

        End If

    Next IsinKey


    MappingReady = (Mapping.Count > 0)

    If MissingRicCount > 0 Then

        MappingIssue = _
            CStr(MissingRicCount) & _
            " underlying RIC(s) were not found in the latest InstrumentList."

    ElseIf Not MappingReady Then

        MappingIssue = _
            "No valid ISIN / Underlying RIC mappings were loaded from " & _
            "the latest certificate reference files."

    End If

    GoTo CloseFiles

LoadFailed:

    MappingReady = False
    Mapping.RemoveAll
    UnderlyingReferenceIsinMap.RemoveAll
    UnderlyingAssetClassMap.RemoveAll

    MappingIssue = _
        "Certificate reference files could not be loaded: " & _
        Err.Description

    Err.Clear

CloseFiles:

    On Error Resume Next

    If InstrumentOpenedByCode Then wbInstrument.Close SaveChanges:=False
    If SearchOpenedByCode Then wbSearch.Close SaveChanges:=False

    On Error GoTo 0


ReturnMapping:

    Set LoadCertificateUnderlyingMap = Mapping

End Function

Private Sub AddUniqueText( _
    ByVal Items As Collection, _
    ByVal NewText As String)

    Dim Item As Variant

    If NewText = "" Then Exit Sub

    For Each Item In Items

        If StrComp( _
                CStr(Item), _
                NewText, _
                vbTextCompare) = 0 Then

            Exit Sub

        End If

    Next Item

    Items.Add NewText

End Sub

Private Function NormalizeExactNameKey( _
    ByVal InputText As String) As String

    Dim Result As String

    Result = Replace(InputText, Chr(160), " ")
    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)

    Do While InStr(Result, "  ") > 0

        Result = Replace(Result, "  ", " ")

    Loop

    NormalizeExactNameKey = UCase(Result)

End Function

Private Function RemoveEntityDiacritics( _
    ByVal InputText As String) As String

    Dim Result As String
    Dim Code As Variant

    Result = UCase(InputText)

    For Each Code In Array( _
        &HC0, &HC1, &HC2, &HC3, &HC4, &HC5, _
        &HE0, &HE1, &HE2, &HE3, &HE4, &HE5)

        Result = Replace(Result, ChrW(CLng(Code)), "A")

    Next Code

    Result = Replace(Result, ChrW(&HC6), "AE")
    Result = Replace(Result, ChrW(&HE6), "AE")

    For Each Code In Array(&HC7, &HE7, &H106, &H10C)

        Result = Replace(Result, ChrW(CLng(Code)), "C")

    Next Code

    Result = Replace(Result, ChrW(&HD0), "D")
    Result = Replace(Result, ChrW(&HF0), "D")

    For Each Code In Array( _
        &HC8, &HC9, &HCA, &HCB, _
        &HE8, &HE9, &HEA, &HEB)

        Result = Replace(Result, ChrW(CLng(Code)), "E")

    Next Code

    For Each Code In Array( _
        &HCC, &HCD, &HCE, &HCF, _
        &HEC, &HED, &HEE, &HEF)

        Result = Replace(Result, ChrW(CLng(Code)), "I")

    Next Code

    For Each Code In Array(&H141, &H142)

        Result = Replace(Result, ChrW(CLng(Code)), "L")

    Next Code

    For Each Code In Array(&HD1, &HF1)

        Result = Replace(Result, ChrW(CLng(Code)), "N")

    Next Code

    For Each Code In Array( _
        &HD2, &HD3, &HD4, &HD5, &HD6, &HD8, _
        &HF2, &HF3, &HF4, &HF5, &HF6, &HF8)

        Result = Replace(Result, ChrW(CLng(Code)), "O")

    Next Code

    Result = Replace(Result, ChrW(&H152), "OE")
    Result = Replace(Result, ChrW(&H153), "OE")

    For Each Code In Array(&H158, &H159)

        Result = Replace(Result, ChrW(CLng(Code)), "R")

    Next Code

    For Each Code In Array(&H15A, &H15B, &H160, &H161)

        Result = Replace(Result, ChrW(CLng(Code)), "S")

    Next Code

    For Each Code In Array( _
        &HD9, &HDA, &HDB, &HDC, _
        &HF9, &HFA, &HFB, &HFC)

        Result = Replace(Result, ChrW(CLng(Code)), "U")

    Next Code

    For Each Code In Array(&HDD, &HFD, &HFF, &H178)

        Result = Replace(Result, ChrW(CLng(Code)), "Y")

    Next Code

    For Each Code In Array(&H179, &H17A, &H17B, &H17C)

        Result = Replace(Result, ChrW(CLng(Code)), "Z")

    Next Code

    Result = Replace(Result, ChrW(&HDE), "TH")
    Result = Replace(Result, ChrW(&HFE), "TH")
    Result = Replace(Result, ChrW(&HDF), "SS")

    RemoveEntityDiacritics = Result

End Function

Private Function IsEntityShareClassSuffix( _
    ByVal SuffixText As String) As Boolean

    Dim CompactSuffix As String
    Dim Character As String
    Dim CharacterCode As Long
    Dim ClassCodeLength As Long
    Dim RemainingText As String
    Dim i As Long

    SuffixText = _
        RemoveEntityDiacritics(UCase(Trim(SuffixText)))

    For i = 1 To Len(SuffixText)

        Character = Mid(SuffixText, i, 1)
        CharacterCode = AscW(Character)

        If (CharacterCode >= 48 And CharacterCode <= 57) Or _
           (CharacterCode >= 65 And CharacterCode <= 90) Then

            CompactSuffix = CompactSuffix & Character

        End If

    Next i

    ' Share-class codes are normally one or two characters: A, B, C, A1,
    ' and similar. Longer words such as "Inc" must not be mistaken for one.
    For ClassCodeLength = 1 To 2

        If Len(CompactSuffix) >= ClassCodeLength Then

            RemainingText = _
                Mid(CompactSuffix, ClassCodeLength + 1)

            Select Case RemainingText

                Case "", _
                     "SHARE", _
                     "SHARES", _
                     "COMMONSTOCK", _
                     "COMMONSHARE", _
                     "COMMONSHARES", _
                     "ORDINARYSHARE", _
                     "ORDINARYSHARES"

                    IsEntityShareClassSuffix = True

                    Exit Function

            End Select

        End If

    Next ClassCodeLength

End Function

Private Function RemoveTrailingEntityShareClass( _
    ByVal InputText As String) As String

    Dim Result As String
    Dim UpperResult As String
    Dim ClassPosition As Long
    Dim SuffixText As String

    Result = Replace(InputText, ChrW(&HA0), " ")
    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)

    Do While InStr(Result, "  ") > 0

        Result = Replace(Result, "  ", " ")

    Loop

    UpperResult = UCase(Result)
    ClassPosition = InStrRev(UpperResult, " CLASS ")

    If ClassPosition > 1 Then

        SuffixText = _
            Mid(Result, ClassPosition + Len(" CLASS "))

        If IsEntityShareClassSuffix(SuffixText) Then

            Result = Trim(Left(Result, ClassPosition - 1))

        End If

    End If

    RemoveTrailingEntityShareClass = Result

End Function

Private Function NormalizeEntityLegalSuffix( _
    ByVal EntityKey As String) As String

    Dim LongSuffixes As Variant
    Dim ShortSuffixes As Variant
    Dim LongSuffix As String
    Dim i As Long

    ' Normalize common legal-form spellings only when they occur at the end
    ' of a name. This is more aggressive than punctuation removal but avoids
    ' replacing the same letter sequence inside an ordinary company name.
    LongSuffixes = Array( _
        "GESELLSCHAFTMITBESCHRAENKTERHAFTUNG", _
        "GESELLSCHAFTMITBESCHRANKTERHAFTUNG", _
        "SOCIETEPARACTIONSSIMPLIFIEE", _
        "LIMITEDLIABILITYCOMPANY", _
        "PUBLICLIMITEDCOMPANY", _
        "NAAMLOZEVENNOOTSCHAP", _
        "SOCIETAPERAZIONI", _
        "SOCIETEANONYME", _
        "SOCIEDADANONIMA", _
        "AKTIENGESELLSCHAFT", _
        "INCORPORATED", _
        "CORPORATION", _
        "COMPANY", _
        "LIMITED")

    ShortSuffixes = Array( _
        "GMBH", _
        "GMBH", _
        "SAS", _
        "LLC", _
        "PLC", _
        "NV", _
        "SPA", _
        "SA", _
        "SA", _
        "AG", _
        "INC", _
        "CORP", _
        "CO", _
        "LTD")

    For i = LBound(LongSuffixes) To UBound(LongSuffixes)

        LongSuffix = CStr(LongSuffixes(i))

        If Len(EntityKey) > Len(LongSuffix) And _
           Right(EntityKey, Len(LongSuffix)) = LongSuffix Then

            NormalizeEntityLegalSuffix = _
                Left( _
                    EntityKey, _
                    Len(EntityKey) - Len(LongSuffix)) & _
                CStr(ShortSuffixes(i))

            Exit Function

        End If

    Next i

    NormalizeEntityLegalSuffix = EntityKey

End Function

Private Function NormalizeEntityKey( _
    ByVal InputText As String) As String

    Dim Result As String
    Dim Character As String
    Dim CharacterCode As Long
    Dim CollapsedKey As String
    Dim i As Long

    Result = _
        RemoveTrailingEntityShareClass(InputText)
    Result = Replace(Result, ChrW(&HA0), " ")
    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)
    Result = UCase(Result)

    ' Bloomberg-style names often append the leading article as /The.
    If Right(Result, 4) = "/THE" Then

        Result = Left(Result, Len(Result) - 4)

    ElseIf Right(Result, 5) = ", THE" Then

        Result = Left(Result, Len(Result) - 5)

    End If

    If Left(Result, 4) = "THE " Then

        Result = Mid(Result, 5)

    End If

    Result = Replace(Result, "&", " AND ")
    Result = RemoveEntityDiacritics(Result)

    For i = 1 To Len(Result)

        Character = Mid(Result, i, 1)
        CharacterCode = AscW(Character)

        If (CharacterCode >= 48 And CharacterCode <= 57) Or _
           (CharacterCode >= 65 And CharacterCode <= 90) Then

            CollapsedKey = CollapsedKey & Character

        End If

    Next i

    NormalizeEntityKey = _
        NormalizeEntityLegalSuffix(CollapsedKey)

End Function

Private Function TrimEntityDisplayPunctuation( _
    ByVal InputText As String) As String

    Dim Result As String
    Dim LastCharacter As String

    Result = Trim(InputText)

    Do While Len(Result) > 0

        LastCharacter = Right(Result, 1)

        Select Case LastCharacter

            Case ".", ",", ";", ":"

                Result = _
                    Trim(Left(Result, Len(Result) - 1))

            Case Else

                Exit Do

        End Select

    Loop

    TrimEntityDisplayPunctuation = Result

End Function

Private Function StandardizeEntityDisplayName( _
    ByVal InputText As String) As String

    Dim Result As String
    Dim UpperResult As String
    Dim CandidateTail As String
    Dim CandidateKey As String
    Dim BaseName As String
    Dim SuffixCodes As Variant
    Dim SuffixDisplays As Variant
    Dim IsWordBoundary As Boolean
    Dim i As Long
    Dim j As Long

    Result = _
        RemoveTrailingEntityShareClass(InputText)
    Result = Replace(Result, ChrW(&HA0), " ")
    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)

    Do While InStr(Result, "  ") > 0

        Result = Replace(Result, "  ", " ")

    Loop

    UpperResult = UCase(Result)

    If Right(UpperResult, 4) = "/THE" Then

        Result = Trim(Left(Result, Len(Result) - 4))

    ElseIf Right(UpperResult, 5) = ", THE" Then

        Result = Trim(Left(Result, Len(Result) - 5))

    End If

    If Left(UCase(Result), 4) = "THE " Then

        Result = Trim(Mid(Result, 5))

    End If

    Result = TrimEntityDisplayPunctuation(Result)

    SuffixCodes = Array( _
        "GMBH", _
        "CORP", _
        "INC", _
        "CO", _
        "LTD", _
        "PLC", _
        "LLC", _
        "AG", _
        "SA", _
        "SPA", _
        "NV", _
        "SAS", _
        "SE", _
        "SCA", _
        "KGAA", _
        "LP", _
        "LLP")

    SuffixDisplays = Array( _
        "GmbH", _
        "Corp", _
        "Inc", _
        "Co", _
        "Ltd", _
        "PLC", _
        "LLC", _
        "AG", _
        "SA", _
        "SpA", _
        "NV", _
        "SAS", _
        "SE", _
        "SCA", _
        "KGaA", _
        "LP", _
        "LLP")

    ' Test each word-boundary tail. Prefixing X lets NormalizeEntityKey
    ' expand a standalone long legal form such as "Corporation" safely.
    For i = 1 To Len(Result)

        IsWordBoundary = (i = 1)

        If Not IsWordBoundary Then

            IsWordBoundary = _
                (Mid(Result, i - 1, 1) = " ")

        End If

        If IsWordBoundary Then

            CandidateTail = Mid(Result, i)
            CandidateKey = _
                NormalizeEntityKey("X " & CandidateTail)

            For j = LBound(SuffixCodes) To UBound(SuffixCodes)

                If CandidateKey = _
                    "X" & CStr(SuffixCodes(j)) Then

                    BaseName = _
                        TrimEntityDisplayPunctuation( _
                            Left(Result, i - 1))

                    If BaseName <> "" Then

                        StandardizeEntityDisplayName = _
                            BaseName & _
                            " " & _
                            CStr(SuffixDisplays(j))

                    Else

                        StandardizeEntityDisplayName = Result

                    End If

                    Exit Function

                End If

            Next j

        End If

    Next i

    StandardizeEntityDisplayName = Result

End Function

Private Function IsEntityLegalSuffixOnlyMatch( _
    ByVal FirstKey As String, _
    ByVal SecondKey As String) As Boolean

    Dim ShortKey As String
    Dim LongKey As String
    Dim AdditionalSuffix As String

    If FirstKey = "" Or SecondKey = "" Then Exit Function
    If FirstKey = SecondKey Then Exit Function

    If Len(FirstKey) < Len(SecondKey) Then

        ShortKey = FirstKey
        LongKey = SecondKey

    Else

        ShortKey = SecondKey
        LongKey = FirstKey

    End If

    If Len(ShortKey) < _
        ENTITY_LEGAL_SUFFIX_ROOT_MIN_LENGTH Then Exit Function

    If Left(LongKey, Len(ShortKey)) <> ShortKey Then Exit Function

    AdditionalSuffix = Mid(LongKey, Len(ShortKey) + 1)

    Select Case AdditionalSuffix

        Case "GMBH", _
             "CORP", _
             "INC", _
             "CO", _
             "LTD", _
             "PLC", _
             "LLC", _
             "AG", _
             "SA", _
             "SPA", _
             "NV", _
             "SAS", _
             "SE", _
             "SCA", _
             "KGAA", _
             "LP", _
             "LLP"

            IsEntityLegalSuffixOnlyMatch = True

    End Select

End Function

Private Function IsLikelyEntityPrefixMatch( _
    ByVal FirstKey As String, _
    ByVal SecondKey As String) As Boolean

    Dim ShortKey As String
    Dim LongKey As String
    Dim MatchRatio As Double

    If FirstKey = "" Or SecondKey = "" Then Exit Function

    If IsEntityLegalSuffixOnlyMatch( _
            FirstKey, _
            SecondKey) Then

        IsLikelyEntityPrefixMatch = True

        Exit Function

    End If

    If Len(FirstKey) <= Len(SecondKey) Then

        ShortKey = FirstKey
        LongKey = SecondKey

    Else

        ShortKey = SecondKey
        LongKey = FirstKey

    End If

    If Len(ShortKey) < ENTITY_PREFIX_MIN_LENGTH Then Exit Function

    If Left(LongKey, Len(ShortKey)) <> ShortKey Then Exit Function

    MatchRatio = Len(ShortKey) / Len(LongKey)

    IsLikelyEntityPrefixMatch = _
        (MatchRatio >= ENTITY_PREFIX_MIN_RATIO)

End Function

Private Function NewExactNameMap() As Object

    Dim Mapping As Object

    Set Mapping = CreateObject("Scripting.Dictionary")
    Mapping.CompareMode = vbTextCompare

    Set NewExactNameMap = Mapping

End Function

Private Function ResolveExactName( _
    ByVal LookupText As String, _
    ByVal Mapping As Object) As String

    Dim LookupKey As String

    If Mapping Is Nothing Then Exit Function

    LookupKey = NormalizeExactNameKey(LookupText)

    If LookupKey = "" Then Exit Function

    If Mapping.Exists(LookupKey) Then

        ResolveExactName = CStr(Mapping(LookupKey))

    End If

End Function

Private Function LoadEquityNameMap( _
    ByRef MappingReady As Boolean, _
    ByVal PreferredWorkbook As Workbook) As Object

    Dim Mapping As Object
    Dim DataTable As ListObject
    Dim TableData As Variant
    Dim IsinColumn As Long
    Dim NameColumn As Long
    Dim r As Long
    Dim LookupKey As String
    Dim SecurityName As String

    Set Mapping = NewExactNameMap()

    Set DataTable = _
        GetReferenceDataTable( _
            EQUITY_NAMES_SHEET, _
            UNMAPPED_EQUITIES_TABLE, _
            PreferredWorkbook)

    If DataTable Is Nothing Then

        Set LoadEquityNameMap = Mapping

        Exit Function

    End If

    IsinColumn = GetTableColumnIndex(DataTable, "ISIN")
    NameColumn = _
        GetTableColumnIndex(DataTable, "Security Name")

    If IsinColumn = 0 Or NameColumn = 0 Or _
       DataTable.DataBodyRange Is Nothing Then

        Set LoadEquityNameMap = Mapping

        Exit Function

    End If

    TableData = DataTable.DataBodyRange.Value2

    For r = 1 To UBound(TableData, 1)

        LookupKey = _
            NormalizeExactNameKey( _
                SafeText( _
                    TableData(r, IsinColumn)))

        SecurityName = _
            SafeText( _
                TableData(r, NameColumn))

        If LookupKey <> "" And SecurityName <> "" Then

            Mapping(LookupKey) = SecurityName

        End If

    Next r

    MappingReady = (Mapping.Count > 0)

    Set LoadEquityNameMap = Mapping

End Function

Private Sub LoadBondIssuerMaps( _
    ByRef IssuerMapping As Object, _
    ByRef TypeMapping As Object, _
    ByRef MappingReady As Boolean, _
    ByVal PreferredWorkbook As Workbook)

    Dim DataTable As ListObject
    Dim TableData As Variant
    Dim TickerColumn As Long
    Dim IssuerColumn As Long
    Dim TypeColumn As Long
    Dim r As Long
    Dim LookupKey As String
    Dim IssuerName As String
    Dim IssuerType As String

    Set IssuerMapping = NewExactNameMap()
    Set TypeMapping = NewExactNameMap()

    Set DataTable = _
        GetReferenceDataTable( _
            BOND_ISSUERS_SHEET, _
            BOND_ISSUERS_TABLE, _
            PreferredWorkbook)

    If DataTable Is Nothing Then Exit Sub

    TickerColumn = _
        GetTableColumnIndex(DataTable, "Issuer Ticker")
    IssuerColumn = _
        GetTableColumnIndex(DataTable, "Issuer Name")
    TypeColumn = _
        GetTableColumnIndex(DataTable, "Bond Type")

    If TickerColumn = 0 Or IssuerColumn = 0 Or _
       TypeColumn = 0 Or DataTable.DataBodyRange Is Nothing Then

        Exit Sub

    End If

    TableData = DataTable.DataBodyRange.Value2

    For r = 1 To UBound(TableData, 1)

        LookupKey = _
            NormalizeExactNameKey( _
                SafeText( _
                    TableData(r, TickerColumn)))

        IssuerName = _
            SafeText( _
                TableData(r, IssuerColumn))

        IssuerType = _
            SafeText( _
                TableData(r, TypeColumn))

        If LookupKey <> "" Then

            If IssuerName <> "" Then

                IssuerMapping(LookupKey) = IssuerName

            End If

            If IssuerType <> "" Then

                TypeMapping(LookupKey) = IssuerType

            End If

        End If

    Next r

    MappingReady = _
        (IssuerMapping.Count > 0 Or TypeMapping.Count > 0)

End Sub

Private Function LoadCountryNameMap( _
    ByVal PreferredWorkbook As Workbook) As Object

    Dim Mapping As Object
    Dim wsCountries As Worksheet
    Dim TableData As Variant
    Dim LastRow As Long
    Dim CandidateRow As Long
    Dim r As Long
    Dim LookupKey As String
    Dim CountryName As String

    Set Mapping = NewExactNameMap()

    Set wsCountries = _
        GetReferenceDataWorksheet( _
            COUNTRIES_SHEET, _
            PreferredWorkbook)

    If wsCountries Is Nothing Then

        Set LoadCountryNameMap = Mapping

        Exit Function

    End If

    LastRow = 2

    CandidateRow = _
        wsCountries.Cells( _
            wsCountries.Rows.Count, _
            2).End(xlUp).Row

    If CandidateRow > LastRow Then LastRow = CandidateRow

    CandidateRow = _
        wsCountries.Cells( _
            wsCountries.Rows.Count, _
            3).End(xlUp).Row

    If CandidateRow > LastRow Then LastRow = CandidateRow

    If LastRow < 3 Then

        Set LoadCountryNameMap = Mapping

        Exit Function

    End If

    TableData = _
        wsCountries.Range( _
            wsCountries.Cells(3, 2), _
            wsCountries.Cells(LastRow, 3)).Value2

    For r = 1 To UBound(TableData, 1)

        LookupKey = _
            NormalizeExactNameKey( _
                SafeText(TableData(r, 1)))

        CountryName = _
            SafeText(TableData(r, 2))

        If LookupKey <> "" And CountryName <> "" Then

            Mapping(LookupKey) = CountryName

        End If

    Next r

    Set LoadCountryNameMap = Mapping

End Function

Private Function LoadFundParentCompanyMap( _
    ByRef MappingReady As Boolean, _
    ByVal PreferredWorkbook As Workbook) As Object

    Dim FundMap As Object
    Dim DataTable As ListObject
    Dim TableData As Variant
    Dim FundNameColumn As Long
    Dim CompanyNameColumn As Long
    Dim r As Long
    Dim LookupKey As String
    Dim CompanyName As String

    Set FundMap = NewExactNameMap()

    Set DataTable = _
        GetReferenceDataTable( _
            FUND_PARENT_COMPANIES_SHEET, _
            FUNDS_TABLE, _
            PreferredWorkbook)

    If DataTable Is Nothing Then

        Set LoadFundParentCompanyMap = FundMap

        Exit Function

    End If

    FundNameColumn = _
        GetTableColumnIndex(DataTable, "Fund Name")
    CompanyNameColumn = _
        GetTableColumnIndex(DataTable, "Company Name")

    If FundNameColumn = 0 Or CompanyNameColumn = 0 Or _
       DataTable.DataBodyRange Is Nothing Then

        Set LoadFundParentCompanyMap = FundMap

        Exit Function

    End If

    DataTable.Range.Calculate

    TableData = DataTable.DataBodyRange.Value2

    For r = 1 To UBound(TableData, 1)

        LookupKey = _
            NormalizeExactNameKey( _
                SafeText( _
                    TableData(r, FundNameColumn)))

        CompanyName = _
            SafeText( _
                TableData(r, CompanyNameColumn))

        If LookupKey <> "" And CompanyName <> "" And _
           CompanyName <> "0" Then

            FundMap(LookupKey) = CompanyName

        End If

    Next r

    MappingReady = (FundMap.Count > 0)

    Set LoadFundParentCompanyMap = FundMap

End Function

Private Function ResolveFundParentCompany( _
    ByVal FundMap As Object, _
    ByVal ISIN As String, _
    ByVal SecurityName As String, _
    ByVal SourceRow As Long) As String

    Dim CompanyName As String
    Dim PlaceholderMode As IssuerPlaceholderMode

    CompanyName = _
        ResolveExactName( _
            SecurityName, _
            FundMap)

    If CompanyName <> "" Then

        ResolveFundParentCompany = CompanyName

        Exit Function

    End If

    If Trim(SecurityName) = "" Then

        PlaceholderMode = PlaceholderFromISIN

    Else

        PlaceholderMode = PlaceholderFromSecurityName

    End If

    ResolveFundParentCompany = _
        BuildTemporaryIssuerName( _
            ISIN, _
            SecurityName, _
            SourceRow, _
            PlaceholderMode)

End Function

Private Function AppendUniqueGeographyType( _
    ByVal ExistingTypes As String, _
    ByVal NewType As String) As String

    Dim NewParts As Variant
    Dim NewPart As Variant
    Dim Parts As Variant
    Dim Part As Variant
    Dim TypeText As String
    Dim ResultText As String
    Dim TypeExists As Boolean

    NewType = Trim(NewType)

    If NewType = "" Then

        AppendUniqueGeographyType = ExistingTypes

        Exit Function

    End If

    ResultText = ExistingTypes
    NewParts = Split(NewType, ";")

    For Each NewPart In NewParts

        TypeText = Trim(CStr(NewPart))

        If TypeText <> "" Then

            TypeExists = False

            If Trim(ResultText) <> "" Then

                Parts = Split(ResultText, ";")

                For Each Part In Parts

                    If StrComp( _
                            Trim(CStr(Part)), _
                            TypeText, _
                            vbTextCompare) = 0 Then

                        TypeExists = True

                        Exit For

                    End If

                Next Part

            End If

            If Not TypeExists Then

                If Trim(ResultText) = "" Then

                    ResultText = TypeText

                Else

                    ResultText = _
                        ResultText & "; " & TypeText

                End If

            End If

        End If

    Next NewPart

    AppendUniqueGeographyType = ResultText

End Function

Private Function AppendUniqueGeographyVariant( _
    ByVal ExistingVariants As String, _
    ByVal NewVariant As String) As String

    Dim Parts As Variant
    Dim Part As Variant

    NewVariant = Trim(NewVariant)

    If NewVariant = "" Then

        AppendUniqueGeographyVariant = ExistingVariants

        Exit Function

    End If

    If Trim(ExistingVariants) = "" Then

        AppendUniqueGeographyVariant = NewVariant

        Exit Function

    End If

    Parts = Split(ExistingVariants, ";")

    For Each Part In Parts

        If StrComp( _
                Trim(CStr(Part)), _
                NewVariant, _
                vbBinaryCompare) = 0 Then

            AppendUniqueGeographyVariant = ExistingVariants

            Exit Function

        End If

    Next Part

    AppendUniqueGeographyVariant = _
        ExistingVariants & "; " & NewVariant

End Function

Private Sub AddGeographyLookupEntry( _
    ByVal GeographyEntries As Object, _
    ByVal ExposureName As String, _
    ByVal ExposureType As String, _
    ByVal ReferenceISIN As String, _
    ByVal IsinRelationship As String, _
    ByVal IsinPriority As Long, _
    ByVal ObservationSet As Object, _
    Optional ByVal NameVariant As String = "")

    Dim EntryKey As String
    Dim ObservationKey As String
    Dim Entry As Object
    Dim IsinCandidates As Object
    Dim Candidate As Object
    Dim ExistingPriority As Long

    If GeographyEntries Is Nothing Then Exit Sub

    ExposureName = Trim(ExposureName)

    If ExposureName = "" Then Exit Sub

    ' Square-bracket values are report placeholders, not real names suitable
    ' for a Bloomberg geography query.
    If Left(ExposureName, 1) = "[" And _
       Right(ExposureName, 1) = "]" Then Exit Sub

    EntryKey = NormalizeExactNameKey(ExposureName)

    If EntryKey = "" Then Exit Sub

    If NameVariant = "" Then NameVariant = ExposureName

    ReferenceISIN = UCase(Trim(ReferenceISIN))

    If Not ObservationSet Is Nothing Then

        ObservationKey = _
            EntryKey & Chr$(30) & _
            UCase$(Trim$(ExposureType)) & Chr$(30) & _
            ReferenceISIN & Chr$(30) & _
            UCase$(Trim$(IsinRelationship)) & Chr$(30) & _
            CStr(IsinPriority) & Chr$(30) & _
            Trim$(NameVariant)

        If ObservationSet.Exists(ObservationKey) Then Exit Sub

        ObservationSet.Add ObservationKey, True

    End If

    If GeographyEntries.Exists(EntryKey) Then

        Set Entry = GeographyEntries(EntryKey)

    Else

        Set Entry = CreateObject("Scripting.Dictionary")
        Entry.CompareMode = vbTextCompare

        Entry.Add "Name", ExposureName
        Entry.Add "ExposureType", ""
        Entry.Add "ReferenceISIN", ""
        Entry.Add "IsinRelationship", ""
        Entry.Add "IsinPriority", 999&
        Entry.Add "Variants", ""

        Set IsinCandidates = _
            CreateObject("Scripting.Dictionary")
        IsinCandidates.CompareMode = vbTextCompare

        Entry.Add "IsinCandidates", IsinCandidates

        GeographyEntries.Add EntryKey, Entry

    End If

    If Not Entry.Exists("IsinCandidates") Then

        Set IsinCandidates = _
            CreateObject("Scripting.Dictionary")
        IsinCandidates.CompareMode = vbTextCompare

        Entry.Add "IsinCandidates", IsinCandidates

    Else

        Set IsinCandidates = Entry("IsinCandidates")

    End If

    Entry("ExposureType") = _
        AppendUniqueGeographyType( _
            CStr(Entry("ExposureType")), _
            ExposureType)

    Entry("Variants") = _
        AppendUniqueGeographyVariant( _
            CStr(Entry("Variants")), _
            NameVariant)

    If ReferenceISIN = "" Then Exit Sub

    If IsinCandidates.Exists(ReferenceISIN) Then

        Set Candidate = IsinCandidates(ReferenceISIN)

        If IsinPriority < CLng(Candidate("Priority")) Then

            Candidate("Relationship") = IsinRelationship
            Candidate("Relationships") = IsinRelationship
            Candidate("Priority") = IsinPriority

        ElseIf IsinPriority = CLng(Candidate("Priority")) Then

            Candidate("Relationships") = _
                AppendUniqueGeographyType( _
                    CStr(Candidate("Relationships")), _
                    IsinRelationship)

            If StrComp( _
                    IsinRelationship, _
                    CStr(Candidate("Relationship")), _
                    vbTextCompare) < 0 Then

                Candidate("Relationship") = IsinRelationship

            End If

        End If

    Else

        Set Candidate = CreateObject("Scripting.Dictionary")
        Candidate.CompareMode = vbTextCompare

        Candidate.Add "Relationship", IsinRelationship
        Candidate.Add "Relationships", IsinRelationship
        Candidate.Add "Priority", IsinPriority

        IsinCandidates.Add ReferenceISIN, Candidate

    End If

    ExistingPriority = CLng(Entry("IsinPriority"))

    If CStr(Entry("ReferenceISIN")) = "" Or _
       IsinPriority < ExistingPriority Or _
       (IsinPriority = ExistingPriority And _
        StrComp( _
            ReferenceISIN, _
            CStr(Entry("ReferenceISIN")), _
            vbTextCompare) < 0) Or _
       (IsinPriority = ExistingPriority And _
        StrComp( _
            ReferenceISIN, _
            CStr(Entry("ReferenceISIN")), _
            vbTextCompare) = 0 And _
        StrComp( _
            IsinRelationship, _
            CStr(Entry("IsinRelationship")), _
            vbTextCompare) < 0) Then

        Entry("ReferenceISIN") = ReferenceISIN
        Entry("IsinRelationship") = IsinRelationship
        Entry("IsinPriority") = IsinPriority

    End If

End Sub

Private Function MergeGeographyVariants( _
    ByVal ExistingVariants As String, _
    ByVal LeadingVariant As String, _
    ByVal NewVariants As String) As String

    Dim Parts As Variant
    Dim Part As Variant
    Dim ResultText As String

    ResultText = _
        AppendUniqueGeographyVariant( _
            ExistingVariants, _
            LeadingVariant)

    Parts = Split(NewVariants, ";")

    For Each Part In Parts

        ResultText = _
            AppendUniqueGeographyVariant( _
                ResultText, _
                Trim(CStr(Part)))

    Next Part

    MergeGeographyVariants = ResultText

End Function

Private Function CopyGeographyCandidate( _
    ByVal SourceCandidate As Object) As Object

    Dim CandidateCopy As Object
    Dim RelationshipText As String

    Set CandidateCopy = _
        CreateObject("Scripting.Dictionary")
    CandidateCopy.CompareMode = vbTextCompare

    RelationshipText = _
        CStr(SourceCandidate("Relationship"))

    CandidateCopy.Add "Relationship", RelationshipText

    If SourceCandidate.Exists("Relationships") Then

        CandidateCopy.Add _
            "Relationships", _
            CStr(SourceCandidate("Relationships"))

    Else

        CandidateCopy.Add "Relationships", RelationshipText

    End If

    CandidateCopy.Add _
        "Priority", _
        CLng(SourceCandidate("Priority"))

    Set CopyGeographyCandidate = CandidateCopy

End Function

Private Sub MergeGeographyCandidateSets( _
    ByVal TargetEntry As Object, _
    ByVal SourceEntry As Object)

    Dim TargetCandidates As Object
    Dim SourceCandidates As Object
    Dim TargetCandidate As Object
    Dim SourceCandidate As Object
    Dim CandidateCopy As Object
    Dim CandidateISIN As Variant
    Dim TargetRelationships As String
    Dim SourceRelationships As String
    Dim TargetPriority As Long
    Dim SourcePriority As Long

    If TargetEntry Is Nothing Then Exit Sub
    If SourceEntry Is Nothing Then Exit Sub
    If Not SourceEntry.Exists("IsinCandidates") Then Exit Sub

    Set SourceCandidates = SourceEntry("IsinCandidates")

    If SourceCandidates Is Nothing Then Exit Sub
    If SourceCandidates.Count = 0 Then Exit Sub

    If TargetEntry.Exists("IsinCandidates") Then

        Set TargetCandidates = TargetEntry("IsinCandidates")

    Else

        Set TargetCandidates = _
            CreateObject("Scripting.Dictionary")
        TargetCandidates.CompareMode = vbTextCompare
        TargetEntry.Add "IsinCandidates", TargetCandidates

    End If

    For Each CandidateISIN In SourceCandidates.Keys

        Set SourceCandidate = _
            SourceCandidates(CandidateISIN)

        If TargetCandidates.Exists(CandidateISIN) Then

            Set TargetCandidate = _
                TargetCandidates(CandidateISIN)

            TargetPriority = _
                CLng(TargetCandidate("Priority"))
            SourcePriority = _
                CLng(SourceCandidate("Priority"))

            If SourcePriority < TargetPriority Then

                TargetCandidate("Relationship") = _
                    CStr(SourceCandidate("Relationship"))

                If SourceCandidate.Exists("Relationships") Then

                    TargetCandidate("Relationships") = _
                        CStr(SourceCandidate("Relationships"))

                Else

                    TargetCandidate("Relationships") = _
                        CStr(SourceCandidate("Relationship"))

                End If

                TargetCandidate("Priority") = SourcePriority

            ElseIf SourcePriority = TargetPriority Then

                If TargetCandidate.Exists("Relationships") Then

                    TargetRelationships = _
                        CStr(TargetCandidate("Relationships"))

                Else

                    TargetRelationships = _
                        CStr(TargetCandidate("Relationship"))

                End If

                If SourceCandidate.Exists("Relationships") Then

                    SourceRelationships = _
                        CStr(SourceCandidate("Relationships"))

                Else

                    SourceRelationships = _
                        CStr(SourceCandidate("Relationship"))

                End If

                TargetCandidate("Relationships") = _
                    AppendUniqueGeographyType( _
                        TargetRelationships, _
                        SourceRelationships)

                If StrComp( _
                        CStr(SourceCandidate("Relationship")), _
                        CStr(TargetCandidate("Relationship")), _
                        vbTextCompare) < 0 Then

                    TargetCandidate("Relationship") = _
                        CStr(SourceCandidate("Relationship"))

                End If

            End If

        Else

            Set CandidateCopy = _
                CopyGeographyCandidate(SourceCandidate)

            TargetCandidates.Add _
                CStr(CandidateISIN), _
                CandidateCopy

        End If

    Next CandidateISIN

End Sub

Private Sub MergeGeographyReferenceSelection( _
    ByVal TargetEntry As Object, _
    ByVal SourceEntry As Object)

    Dim TargetISIN As String
    Dim SourceISIN As String
    Dim TargetPriority As Long
    Dim SourcePriority As Long
    Dim SourceIsBetter As Boolean

    If TargetEntry Is Nothing Then Exit Sub
    If SourceEntry Is Nothing Then Exit Sub

    SourceISIN = _
        UCase(Trim(CStr(SourceEntry("ReferenceISIN"))))

    If SourceISIN = "" Then Exit Sub

    TargetISIN = _
        UCase(Trim(CStr(TargetEntry("ReferenceISIN"))))
    SourcePriority = CLng(SourceEntry("IsinPriority"))
    TargetPriority = CLng(TargetEntry("IsinPriority"))

    SourceIsBetter = _
        (TargetISIN = "" Or _
         SourcePriority < TargetPriority Or _
         (SourcePriority = TargetPriority And _
          StrComp( _
                SourceISIN, _
                TargetISIN, _
                vbTextCompare) < 0) Or _
         (SourcePriority = TargetPriority And _
          StrComp( _
                SourceISIN, _
                TargetISIN, _
                vbTextCompare) = 0 And _
          StrComp( _
                CStr(SourceEntry("IsinRelationship")), _
                CStr(TargetEntry("IsinRelationship")), _
                vbTextCompare) < 0))

    If SourceIsBetter Then

        TargetEntry("ReferenceISIN") = SourceISIN
        TargetEntry("IsinRelationship") = _
            CStr(SourceEntry("IsinRelationship"))
        TargetEntry("IsinPriority") = SourcePriority

    End If

End Sub

Private Sub AddGeographyEntryWithCandidates( _
    ByVal TargetEntries As Object, _
    ByVal TargetName As String, _
    ByVal SourceEntry As Object, _
    ByVal NameVariant As String)

    Dim EntryKey As String
    Dim TargetEntry As Object
    Dim SourceVariants As String

    If TargetEntries Is Nothing Then Exit Sub
    If SourceEntry Is Nothing Then Exit Sub

    TargetName = Trim(TargetName)

    If TargetName = "" Then Exit Sub

    If Left(TargetName, 1) = "[" And _
       Right(TargetName, 1) = "]" Then Exit Sub

    EntryKey = NormalizeExactNameKey(TargetName)

    If EntryKey = "" Then Exit Sub

    SourceVariants = CStr(SourceEntry("Variants"))

    If Not TargetEntries.Exists(EntryKey) Then

        SourceEntry("Name") = TargetName
        SourceEntry("Variants") = _
            MergeGeographyVariants( _
                "", _
                NameVariant, _
                SourceVariants)

        TargetEntries.Add EntryKey, SourceEntry

        Exit Sub

    End If

    Set TargetEntry = TargetEntries(EntryKey)

    TargetEntry("ExposureType") = _
        AppendUniqueGeographyType( _
            CStr(TargetEntry("ExposureType")), _
            CStr(SourceEntry("ExposureType")))

    TargetEntry("Variants") = _
        MergeGeographyVariants( _
            CStr(TargetEntry("Variants")), _
            NameVariant, _
            SourceVariants)

    MergeGeographyCandidateSets _
        TargetEntry, _
        SourceEntry

    MergeGeographyReferenceSelection _
        TargetEntry, _
        SourceEntry

End Sub

Private Function GetNameVariantsLastRow( _
    ByVal wsVariants As Worksheet) As Long

    Dim ColNo As Long
    Dim CandidateRow As Long

    GetNameVariantsLastRow = 1

    If wsVariants Is Nothing Then Exit Function

    For ColNo = 1 To 4

        CandidateRow = _
            wsVariants.Cells( _
                wsVariants.Rows.Count, _
                ColNo).End(xlUp).Row

        If CandidateRow > GetNameVariantsLastRow Then

            GetNameVariantsLastRow = CandidateRow

        End If

    Next ColNo

End Function

Private Function EnsureNameVariantsWorksheet() As Worksheet

    Dim wsVariants As Worksheet

    Set wsVariants = GetOptionalWorksheet(NAME_VARIANTS_SHEET)

    If wsVariants Is Nothing Then

        On Error GoTo CreateFailed

        Set wsVariants = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        wsVariants.name = NAME_VARIANTS_SHEET

    End If

    wsVariants.Cells(1, 1).Value = "Name Variant"
    wsVariants.Cells(1, 2).Value = "Canonical Name"
    wsVariants.Cells(1, 3).Value = "Normalized Key"
    wsVariants.Cells(1, 4).Value = "Manual Override"

    With wsVariants.Range( _
        wsVariants.Cells(1, 1), _
        wsVariants.Cells(1, 4))

        .Font.Bold = True

    End With

    If Not wsVariants.AutoFilterMode Then

        wsVariants.Range( _
            wsVariants.Cells(1, 1), _
            wsVariants.Cells(1, 4)).AutoFilter

    End If

    Set EnsureNameVariantsWorksheet = wsVariants

    Exit Function

CreateFailed:

    Err.Clear
    Set EnsureNameVariantsWorksheet = Nothing

End Function

Private Function LoadManualNameVariantMap() As Object

    Dim Mapping As Object
    Dim wsVariants As Worksheet
    Dim TableData As Variant
    Dim LastRow As Long
    Dim r As Long
    Dim NameVariant As String
    Dim CanonicalName As String
    Dim ManualOverride As String
    Dim VariantKey As String
    Dim HasManualOverrideColumn As Boolean

    Set Mapping = NewExactNameMap()
    Set wsVariants = GetOptionalWorksheet(NAME_VARIANTS_SHEET)

    If wsVariants Is Nothing Then

        Set LoadManualNameVariantMap = Mapping

        Exit Function

    End If

    LastRow = GetNameVariantsLastRow(wsVariants)

    HasManualOverrideColumn = _
        (NormalizeHeader( _
            SafeText(wsVariants.Cells(1, 4).Value)) = _
         "MANUALOVERRIDE")

    If LastRow < 2 Then

        Set LoadManualNameVariantMap = Mapping

        Exit Function

    End If

    TableData = _
        wsVariants.Range( _
            wsVariants.Cells(2, 1), _
            wsVariants.Cells(LastRow, 4)).Value2

    For r = 1 To UBound(TableData, 1)

        NameVariant = SafeText(TableData(r, 1))
        CanonicalName = SafeText(TableData(r, 2))

        If HasManualOverrideColumn Then

            ManualOverride = _
                SafeText(TableData(r, 4))

            If ManualOverride <> "" Then

                CanonicalName = ManualOverride

            Else

                CanonicalName = ""

            End If

        ElseIf StrComp( _
                    NormalizeExactNameKey(NameVariant), _
                    NormalizeExactNameKey(CanonicalName), _
                    vbTextCompare) = 0 Then

            ' v12 auto-filled identity rows were not manual overrides.
            CanonicalName = ""

        End If

        If NameVariant <> "" And CanonicalName <> "" Then

            VariantKey = NormalizeExactNameKey(NameVariant)
            Mapping(VariantKey) = CanonicalName

        End If

    Next r

    Set LoadManualNameVariantMap = Mapping

End Function

Private Function CountNonAsciiCharacters( _
    ByVal InputText As String) As Long

    Dim i As Long

    For i = 1 To Len(InputText)

        If AscW(Mid(InputText, i, 1)) > 127 Then

            CountNonAsciiCharacters = _
                CountNonAsciiCharacters + 1

        End If

    Next i

End Function

Private Function CachedEntityVariantComesFirst( _
    ByVal FirstVariant As String, _
    ByVal FirstHasShareClass As Boolean, _
    ByVal FirstKeyLength As Long, _
    ByVal FirstTextLength As Long, _
    ByVal FirstAccentCount As Long, _
    ByVal SecondVariant As String, _
    ByVal SecondHasShareClass As Boolean, _
    ByVal SecondKeyLength As Long, _
    ByVal SecondTextLength As Long, _
    ByVal SecondAccentCount As Long) As Boolean

    If FirstHasShareClass <> SecondHasShareClass Then

        ' Prefer the issuer-level spelling over a security-level variant.
        CachedEntityVariantComesFirst = _
            Not FirstHasShareClass

        Exit Function

    End If

    If FirstKeyLength <> SecondKeyLength Then

        CachedEntityVariantComesFirst = _
            (FirstKeyLength > SecondKeyLength)

        Exit Function

    End If

    If FirstTextLength <> SecondTextLength Then

        CachedEntityVariantComesFirst = _
            (FirstTextLength > SecondTextLength)

        Exit Function

    End If

    If FirstAccentCount <> SecondAccentCount Then

        CachedEntityVariantComesFirst = _
            (FirstAccentCount > SecondAccentCount)

        Exit Function

    End If

    CachedEntityVariantComesFirst = _
        (StrComp( _
            FirstVariant, _
            SecondVariant, _
            vbTextCompare) < 0)

End Function

Private Sub SortEntityVariants( _
    ByRef Variants() As String)

    Dim HasShareClass() As Boolean
    Dim EntityKeyLength() As Long
    Dim TextLength() As Long
    Dim AccentCount() As Long
    Dim FirstIndex As Long
    Dim LastIndex As Long
    Dim i As Long
    Dim j As Long
    Dim TemporaryVariant As String
    Dim TemporaryHasShareClass As Boolean
    Dim TemporaryEntityKeyLength As Long
    Dim TemporaryTextLength As Long
    Dim TemporaryAccentCount As Long

    FirstIndex = LBound(Variants)
    LastIndex = UBound(Variants)

    ReDim HasShareClass(FirstIndex To LastIndex)
    ReDim EntityKeyLength(FirstIndex To LastIndex)
    ReDim TextLength(FirstIndex To LastIndex)
    ReDim AccentCount(FirstIndex To LastIndex)

    ' The former comparison recalculated all four properties for every pair.
    ' Cache them once, then retain the same comparison and swap sequence.
    For i = FirstIndex To LastIndex

        HasShareClass(i) = _
            (StrComp( _
                NormalizeExactNameKey(Variants(i)), _
                NormalizeExactNameKey( _
                    RemoveTrailingEntityShareClass(Variants(i))), _
                vbTextCompare) <> 0)
        EntityKeyLength(i) = _
            Len(NormalizeEntityKey(Variants(i)))
        TextLength(i) = Len(Variants(i))
        AccentCount(i) = _
            CountNonAsciiCharacters(Variants(i))

    Next i

    For i = FirstIndex To LastIndex - 1

        For j = i + 1 To LastIndex

            If CachedEntityVariantComesFirst( _
                    Variants(j), _
                    HasShareClass(j), _
                    EntityKeyLength(j), _
                    TextLength(j), _
                    AccentCount(j), _
                    Variants(i), _
                    HasShareClass(i), _
                    EntityKeyLength(i), _
                    TextLength(i), _
                    AccentCount(i)) Then

                TemporaryVariant = Variants(i)
                Variants(i) = Variants(j)
                Variants(j) = TemporaryVariant

                TemporaryHasShareClass = HasShareClass(i)
                HasShareClass(i) = HasShareClass(j)
                HasShareClass(j) = TemporaryHasShareClass

                TemporaryEntityKeyLength = EntityKeyLength(i)
                EntityKeyLength(i) = EntityKeyLength(j)
                EntityKeyLength(j) = TemporaryEntityKeyLength

                TemporaryTextLength = TextLength(i)
                TextLength(i) = TextLength(j)
                TextLength(j) = TemporaryTextLength

                TemporaryAccentCount = AccentCount(i)
                AccentCount(i) = AccentCount(j)
                AccentCount(j) = TemporaryAccentCount

            End If

        Next j

    Next i

End Sub

Private Function FindEntityPrefixCanonicalName( _
    ByVal EntityKey As String, _
    ByVal EntityByKey As Object) As String

    Dim ExistingKey As Variant
    Dim ExistingCanonicalName As String
    Dim CandidateName As String
    Dim IsAmbiguous As Boolean

    If EntityByKey Is Nothing Then Exit Function

    For Each ExistingKey In EntityByKey.Keys

        If IsLikelyEntityPrefixMatch( _
                EntityKey, _
                CStr(ExistingKey)) Then

            ExistingCanonicalName = _
                CStr(EntityByKey(ExistingKey))

            If CandidateName = "" Then

                CandidateName = ExistingCanonicalName

            ElseIf StrComp( _
                    CandidateName, _
                    ExistingCanonicalName, _
                    vbTextCompare) <> 0 Then

                IsAmbiguous = True

            End If

        End If

    Next ExistingKey

    If Not IsAmbiguous Then

        FindEntityPrefixCanonicalName = CandidateName

    End If

End Function

Private Function BuildCanonicalEntityNameMap( _
    ByVal GeographyEntries As Object) As Object

    Dim CanonicalByVariant As Object
    Dim EntityByKey As Object
    Dim ManualVariantMap As Object
    Dim VariantSet As Object
    Dim Entry As Object

    Dim GeographyKey As Variant
    Dim ManualKey As Variant
    Dim Parts As Variant
    Dim Part As Variant
    Dim VariantKeys As Variant
    Dim Variants() As String

    Dim NameVariant As String
    Dim CanonicalName As String
    Dim VariantKey As String
    Dim EntityKey As String
    Dim CanonicalEntityKey As String
    Dim i As Long

    Set CanonicalByVariant = NewExactNameMap()
    Set EntityByKey = NewExactNameMap()
    Set ManualVariantMap = LoadManualNameVariantMap()


    Set VariantSet = CreateObject("Scripting.Dictionary")
    VariantSet.CompareMode = vbBinaryCompare

    For Each ManualKey In ManualVariantMap.Keys

        CanonicalName = _
            StandardizeEntityDisplayName( _
                CStr(ManualVariantMap(ManualKey)))
        CanonicalByVariant(CStr(ManualKey)) = CanonicalName

        ' Register both sides of a maintained override. This lets another
        ' spelling that normalizes like the maintained variant inherit the
        ' same canonical name, even when the canonical display name itself
        ' has a substantially different form.
        EntityKey = NormalizeEntityKey(CStr(ManualKey))

        If EntityKey <> "" Then

            EntityByKey(EntityKey) = CanonicalName

        End If

        CanonicalEntityKey = NormalizeEntityKey(CanonicalName)

        If CanonicalEntityKey <> "" Then

            EntityByKey(CanonicalEntityKey) = CanonicalName

        End If

    Next ManualKey


    If Not GeographyEntries Is Nothing Then

        For Each GeographyKey In GeographyEntries.Keys

            Set Entry = GeographyEntries(GeographyKey)
            Parts = Split(CStr(Entry("Variants")), ";")

            For Each Part In Parts

                NameVariant = Trim(CStr(Part))

                If NameVariant <> "" And _
                   Not VariantSet.Exists(NameVariant) Then

                    VariantSet.Add NameVariant, NameVariant

                End If

            Next Part

        Next GeographyKey

    End If


    If VariantSet.Count > 0 Then

        VariantKeys = VariantSet.Keys
        ReDim Variants(0 To VariantSet.Count - 1)

        For i = LBound(VariantKeys) To UBound(VariantKeys)

            Variants(i) = CStr(VariantKeys(i))

        Next i

        SortEntityVariants Variants


        For i = LBound(Variants) To UBound(Variants)

            NameVariant = Variants(i)
            VariantKey = NormalizeExactNameKey(NameVariant)

            If CanonicalByVariant.Exists(VariantKey) Then

                CanonicalName = _
                    CStr(CanonicalByVariant(VariantKey))

            Else

                EntityKey = NormalizeEntityKey(NameVariant)

                If EntityKey = "" Then

                    CanonicalName = _
                        StandardizeEntityDisplayName(NameVariant)

                ElseIf EntityByKey.Exists(EntityKey) Then

                    CanonicalName = _
                        CStr(EntityByKey(EntityKey))

                Else

                    CanonicalName = _
                        FindEntityPrefixCanonicalName( _
                            EntityKey, _
                            EntityByKey)

                    If CanonicalName = "" Then

                        CanonicalName = _
                            StandardizeEntityDisplayName(NameVariant)

                    End If

                    EntityByKey.Add EntityKey, CanonicalName

                End If

                CanonicalByVariant(VariantKey) = CanonicalName

            End If

        Next i


    End If

    Set BuildCanonicalEntityNameMap = CanonicalByVariant

End Function

Private Function ResolveCanonicalEntityName( _
    ByVal RawName As String, _
    ByVal CanonicalNameMap As Object) As String

    Dim VariantKey As String

    If Left( _
            RawName, _
            Len(UNKNOWN_UNDERLYING_PREFIX)) = _
       UNKNOWN_UNDERLYING_PREFIX Then

        ResolveCanonicalEntityName = RawName

        Exit Function

    End If

    ResolveCanonicalEntityName = _
        StandardizeEntityDisplayName(RawName)

    If CanonicalNameMap Is Nothing Then Exit Function

    VariantKey = NormalizeExactNameKey(RawName)

    If CanonicalNameMap.Exists(VariantKey) Then

        ResolveCanonicalEntityName = _
            CStr(CanonicalNameMap(VariantKey))

    End If

End Function

Private Function CanonicalizeGeographyEntries( _
    ByVal RawEntries As Object, _
    ByVal CanonicalNameMap As Object) As Object

    Dim CanonicalEntries As Object
    Dim Entry As Object
    Dim RawKey As Variant
    Dim Parts As Variant
    Dim Part As Variant
    Dim RawName As String
    Dim CanonicalName As String
    Dim NameVariant As String

    Set CanonicalEntries = NewExactNameMap()

    If Not RawEntries Is Nothing Then

        For Each RawKey In RawEntries.Keys

            Set Entry = RawEntries(RawKey)
            RawName = CStr(Entry("Name"))

            CanonicalName = _
                ResolveCanonicalEntityName( _
                    RawName, _
                    CanonicalNameMap)

            NameVariant = ""
            Parts = Split(CStr(Entry("Variants")), ";")

            For Each Part In Parts

                If Trim(CStr(Part)) <> "" Then

                    NameVariant = Trim(CStr(Part))

                    Exit For

                End If

            Next Part

            If NameVariant <> "" Then

                AddGeographyEntryWithCandidates _
                    CanonicalEntries, _
                    CanonicalName, _
                    Entry, _
                    NameVariant

            End If

        Next RawKey

    End If


    Set CanonicalizeGeographyEntries = CanonicalEntries

End Function

Private Sub UpdateNameVariantsWorksheet( _
    ByVal RawEntries As Object, _
    ByVal CanonicalNameMap As Object)

    Dim wsVariants As Worksheet
    Dim ExistingRows As Object
    Dim PendingRows As Object
    Dim ExistingData As Variant
    Dim ExistingUpdateData() As Variant
    Dim PendingOutputData() As Variant
    Dim PendingOrder As Collection
    Dim Entry As Object
    Dim RawKey As Variant
    Dim Parts As Variant
    Dim Part As Variant
    Dim PendingName As Variant
    Dim PendingData As Variant
    Dim NameVariant As String
    Dim CanonicalName As String
    Dim LastRow As Long
    Dim ExistingRowCount As Long
    Dim PendingRowCount As Long
    Dim PendingRow As Long
    Dim r As Long
    Dim ColNo As Long

    Set wsVariants = EnsureNameVariantsWorksheet()

    If wsVariants Is Nothing Then Exit Sub

    Set ExistingRows = CreateObject("Scripting.Dictionary")
    ExistingRows.CompareMode = vbBinaryCompare
    Set PendingRows = CreateObject("Scripting.Dictionary")
    PendingRows.CompareMode = vbBinaryCompare
    Set PendingOrder = New Collection

    LastRow = GetNameVariantsLastRow(wsVariants)
    ExistingRowCount = LastRow - 1

    If ExistingRowCount > 0 Then

        ExistingData = _
            wsVariants.Range( _
                wsVariants.Cells(2, 1), _
                wsVariants.Cells(LastRow, 3)).Value2

    End If

    For r = 1 To ExistingRowCount

        NameVariant = SafeText(ExistingData(r, 1))

        If NameVariant <> "" And _
           Not ExistingRows.Exists(NameVariant) Then

            ExistingRows.Add NameVariant, r

        End If

    Next r

    If Not RawEntries Is Nothing Then

        For Each RawKey In RawEntries.Keys

            Set Entry = RawEntries(RawKey)
            Parts = Split(CStr(Entry("Variants")), ";")

            For Each Part In Parts

                NameVariant = Trim(CStr(Part))

                If NameVariant <> "" Then

                    CanonicalName = _
                        ResolveCanonicalEntityName( _
                            NameVariant, _
                            CanonicalNameMap)

                    If ExistingRows.Exists(NameVariant) Then

                        r = CLng(ExistingRows(NameVariant))
                        ExistingData(r, 2) = _
                            CanonicalName
                        ExistingData(r, 3) = _
                            NormalizeEntityKey(NameVariant)

                    ElseIf PendingRows.Exists(NameVariant) Then

                        PendingRows(NameVariant) = _
                            Array( _
                                NameVariant, _
                                CanonicalName, _
                                NormalizeEntityKey(NameVariant))

                    Else

                        PendingRows.Add _
                            NameVariant, _
                            Array( _
                                NameVariant, _
                                CanonicalName, _
                                NormalizeEntityKey(NameVariant))
                        PendingOrder.Add NameVariant

                    End If

                End If

            Next Part

        Next RawKey

    End If

    If ExistingRowCount > 0 Then

        ReDim ExistingUpdateData( _
            1 To ExistingRowCount, _
            1 To 2)

        For r = 1 To ExistingRowCount

            For ColNo = 1 To 2

                ExistingUpdateData(r, ColNo) = _
                    ExistingData(r, ColNo + 1)

            Next ColNo

        Next r

        wsVariants.Range( _
            wsVariants.Cells(2, 2), _
            wsVariants.Cells(ExistingRowCount + 1, 3)).Value2 = _
                ExistingUpdateData

    End If

    PendingRowCount = PendingOrder.Count

    If PendingRowCount > 0 Then

        ReDim PendingOutputData(1 To PendingRowCount, 1 To 3)

        For Each PendingName In PendingOrder

            PendingRow = PendingRow + 1
            PendingData = PendingRows(CStr(PendingName))

            For ColNo = 1 To 3

                PendingOutputData(PendingRow, ColNo) = _
                    PendingData(ColNo - 1)

            Next ColNo

        Next PendingName

        wsVariants.Range( _
            wsVariants.Cells(ExistingRowCount + 2, 1), _
            wsVariants.Cells( _
                ExistingRowCount + PendingRowCount + 1, _
                3)).Value2 = _
                    PendingOutputData

    End If

    LastRow = ExistingRowCount + PendingRowCount + 1

    If LastRow >= 2 Then

        wsVariants.Range( _
            wsVariants.Cells(1, 1), _
            wsVariants.Cells(LastRow, 4)).Sort _
                Key1:=wsVariants.Cells(2, 1), _
                Order1:=xlAscending, _
                Header:=xlYes

    End If

    wsVariants.Columns("A:D").AutoFit

End Sub

Private Sub AddCertificateGeographyEntries( _
    ByVal GeographyEntries As Object, _
    ByVal CertificateMap As Object, _
    ByVal UnderlyingReferenceIsinMap As Object, _
    ByVal UnderlyingAssetClassMap As Object, _
    ByVal FundMap As Object, _
    ByVal CertificateISIN As String, _
    ByVal CertificateSecurityName As String, _
    ByVal ObservationSet As Object)

    Dim Underlyings As Collection
    Dim Underlying As Variant
    Dim Component As Object
    Dim UnderlyingName As String
    Dim GeographyName As String
    Dim UnderlyingKey As String
    Dim UnderlyingISIN As String
    Dim UnderlyingAssetClass As String
    Dim ParentCompany As String
    Dim ExposureType As String
    Dim IsinRelationship As String
    Dim IsinPriority As Long

    If GeographyEntries Is Nothing Then Exit Sub
    If CertificateMap Is Nothing Then Exit Sub

    CertificateISIN = UCase(Trim(CertificateISIN))

    If CertificateISIN <> "" Then

        If CertificateMap.Exists(CertificateISIN) Then

            Set Underlyings = CertificateMap(CertificateISIN)

        End If

    End If

    If Underlyings Is Nothing Then

        CertificateSecurityName = Trim(CertificateSecurityName)

        If CertificateSecurityName <> "" Then

            If CertificateMap.Exists(CertificateSecurityName) Then

                Set Underlyings = _
                    CertificateMap(CertificateSecurityName)

            End If

        End If

    End If

    If Underlyings Is Nothing Then Exit Sub

    For Each Underlying In Underlyings

        Set Component = Underlying
        UnderlyingName = CStr(Component("Name"))

        If Left( _
                UnderlyingName, _
                Len(MISSING_CERTIFICATE_RIC_PREFIX)) <> _
           MISSING_CERTIFICATE_RIC_PREFIX And _
           Left( _
                UnderlyingName, _
                Len(UNKNOWN_UNDERLYING_PREFIX)) <> _
           UNKNOWN_UNDERLYING_PREFIX Then

            GeographyName = UnderlyingName
            UnderlyingKey = NormalizeExactNameKey(UnderlyingName)
            UnderlyingISIN = _
                CStr(Component("ReferenceISIN"))
            UnderlyingAssetClass = _
                CStr(Component("AssetClass"))

            If UnderlyingISIN = "" Then

                If Not UnderlyingReferenceIsinMap Is Nothing Then

                    If UnderlyingReferenceIsinMap.Exists(UnderlyingKey) Then

                        UnderlyingISIN = _
                            CStr( _
                                UnderlyingReferenceIsinMap(UnderlyingKey))

                    End If

                End If

            End If

            If UnderlyingAssetClass = "" Then

                If Not UnderlyingAssetClassMap Is Nothing Then

                    If UnderlyingAssetClassMap.Exists(UnderlyingKey) Then

                        UnderlyingAssetClass = _
                            CStr( _
                                UnderlyingAssetClassMap(UnderlyingKey))

                    End If

                End If

            End If

            ExposureType = _
                "Certificate underlying"

            If UnderlyingAssetClass <> "" Then

                ExposureType = _
                    ExposureType & " - " & UnderlyingAssetClass

            End If

            If InStr( _
                    1, _
                    UCase(UnderlyingAssetClass), _
                    "INDEX", _
                    vbTextCompare) > 0 Then

                IsinRelationship = "Index identifier"
                IsinPriority = GEO_ISIN_PRIORITY_UNDERLYING

            ElseIf (StrComp( _
                        UnderlyingAssetClass, _
                        "Fund", _
                        vbTextCompare) = 0 Or _
                    StrComp( _
                        UnderlyingAssetClass, _
                        "ETF", _
                        vbTextCompare) = 0) Then

                ParentCompany = _
                    ResolveExactName( _
                        UnderlyingName, _
                        FundMap)

                If ParentCompany <> "" Then

                    GeographyName = ParentCompany
                    ExposureType = _
                        "Certificate underlying fund parent company"

                End If

                IsinRelationship = "Managed underlying fund"
                IsinPriority = GEO_ISIN_PRIORITY_MANAGED_FUND

            Else

                IsinRelationship = "Underlying security"
                IsinPriority = GEO_ISIN_PRIORITY_UNDERLYING

            End If

            AddGeographyLookupEntry _
                GeographyEntries, _
                GeographyName, _
                ExposureType, _
                UnderlyingISIN, _
                IsinRelationship, _
                IsinPriority, _
                ObservationSet

        End If

    Next Underlying

End Sub

Private Function ReadWorksheetColumnValues( _
    ByVal ws As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal LastRow As Long, _
    ByVal ColumnNo As Long) As Variant

    Dim SingleRowData() As Variant

    If ws Is Nothing Then Exit Function
    If FirstRow < 1 Or LastRow < FirstRow Or _
       ColumnNo < 1 Then Exit Function

    If FirstRow = LastRow Then

        ReDim SingleRowData(1 To 1, 1 To 1)
        SingleRowData(1, 1) = _
            ws.Cells(FirstRow, ColumnNo).Value2

        ReadWorksheetColumnValues = SingleRowData

    Else

        ReadWorksheetColumnValues = _
            ws.Range( _
                ws.Cells(FirstRow, ColumnNo), _
                ws.Cells(LastRow, ColumnNo)).Value2

    End If

End Function

Private Function GetCompaniesDataTable( _
    ByVal PreferredWorkbook As Workbook) As ListObject

    Dim wsCompanies As Worksheet
    Dim CandidateTable As ListObject

    Set GetCompaniesDataTable = _
        GetReferenceDataTable( _
            COMPANIES_SHEET, _
            COMPANIES_TABLE, _
            PreferredWorkbook)

    If Not GetCompaniesDataTable Is Nothing Then Exit Function

    Set wsCompanies = _
        GetReferenceDataWorksheet( _
            COMPANIES_SHEET, _
            PreferredWorkbook)

    If wsCompanies Is Nothing Then Exit Function

    For Each CandidateTable In wsCompanies.ListObjects

        If GetTableColumnIndex(CandidateTable, "Name") > 0 And _
           GetTableColumnIndex( _
                CandidateTable, _
                "Geography Final") > 0 And _
           GetTableColumnIndex( _
                CandidateTable, _
                "Sector Final") > 0 Then

            Set GetCompaniesDataTable = CandidateTable

            Exit Function

        End If

    Next CandidateTable

End Function

Private Sub LoadCompaniesLookup( _
    ByRef CompaniesByName As Object, _
    ByRef CompaniesByVariant As Object, _
    ByRef CompaniesReady As Boolean)

    Dim DataTable As ListObject
    Dim TableData As Variant
    Dim Entry As Object
    Dim Parts As Variant
    Dim Part As Variant
    Dim CompanyName As String
    Dim NameVariants As String
    Dim LookupKey As String
    Dim r As Long
    Dim NameColumn As Long
    Dim NameVariantsColumn As Long
    Dim ExposureTypeColumn As Long
    Dim ReferenceISINColumn As Long
    Dim IsinRelationshipColumn As Long
    Dim GeographyFinalColumn As Long
    Dim SectorFinalColumn As Long

    Set CompaniesByName = NewExactNameMap()
    Set CompaniesByVariant = NewExactNameMap()
    CompaniesReady = False

    Set DataTable = GetCompaniesDataTable(ThisWorkbook)

    If DataTable Is Nothing Then Exit Sub

    NameColumn = GetTableColumnIndex(DataTable, "Name")
    NameVariantsColumn = _
        GetTableColumnIndex(DataTable, "Name Variants")
    ExposureTypeColumn = _
        GetTableColumnIndex(DataTable, "Exposure Type")
    ReferenceISINColumn = _
        GetTableColumnIndex(DataTable, "Reference ISIN")
    IsinRelationshipColumn = _
        GetTableColumnIndex(DataTable, "ISIN Relationship")
    GeographyFinalColumn = _
        GetTableColumnIndex(DataTable, "Geography Final")
    SectorFinalColumn = _
        GetTableColumnIndex(DataTable, "Sector Final")

    If NameColumn = 0 Or NameVariantsColumn = 0 Or _
       ExposureTypeColumn = 0 Or ReferenceISINColumn = 0 Or _
       IsinRelationshipColumn = 0 Or _
       GeographyFinalColumn = 0 Or _
       SectorFinalColumn = 0 Then Exit Sub

    CompaniesReady = True

    If DataTable.DataBodyRange Is Nothing Then Exit Sub

    TableData = DataTable.DataBodyRange.Value2

    For r = 1 To UBound(TableData, 1)

        CompanyName = _
            SafeText( _
                TableData(r, NameColumn))

        If CompanyName <> "" Then

            Set Entry = CreateObject("Scripting.Dictionary")
            Entry.CompareMode = vbTextCompare

            Entry.Add "Name", CompanyName
            Entry.Add "NameVariants", _
                SafeText( _
                    TableData(r, NameVariantsColumn))
            Entry.Add "ExposureType", _
                SafeText( _
                    TableData(r, ExposureTypeColumn))
            Entry.Add "ReferenceISIN", _
                SafeText( _
                    TableData(r, ReferenceISINColumn))
            Entry.Add "IsinRelationship", _
                SafeText( _
                    TableData(r, IsinRelationshipColumn))
            Entry.Add "GeographyFinal", _
                TableData(r, GeographyFinalColumn)
            Entry.Add "SectorFinal", _
                TableData(r, SectorFinalColumn)

            LookupKey = NormalizeExactNameKey(CompanyName)

            If LookupKey <> "" Then

                If Not CompaniesByName.Exists(LookupKey) Then

                    CompaniesByName.Add LookupKey, Entry

                End If

                If Not CompaniesByVariant.Exists(LookupKey) Then

                    CompaniesByVariant.Add LookupKey, Entry

                End If

            End If

            NameVariants = CStr(Entry("NameVariants"))
            Parts = Split(NameVariants, ";")

            For Each Part In Parts

                LookupKey = _
                    NormalizeExactNameKey(Trim(CStr(Part)))

                If LookupKey <> "" Then

                    If Not CompaniesByVariant.Exists(LookupKey) Then

                        CompaniesByVariant.Add LookupKey, Entry

                    End If

                End If

            Next Part

        End If

    Next r

End Sub

Private Function ResolveCompanyEntry( _
    ByVal ExposureName As String, _
    ByVal CompaniesByName As Object, _
    ByVal CompaniesByVariant As Object) As Object

    Dim LookupKey As String

    LookupKey = NormalizeExactNameKey(ExposureName)

    If LookupKey = "" Then Exit Function

    If Not CompaniesByName Is Nothing Then

        If CompaniesByName.Exists(LookupKey) Then

            Set ResolveCompanyEntry = CompaniesByName(LookupKey)

            Exit Function

        End If

    End If

    If Not CompaniesByVariant Is Nothing Then

        If CompaniesByVariant.Exists(LookupKey) Then

            Set ResolveCompanyEntry = CompaniesByVariant(LookupKey)

        End If

    End If

End Function

Private Function IsUsableCompanyAttribute( _
    ByVal RawValue As Variant) As Boolean

    Dim AttributeText As String
    Dim FirstCharacter As String

    If IsError(RawValue) Then Exit Function
    If IsNull(RawValue) Then Exit Function
    If IsEmpty(RawValue) Then Exit Function

    AttributeText = Trim(CStr(RawValue))

    If AttributeText = "" Then Exit Function

    FirstCharacter = Left(AttributeText, 1)

    If FirstCharacter = "#" Or FirstCharacter = "*" Then Exit Function

    If AscW(FirstCharacter) = &H2605 Or _
       AscW(FirstCharacter) = &H2606 Then Exit Function

    If StrComp(AttributeText, "N/A", vbTextCompare) = 0 Then Exit Function
    If StrComp(AttributeText, "NA", vbTextCompare) = 0 Then Exit Function

    IsUsableCompanyAttribute = True

End Function

Private Function CompanyDimensionValue( _
    ByVal CompanyEntry As Object, _
    ByVal DimensionName As String) As String

    Dim RawValue As Variant

    CompanyDimensionValue = OTHER_RISK_DIMENSION

    If CompanyEntry Is Nothing Then Exit Function

    If StrComp( _
            DimensionName, _
            "Country", _
            vbTextCompare) = 0 Then

        RawValue = CompanyEntry("GeographyFinal")

    Else

        RawValue = CompanyEntry("SectorFinal")

    End If

    If IsUsableCompanyAttribute(RawValue) Then

        CompanyDimensionValue = Trim(CStr(RawValue))

    End If

End Function

Private Function DelimitedTextContainsToken( _
    ByVal ExistingText As String, _
    ByVal AdditionalExistingText As String, _
    ByVal RequiredTokenKey As String) As Boolean

    Dim SourceText As Variant
    Dim Parts As Variant
    Dim Part As Variant

    If RequiredTokenKey = "" Then

        DelimitedTextContainsToken = True

        Exit Function

    End If

    For Each SourceText In _
        Array(ExistingText, AdditionalExistingText)

        Parts = Split(CStr(SourceText), ";")

        For Each Part In Parts

            If StrComp( _
                    NormalizeExactNameKey( _
                        Trim(CStr(Part))), _
                    RequiredTokenKey, _
                    vbBinaryCompare) = 0 Then

                DelimitedTextContainsToken = True

                Exit Function

            End If

        Next Part

    Next SourceText

End Function

Private Function DelimitedListContainsAll( _
    ByVal ExistingText As String, _
    ByVal RequiredText As String, _
    Optional ByVal AdditionalExistingText As String = "") As Boolean

    Dim RequiredParts As Variant
    Dim RequiredPart As Variant
    Dim RequiredTokenKey As String

    RequiredParts = Split(RequiredText, ";")

    For Each RequiredPart In RequiredParts

        RequiredTokenKey = _
            NormalizeExactNameKey( _
                Trim(CStr(RequiredPart)))

        If RequiredTokenKey <> "" Then

            If Not DelimitedTextContainsToken( _
                    ExistingText, _
                    AdditionalExistingText, _
                    RequiredTokenKey) Then Exit Function

        End If

    Next RequiredPart

    DelimitedListContainsAll = True

End Function

Private Function MergeDelimitedText( _
    ByVal ExistingText As String, _
    ByVal NewText As String) As String

    Dim Seen As Object
    Dim SourceText As Variant
    Dim Parts As Variant
    Dim Part As Variant
    Dim TokenText As String
    Dim TokenKey As String

    Set Seen = NewExactNameMap()

    For Each SourceText In Array(ExistingText, NewText)

        Parts = Split(CStr(SourceText), ";")

        For Each Part In Parts

            TokenText = Trim(CStr(Part))
            TokenKey = NormalizeExactNameKey(TokenText)

            If TokenKey <> "" Then

                If Not Seen.Exists(TokenKey) Then

                    Seen.Add TokenKey, True

                    If MergeDelimitedText <> "" Then

                        MergeDelimitedText = _
                            MergeDelimitedText & "; "

                    End If

                    MergeDelimitedText = _
                        MergeDelimitedText & TokenText

                End If

            End If

        Next Part

    Next SourceText

End Function

Private Function CurrentGeographyRelationship( _
    ByVal GeographyEntry As Object) As String

    If GeographyEntry Is Nothing Then Exit Function

    If CStr(GeographyEntry("ReferenceISIN")) = "" Then

        CurrentGeographyRelationship = _
            "No representative ISIN available"

    Else

        CurrentGeographyRelationship = _
            CStr(GeographyEntry("IsinRelationship"))

    End If

End Function

Private Function GeographyRelationshipPriority( _
    ByVal IsinRelationship As String) As Long

    Dim RelationshipKey As String

    RelationshipKey = _
        NormalizeExactNameKey(IsinRelationship)

    Select Case RelationshipKey

        Case "ISSUED SECURITY"

            GeographyRelationshipPriority = _
                GEO_ISIN_PRIORITY_ISSUED_SECURITY

        Case "UNDERLYING SECURITY", _
             "INDEX IDENTIFIER"

            GeographyRelationshipPriority = _
                GEO_ISIN_PRIORITY_UNDERLYING

        Case "MANAGED FUND", _
             "MANAGED UNDERLYING FUND"

            GeographyRelationshipPriority = _
                GEO_ISIN_PRIORITY_MANAGED_FUND

        Case Else

            GeographyRelationshipPriority = 999&

    End Select

End Function

Private Function ReferenceIsinNeedsRevision( _
    ByVal GeographyEntry As Object, _
    ByVal CompanyEntry As Object) As Boolean

    Dim IsinCandidates As Object
    Dim ExistingCandidate As Object
    Dim ExistingISIN As String
    Dim CurrentISIN As String
    Dim ExistingRelationship As String
    Dim CandidateRelationships As String
    Dim CandidateCount As Long
    Dim BestPriority As Long
    Dim ObservedExistingPriority As Long
    Dim StoredPriority As Long

    If GeographyEntry Is Nothing Then Exit Function
    If CompanyEntry Is Nothing Then Exit Function

    ExistingISIN = _
        UCase(Trim(CStr(CompanyEntry("ReferenceISIN"))))
    CurrentISIN = _
        UCase(Trim(CStr(GeographyEntry("ReferenceISIN"))))

    If GeographyEntry.Exists("IsinCandidates") Then

        Set IsinCandidates = _
            GeographyEntry("IsinCandidates")

        CandidateCount = IsinCandidates.Count

    End If

    If ExistingISIN = "" Then

        ReferenceIsinNeedsRevision = (CurrentISIN <> "")

        Exit Function

    End If

    ' No current candidate is not enough evidence to overwrite a maintained
    ' historical reference. Keep it until an alternative is actually present.
    If CandidateCount = 0 Then Exit Function

    If Not IsinCandidates.Exists(ExistingISIN) Then

        ' The maintained ISIN is no longer among the currently valid
        ' candidates, while at least one replacement candidate is available.
        ReferenceIsinNeedsRevision = True

        Exit Function

    End If

    Set ExistingCandidate = _
        IsinCandidates(ExistingISIN)

    ExistingRelationship = _
        Trim(CStr(CompanyEntry("IsinRelationship")))

    BestPriority = CLng(GeographyEntry("IsinPriority"))
    ObservedExistingPriority = _
        CLng(ExistingCandidate("Priority"))
    StoredPriority = _
        GeographyRelationshipPriority( _
            ExistingRelationship)

    If StoredPriority = 999& Then

        StoredPriority = ObservedExistingPriority

    End If

    If BestPriority < StoredPriority Then

        ' A genuinely higher-priority relationship has become available.
        ReferenceIsinNeedsRevision = True

        Exit Function

    End If

    ' Seeing the maintained ISIN only in a lower-priority role this week does
    ' not invalidate its maintained higher-priority relationship.
    If ObservedExistingPriority > StoredPriority Then Exit Function

    If ExistingCandidate.Exists("Relationships") Then

        CandidateRelationships = _
            CStr(ExistingCandidate("Relationships"))

    Else

        CandidateRelationships = _
            CStr(ExistingCandidate("Relationship"))

    End If

    If ExistingRelationship = "" Then

        ReferenceIsinNeedsRevision = _
            (Trim(CandidateRelationships) <> "")

        Exit Function

    End If

    If Not DelimitedListContainsAll( _
            CandidateRelationships, _
            ExistingRelationship) Then

        ' The same ISIN is now observed under a different relationship.
        ReferenceIsinNeedsRevision = True

    End If

End Function

Private Sub GetReferenceRevisionValues( _
    ByVal GeographyEntry As Object, _
    ByVal CompanyEntry As Object, _
    ByRef RevisedISIN As String, _
    ByRef RevisedRelationship As String)

    Dim IsinCandidates As Object
    Dim ExistingCandidate As Object
    Dim ExistingISIN As String
    Dim BestPriority As Long

    RevisedISIN = _
        UCase(Trim(CStr(GeographyEntry("ReferenceISIN"))))
    RevisedRelationship = _
        CurrentGeographyRelationship(GeographyEntry)

    If CompanyEntry Is Nothing Then Exit Sub

    ExistingISIN = _
        UCase(Trim(CStr(CompanyEntry("ReferenceISIN"))))

    If ExistingISIN = "" Then Exit Sub

    If Not GeographyEntry.Exists("IsinCandidates") Then Exit Sub

    Set IsinCandidates = _
        GeographyEntry("IsinCandidates")

    If Not IsinCandidates.Exists(ExistingISIN) Then Exit Sub

    Set ExistingCandidate = _
        IsinCandidates(ExistingISIN)

    BestPriority = CLng(GeographyEntry("IsinPriority"))

    ' When the maintained ISIN itself is still one of the best candidates,
    ' keep it and revise only its relationship. This avoids replacing it with
    ' another same-priority ISIN selected by tie-breaking.
    If CLng(ExistingCandidate("Priority")) = BestPriority Then

        RevisedISIN = ExistingISIN
        RevisedRelationship = _
            CStr(ExistingCandidate("Relationship"))

    End If

End Sub

Private Function CompanyEntryNeedsLookup( _
    ByVal GeographyEntry As Object, _
    ByVal CompanyEntry As Object) As Boolean

    Dim CurrentName As String
    Dim CurrentVariants As String
    Dim CurrentTypes As String
    Dim ExistingNames As String

    If GeographyEntry Is Nothing Then Exit Function

    If CompanyEntry Is Nothing Then

        CompanyEntryNeedsLookup = True

        Exit Function

    End If

    CurrentName = CStr(GeographyEntry("Name"))
    CurrentVariants = CStr(GeographyEntry("Variants"))
    CurrentTypes = CStr(GeographyEntry("ExposureType"))
    ExistingNames = _
        CStr(CompanyEntry("NameVariants"))

    If Not DelimitedListContainsAll( _
            ExistingNames, _
            CurrentName & "; " & CurrentVariants, _
            CStr(CompanyEntry("Name"))) Then

        CompanyEntryNeedsLookup = True

        Exit Function

    End If

    If Not DelimitedListContainsAll( _
            CStr(CompanyEntry("ExposureType")), _
            CurrentTypes) Then

        CompanyEntryNeedsLookup = True

        Exit Function

    End If

    If ReferenceIsinNeedsRevision( _
            GeographyEntry, _
            CompanyEntry) Then

        CompanyEntryNeedsLookup = True

    End If

End Function

Private Sub DeleteLegacyGeographyLookupWorksheet()

    Dim wsLegacy As Worksheet
    Dim PreviousDisplayAlerts As Boolean

    Set wsLegacy = _
        GetOptionalWorksheet(LEGACY_GEOGRAPHY_LOOKUP_SHEET)

    If wsLegacy Is Nothing Then Exit Sub

    PreviousDisplayAlerts = Application.DisplayAlerts

    On Error Resume Next

    Application.DisplayAlerts = False
    wsLegacy.Delete
    Application.DisplayAlerts = PreviousDisplayAlerts

    On Error GoTo 0

End Sub

Private Sub WriteNewGeoSecLookupWorksheet( _
    ByVal GeographyEntries As Object, _
    ByVal CompaniesByName As Object, _
    ByVal CompaniesByVariant As Object)

    Dim wsLookup As Worksheet
    Dim GeographyEntry As Object
    Dim CompanyEntry As Object
    Dim PendingKeys As Collection
    Dim EntryKey As Variant
    Dim Output() As Variant
    Dim CurrentISIN As String
    Dim CurrentRelationship As String
    Dim OutputCount As Long
    Dim OutputRow As Long
    Dim LastRow As Long
    Set wsLookup = _
        CreateOrReplaceSheet(GEO_SEC_LOOKUP_SHEET)

    If wsLookup Is Nothing Then Exit Sub

    wsLookup.Cells(1, 1).Value = "Name"
    wsLookup.Cells(1, 2).Value = "Name Variants"
    wsLookup.Cells(1, 3).Value = "Exposure Type"
    wsLookup.Cells(1, 4).Value = "Reference ISIN"
    wsLookup.Cells(1, 5).Value = "ISIN Relationship"
    wsLookup.Cells(1, 6).Value = "Country of Risk"
    wsLookup.Cells(1, 7).Value = "Fallback Geography"
    wsLookup.Cells(1, 8).Value = "Sector"
    wsLookup.Cells(1, 9).Value = "Fallback Sector"

    LastRow = 1
    Set PendingKeys = New Collection


    If Not GeographyEntries Is Nothing Then

        For Each EntryKey In GeographyEntries.Keys

            Set GeographyEntry = GeographyEntries(EntryKey)
            Set CompanyEntry = _
                ResolveCompanyEntry( _
                    CStr(GeographyEntry("Name")), _
                    CompaniesByName, _
                    CompaniesByVariant)

            If CompanyEntryNeedsLookup( _
                    GeographyEntry, _
                    CompanyEntry) Then

                PendingKeys.Add CStr(EntryKey)

            End If

        Next EntryKey

        OutputCount = PendingKeys.Count


        If OutputCount > 0 Then

            ReDim Output(1 To OutputCount, 1 To 9)
            OutputRow = 0

            For Each EntryKey In PendingKeys

                Set GeographyEntry = GeographyEntries(EntryKey)
                Set CompanyEntry = _
                    ResolveCompanyEntry( _
                        CStr(GeographyEntry("Name")), _
                        CompaniesByName, _
                        CompaniesByVariant)

                OutputRow = OutputRow + 1
                CurrentISIN = _
                    CStr(GeographyEntry("ReferenceISIN"))
                CurrentRelationship = _
                    CurrentGeographyRelationship( _
                        GeographyEntry)

                If CompanyEntry Is Nothing Then

                    Output(OutputRow, 1) = _
                        CStr(GeographyEntry("Name"))
                    Output(OutputRow, 2) = _
                        CStr(GeographyEntry("Variants"))
                    Output(OutputRow, 3) = _
                        CStr(GeographyEntry("ExposureType"))
                    Output(OutputRow, 4) = CurrentISIN
                    Output(OutputRow, 5) = CurrentRelationship

                Else

                    Output(OutputRow, 1) = _
                        CStr(CompanyEntry("Name"))
                    Output(OutputRow, 2) = _
                        MergeDelimitedText( _
                            CStr( _
                                CompanyEntry("NameVariants")), _
                            CStr( _
                                GeographyEntry("Variants")) & _
                            "; " & _
                            CStr( _
                                GeographyEntry("Name")))
                    Output(OutputRow, 3) = _
                        MergeDelimitedText( _
                            CStr( _
                                CompanyEntry("ExposureType")), _
                            CStr( _
                                GeographyEntry("ExposureType")))

                    If ReferenceIsinNeedsRevision( _
                            GeographyEntry, _
                            CompanyEntry) Then

                        GetReferenceRevisionValues _
                            GeographyEntry, _
                            CompanyEntry, _
                            CurrentISIN, _
                            CurrentRelationship

                        Output(OutputRow, 4) = CurrentISIN
                        Output(OutputRow, 5) = _
                            CurrentRelationship

                    Else

                        Output(OutputRow, 4) = _
                            CStr( _
                                CompanyEntry("ReferenceISIN"))
                        Output(OutputRow, 5) = _
                            CStr( _
                                CompanyEntry("IsinRelationship"))

                    End If

                End If

            Next EntryKey

            LastRow = OutputCount + 1

            wsLookup.Range( _
                wsLookup.Cells(2, 1), _
                wsLookup.Cells(LastRow, 9)).Value = Output

            wsLookup.Range( _
                wsLookup.Cells(1, 1), _
                wsLookup.Cells(LastRow, 9)).Sort _
                    Key1:=wsLookup.Cells(2, 1), _
                    Order1:=xlAscending, _
                    Header:=xlYes

            ' F:I are a fresh Bloomberg / manual lookup area. Never copy
            ' values from Companies into this incremental revision sheet.
            wsLookup.Range( _
                wsLookup.Cells(2, 6), _
                wsLookup.Cells(LastRow, 9)).ClearContents

            wsLookup.Range( _
                wsLookup.Cells(2, 6), _
                wsLookup.Cells(LastRow, 6)).FormulaR1C1 = _
                "=IF(RC[-2]="""",""""," & _
                "BDP(RC[-2]&"" ISIN""," & _
                """CNTRY_OF_RISK""))"

            wsLookup.Range( _
                wsLookup.Cells(2, 8), _
                wsLookup.Cells(LastRow, 8)).FormulaR1C1 = _
                "=IF(RC[-4]="""",""""," & _
                "BDP(RC[-4]&"" ISIN""," & _
                """INDUSTRY_SECTOR""))"

        End If

    End If


    FormatReportTable _
        wsLookup.Range( _
            wsLookup.Cells(1, 1), _
            wsLookup.Cells(LastRow, 9)), _
        1

    If LastRow >= 2 Then

        FormatFirstColumn _
            wsLookup, _
            2, _
            LastRow, _
            1

    End If

    wsLookup.Columns(4).NumberFormat = "@"
    wsLookup.Range( _
        wsLookup.Cells(1, 1), _
        wsLookup.Cells(LastRow, 9)).HorizontalAlignment = xlLeft

    wsLookup.Range( _
        wsLookup.Cells(1, 1), _
        wsLookup.Cells(LastRow, 9)).Font.name = "Aptos Display"
    wsLookup.Columns("A:I").AutoFit

    If wsLookup.Columns(1).ColumnWidth > 45 Then _
        wsLookup.Columns(1).ColumnWidth = 45

    If wsLookup.Columns(2).ColumnWidth > 60 Then _
        wsLookup.Columns(2).ColumnWidth = 60

    If wsLookup.Columns(3).ColumnWidth > 42 Then _
        wsLookup.Columns(3).ColumnWidth = 42

    If wsLookup.Columns(5).ColumnWidth > 35 Then _
        wsLookup.Columns(5).ColumnWidth = 35

    wsLookup.Columns(2).WrapText = True

    If Not wsLookup.AutoFilterMode Then

        wsLookup.Range( _
            wsLookup.Cells(1, 1), _
            wsLookup.Cells(LastRow, 9)).AutoFilter

    End If


    DeleteLegacyGeographyLookupWorksheet


End Sub

Private Function CanonicalizeGeographyUsingCompanies( _
    ByVal RawEntries As Object, _
    ByVal CompaniesByName As Object, _
    ByVal CompaniesByVariant As Object) As Object

    Dim CanonicalEntries As Object
    Dim Entry As Object
    Dim CompanyEntry As Object
    Dim RawKey As Variant
    Dim RawName As String
    Dim CanonicalName As String

    Set CanonicalEntries = NewExactNameMap()

    If Not RawEntries Is Nothing Then

        For Each RawKey In RawEntries.Keys

            Set Entry = RawEntries(RawKey)
            RawName = CStr(Entry("Name"))
            CanonicalName = RawName

            Set CompanyEntry = _
                ResolveCompanyEntry( _
                    RawName, _
                    CompaniesByName, _
                    CompaniesByVariant)

            If Not CompanyEntry Is Nothing Then

                CanonicalName = CStr(CompanyEntry("Name"))

            End If

            AddGeographyEntryWithCandidates _
                CanonicalEntries, _
                CanonicalName, _
                Entry, _
                RawName

        Next RawKey

    End If


    Set CanonicalizeGeographyUsingCompanies = _
        CanonicalEntries

End Function


'
' Whether the Excl. DPM Certificate subtable would repeat the Full one.  Both
' are formulas now, so the two are evaluated off to the side of the staging
' sheet and what they produce is compared; the scratch is cleared before the
' answer is returned.
'
Private Function CertificateResultsMatch( _
    ByVal DimensionKey As String, _
    ByVal Visibility As Object) As Boolean

    Dim wsStage As Worksheet
    Dim FullCell As Range
    Dim ExDPMCell As Range
    Dim ScratchCol As Long

    If Not RiskSubtableIsVisible( _
            Visibility, _
            DimensionKey, _
            "Certificates") Then

        Exit Function

    End If

    On Error Resume Next
    Set wsStage = ThisWorkbook.Worksheets(RISK_STAGE_SHEET)
    On Error GoTo 0

    If wsStage Is Nothing Then Exit Function

    ScratchCol = RISK_STAGE_FIELD_COUNT + 4

    Set FullCell = wsStage.Cells(1, ScratchCol)
    Set ExDPMCell = wsStage.Cells(1, ScratchCol + 6)

    FullCell.Formula2 = _
        RiskRankedFormula(DimensionKey, "Certificates", False, Visibility)

    ExDPMCell.Formula2 = _
        RiskRankedFormula(DimensionKey, "Certificates", True, Visibility)

    FullCell.Calculate
    ExDPMCell.Calculate

    CertificateResultsMatch = _
        ExposureGroupsMatch( _
            SpilledRange(FullCell), _
            SpilledRange(ExDPMCell))

    wsStage.Range( _
        wsStage.Cells(1, ScratchCol), _
        wsStage.Cells(TOP_NAME_COUNT + 2, ScratchCol + 11)).Clear

End Function

Private Function SpilledRange( _
    ByVal AnchorCell As Range) As Range

    On Error Resume Next
    Set SpilledRange = AnchorCell.SpillingToRange
    On Error GoTo 0

End Function

'
' Two ranked tables are the same when they are the same shape and every cell
' reads the same.  At most ten rows of four, so a cell at a time is fine.
'
Private Function ExposureGroupsMatch( _
    ByVal LeftTable As Range, _
    ByVal RightTable As Range) As Boolean

    Dim r As Long
    Dim c As Long

    If LeftTable Is Nothing Or RightTable Is Nothing Then

        ExposureGroupsMatch = _
            (LeftTable Is Nothing) And (RightTable Is Nothing)

        Exit Function

    End If

    If LeftTable.Rows.Count <> RightTable.Rows.Count Then Exit Function
    If LeftTable.Columns.Count <> RightTable.Columns.Count Then Exit Function

    For r = 1 To LeftTable.Rows.Count
        For c = 1 To LeftTable.Columns.Count

            If CStr(LeftTable.Cells(r, c).Value2) <> _
               CStr(RightTable.Cells(r, c).Value2) Then

                Exit Function

            End If

        Next c
    Next r

    ExposureGroupsMatch = True

End Function

'
' A concentration subtable is three things: the dimension it groups by, the
' asset class it covers, and the account scope it counts.  Every formula
' below is built from those against the RiskExposure staging table, so a
' number on the report and the rows behind it are one click apart.
'
' The rules are the ones the row-by-row aggregation used to apply.  Issuer
' takes every row of a class, and a row that resolved to no entity is
' aggregated under a prefixed name the ranked list skips and the share
' denominator still counts - which is why Issuer builds two filters where
' Geography and Sector build one.  Blank geography or sector reads as Others.
'

'
' The order the subtables appear in, and the names printed above them.
' Sovereign Bonds and Funds sit between Corporate Bonds and Certificates,
' which is where v58 put them.
'
Private Function RiskDisplayedClasses() As Variant

    RiskDisplayedClasses = Array( _
        "Equity", _
        CORPORATE_BONDS_CLASS, _
        SOVEREIGN_BONDS_CLASS, _
        "Funds", _
        "Certificates", _
        "Overall")

End Function

Private Function RiskClassDisplayName( _
    ByVal AssetClassKey As String) As String

    Select Case AssetClassKey

        Case CORPORATE_BONDS_CLASS

            RiskClassDisplayName = CORPORATE_BONDS_DISPLAY_NAME

        Case "Certificates"

            RiskClassDisplayName = "Certificates (Excl. Protected)"

        Case Else

            RiskClassDisplayName = AssetClassKey

    End Select

End Function

Private Function RiskStageColumn( _
    ByVal HeaderName As String) As String

    RiskStageColumn = RISK_STAGE_TABLE & "[" & HeaderName & "]"

End Function

Private Function RiskDimensionHeader( _
    ByVal DimensionKey As String) As String

    Select Case DimensionKey

        Case "Issuer"

            RiskDimensionHeader = "Exposure Name"

        Case "Country"

            RiskDimensionHeader = "Geography"

        Case "Sector"

            RiskDimensionHeader = "Sector"

    End Select

End Function

Private Function RiskDimensionBinding( _
    ByVal DimensionKey As String) As String

    Select Case DimensionKey

        Case "Issuer"

            RiskDimensionBinding = "nme"

        Case "Country"

            RiskDimensionBinding = "geo"

        Case "Sector"

            RiskDimensionBinding = "sec"

    End Select

End Function

'
' Only Issuer measures a share against rows it cannot rank.
'
Private Function RiskDimensionRanksEveryRow( _
    ByVal DimensionKey As String) As Boolean

    RiskDimensionRanksEveryRow = (DimensionKey <> "Issuer")

End Function

Private Function RiskRankedClasses() As Variant

    RiskRankedClasses = Array( _
        "Equity", _
        CORPORATE_BONDS_CLASS, _
        SOVEREIGN_BONDS_CLASS, _
        "Certificates", _
        "Funds")

End Function

'
' The Risk Asset Class values a subtable answers to.  Certificates is written
' into the staging table as "Certificates (Excl. Protected)"; the bare name is
' matched as well, exactly as RiskAssetIndexFromClass used to accept both.
'
Private Function RiskClassStagingValues( _
    ByVal AssetClassKey As String) As Variant

    Select Case AssetClassKey

        Case "Certificates"

            RiskClassStagingValues = _
                Array("Certificates", "Certificates (Excl. Protected)")

        Case Else

            RiskClassStagingValues = Array(AssetClassKey)

    End Select

End Function

'
' Overall covers whatever the dimension still shows, so hiding a class takes
' it out of the Overall denominator too.
'
Private Function RiskClassTest( _
    ByVal AssetClassKey As String, _
    ByVal DimensionKey As String, _
    ByVal Visibility As Object) As String

    Dim Names As Collection
    Dim ClassKey As Variant
    Dim StagingValue As Variant
    Dim List As String

    Set Names = New Collection

    If StrComp(AssetClassKey, "Overall", vbTextCompare) = 0 Then

        For Each ClassKey In RiskRankedClasses()

            If RiskSubtableIsVisible( _
                    Visibility, _
                    DimensionKey, _
                    CStr(ClassKey)) Then

                For Each StagingValue In RiskClassStagingValues(CStr(ClassKey))
                    Names.Add CStr(StagingValue)
                Next StagingValue

            End If

        Next ClassKey

    Else

        For Each StagingValue In RiskClassStagingValues(AssetClassKey)
            Names.Add CStr(StagingValue)
        Next StagingValue

    End If

    If Names.Count = 0 Then

        RiskClassTest = "FALSE"

    ElseIf Names.Count = 1 Then

        RiskClassTest = "(cls = """ & Names(1) & """)"

    Else

        For Each StagingValue In Names

            If List <> "" Then List = List & ", "
            List = List & """" & CStr(StagingValue) & """"

        Next StagingValue

        RiskClassTest = "ISNUMBER(MATCH(cls, {" & List & "}, 0))"

    End If

End Function

'
' One LET binding, its name padded so the expressions line up under each
' other in the formula bar.
'
Private Function RiskBindLine( _
    ByVal BindName As String, _
    ByVal Expression As String) As String

    Dim Label As String

    Label = BindName & ","

    RiskBindLine = _
        RISK_FORMULA_INDENT & Label & _
        Space$(RISK_BIND_WIDTH - Len(Label)) & Expression

End Function

'
' The LET preamble, emitting exactly the bindings the caller says it reads.
' A LET that declares a name nothing reads is not worth finding out about the
' hard way, and the denominators read fewer of these than the ranked table
' does: no dimension column, and for Issuer no entity test either.
'
Private Function RiskBindings( _
    ByVal DimensionKey As String, _
    ByVal Wanted As String, _
    ByVal ExcludeDPM As Boolean) As String

    Dim DimensionName As String
    Dim Header As String
    Dim Lines As String

    DimensionName = RiskDimensionBinding(DimensionKey)
    Header = RiskStageColumn(RiskDimensionHeader(DimensionKey))

    If InStr(1, Wanted, "cls") > 0 Then

        Lines = Lines & _
            RiskBindLine("cls", RiskStageColumn("Risk Asset Class") & ",") & _
            vbLf

    End If

    If InStr(1, Wanted, "typ") > 0 Then

        Lines = Lines & _
            RiskBindLine("typ", RiskStageColumn("Exposure Type") & ",") & vbLf

    End If

    If InStr(1, Wanted, "nme") > 0 Then

        Lines = Lines & _
            RiskBindLine("nme", RiskStageColumn("Exposure Name") & ",") & vbLf

    End If

    If InStr(1, Wanted, "dim") > 0 And DimensionName <> "nme" Then

        Lines = Lines & _
            RiskBindLine( _
                DimensionName, _
                "IF(TRIM(" & Header & ") = """", """ & _
                OTHER_RISK_DIMENSION & """, TRIM(" & Header & ")),") & vbLf

    End If

    If ExcludeDPM Then

        Lines = Lines & _
            RiskBindLine("scp", RiskStageColumn("Account Scope") & ",") & vbLf

    End If

    If InStr(1, Wanted, "ndg") > 0 Then

        Lines = Lines & _
            RiskBindLine("ndg", RiskStageColumn("NDG") & ",") & vbLf

    End If

    If InStr(1, Wanted, "val") > 0 Then

        Lines = Lines & _
            RiskBindLine( _
                "val", _
                RiskStageColumn("Allocated Collateral Value") & ",") & vbLf

    End If

    RiskBindings = Lines

End Function

'
' The rows one subtable aggregates, a factor per line so the conditions read
' as a list.  EntityOnly is False only for the Issuer share denominator.
'
Private Function RiskKeepExpression( _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal EntityOnly As Boolean, _
    ByVal Visibility As Object) As String

    Dim Factors As Collection
    Dim Factor As Variant

    Set Factors = New Collection

    Factors.Add RiskClassTest(AssetClassKey, DimensionKey, Visibility)

    If EntityOnly Then

        Factors.Add "(typ <> """ & UNKNOWN_UNDERLYING_TYPE & """)"

        Factors.Add _
            "(LEFT(nme, " & CStr(Len(UNKNOWN_UNDERLYING_PREFIX)) & _
            ") <> """ & UNKNOWN_UNDERLYING_PREFIX & """)"

    End If

    If ExcludeDPM Then
        Factors.Add "(TRIM(scp) = """ & NON_DPM_SCOPE & """)"
    End If

    For Each Factor In Factors

        If Len(RiskKeepExpression) = 0 Then

            RiskKeepExpression = CStr(Factor)

        Else

            RiskKeepExpression = _
                RiskKeepExpression & vbLf & RISK_FORMULA_INDENT & _
                Space$(RISK_BIND_WIDTH - 2) & "* " & CStr(Factor)

        End If

    Next Factor

End Function

'
' The ranked table in one formula: name, value, share and distinct NDG count
' for the top ten, ordered by value descending with ties broken by name
' ascending.  GROUPBY takes a LAMBDA where an aggregate is wanted, so the
' distinct NDG count is written where SUM sits.
'
' GROUPBY names its value columns in a row of its own, and that row is text,
' so sorting by value descending would carry it to the top and push the tenth
' name out of the table.  DROP takes it off before the sort.
'
Private Function RiskRankedFormula( _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal Visibility As Object) As String

    Dim RanksEveryRow As Boolean
    Dim Formula As String

    RanksEveryRow = RiskDimensionRanksEveryRow(DimensionKey)

    Formula = _
        "=LET(" & vbLf & _
        RiskBindings(DimensionKey, "cls,typ,nme,dim,ndg,val", ExcludeDPM) & _
        vbLf & _
        RiskBindLine( _
            "k", _
            RiskKeepExpression( _
                DimensionKey, AssetClassKey, ExcludeDPM, True, Visibility) & ",")

    If Not RanksEveryRow Then

        Formula = Formula & vbLf & _
            RiskBindLine( _
                "kAll", _
                RiskKeepExpression( _
                    DimensionKey, AssetClassKey, ExcludeDPM, False, Visibility) & ",")

    End If

    Formula = Formula & vbLf & vbLf & _
        RiskBindLine( _
            "g", _
            "FILTER(" & RiskDimensionBinding(DimensionKey) & ", k),") & vbLf & _
        RiskBindLine("v", "FILTER(val, k),") & vbLf & _
        RiskBindLine("n", "FILTER(ndg, k),") & vbLf & vbLf & _
        RiskBindLine( _
            "a", _
            "DROP(GROUPBY(g, HSTACK(v, n), " & _
            "HSTACK(SUM, LAMBDA(x, COUNTA(UNIQUE(x)))), 0, 0), 1),") & vbLf & _
        RiskBindLine( _
            "s", _
            "TAKE(SORTBY(a, CHOOSECOLS(a, 2), -1, CHOOSECOLS(a, 1), 1), " & _
            CStr(TOP_NAME_COUNT) & "),") & vbLf & _
        RiskBindLine( _
            "d", _
            IIf(RanksEveryRow, "SUM(v),", "SUM(FILTER(val, kAll, 0)),")) & _
        vbLf & vbLf & _
        RISK_FORMULA_INDENT & "IFERROR(" & vbLf & _
        RISK_FORMULA_INDENT & RISK_FORMULA_INDENT & _
        "HSTACK(CHOOSECOLS(s, 1), CHOOSECOLS(s, 2), " & _
        "CHOOSECOLS(s, 2) / d, CHOOSECOLS(s, 3))," & vbLf & _
        RISK_FORMULA_INDENT & RISK_FORMULA_INDENT & """"")" & vbLf & _
        ")"

    RiskRankedFormula = Formula

End Function

'
' The share denominator: for Geography and Sector the total of the rows that
' were ranked, for Issuer the larger total that includes the rows it cannot
' rank.
'
Private Function RiskCategoryTotalFormula( _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal Visibility As Object) As String

    Dim EntityOnly As Boolean

    EntityOnly = RiskDimensionRanksEveryRow(DimensionKey)

    RiskCategoryTotalFormula = _
        "LET(" & vbLf & _
        RiskBindings( _
            DimensionKey, _
            IIf(EntityOnly, "cls,typ,nme,val", "cls,val"), _
            ExcludeDPM) & _
        vbLf & _
        RiskBindLine( _
            "k", _
            RiskKeepExpression( _
                DimensionKey, AssetClassKey, ExcludeDPM, EntityOnly, _
                Visibility) & ",") & _
        vbLf & vbLf & _
        RISK_FORMULA_INDENT & "SUM(FILTER(val, k, 0))" & vbLf & _
        ")"

End Function

'
' The distinct NDGs of the ten names taken together - a union, not the sum of
' the ten counts above it.
'
Private Function RiskUnionNdgFormula( _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal Visibility As Object, _
    ByVal NameRange As String) As String

    RiskUnionNdgFormula = _
        "=LET(" & vbLf & _
        RiskBindings(DimensionKey, "cls,typ,nme,dim,ndg", ExcludeDPM) & _
        vbLf & _
        RiskBindLine( _
            "k", _
            RiskKeepExpression( _
                DimensionKey, AssetClassKey, ExcludeDPM, True, Visibility) & ",") & _
        vbLf & vbLf & _
        RISK_FORMULA_INDENT & "IFERROR(COUNTA(UNIQUE(FILTER(ndg," & vbLf & _
        RISK_FORMULA_INDENT & RISK_FORMULA_INDENT & "k * ISNUMBER(MATCH(" & _
        RiskDimensionBinding(DimensionKey) & ", " & NameRange & ", 0))))), 0)" & _
        vbLf & ")"

End Function

'
' How many rows the ranked formula produced.  A class with nothing in it
' answers with the empty string rather than an array, which is the None
' case and the one reason this returns nought.
'
Private Function RiskSpilledRowCount( _
    ByVal AnchorCell As Range) As Long

    Dim Spill As Range

    If Len(CStr(AnchorCell.Value2)) = 0 Then Exit Function

    On Error Resume Next
    Set Spill = AnchorCell.SpillingToRange
    On Error GoTo 0

    If Spill Is Nothing Then Exit Function

    RiskSpilledRowCount = Spill.Rows.Count

End Function

Private Function WriteSameAsLeftExposureGroup( _
    ByVal ws As Worksheet, _
    ByVal StartRow As Long, _
    ByVal LeftCol As Long, _
    ByVal AssetClass As String) As Long

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4))

        .Merge
        .Value = AssetClass

    End With

    With ws.Range( _
            ws.Cells(StartRow + 1, LeftCol), _
            ws.Cells(StartRow + 1, LeftCol + 4))

        .Merge
        .Value = "Same as left"

    End With

    FormatReportTable _
        ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow + 1, LeftCol + 4)), _
        1

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4))

        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(111, 38, 61)

        With .Borders(xlEdgeBottom)

            .LineStyle = xlNone

        End With

        With .Borders(xlEdgeTop)

            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(111, 38, 61)

        End With

    End With

    With ws.Range( _
            ws.Cells(StartRow + 1, LeftCol), _
            ws.Cells(StartRow + 1, LeftCol + 4))

        .HorizontalAlignment = xlLeft
        .Font.Italic = True

    End With

    WriteSameAsLeftExposureGroup = StartRow + 2

End Function

'
' One concentration table: a title, a header row, and a subtable for each
' asset class the visibility list still shows.  Issuer, Geography and Sector
' differ only in the column they group by and what the table is called, so
' they share this - it used to be two copies, WriteRiskGranularityTable
' being the Issuer one.
'
Private Function WriteRiskDimensionTable( _
    ByVal ws As Worksheet, _
    ByVal TopRow As Long, _
    ByVal LeftCol As Long, _
    ByVal SectionTitle As String, _
    ByVal DimensionLabel As String, _
    ByVal DimensionKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal Visibility As Object, _
    Optional ByVal CertificateSameAsLeft As Boolean = False) As Long

    Dim ClassKey As Variant
    Dim CurrentRow As Long

    WriteSectionTitle _
        ws, _
        TopRow, _
        LeftCol, _
        5, _
        SectionTitle

    ws.Cells(TopRow + 1, LeftCol).Value = "Rank"
    ws.Cells(TopRow + 1, LeftCol + 1).Value = DimensionLabel
    ws.Cells(TopRow + 1, LeftCol + 2).Value = "Collateral Value"
    ws.Cells(TopRow + 1, LeftCol + 3).Value = "% of Category"
    ws.Cells(TopRow + 1, LeftCol + 4).Value = "#NDG"

    CurrentRow = TopRow + 2

    For Each ClassKey In RiskDisplayedClasses()

        If RiskSubtableIsVisible( _
                Visibility, _
                DimensionKey, _
                CStr(ClassKey)) Then

            If CertificateSameAsLeft _
               And StrComp(CStr(ClassKey), "Certificates", _
                           vbTextCompare) = 0 Then

                CurrentRow = _
                    WriteSameAsLeftExposureGroup( _
                        ws, _
                        CurrentRow, _
                        LeftCol, _
                        RiskClassDisplayName(CStr(ClassKey)))

            Else

                CurrentRow = _
                    WriteTopExposureGroup( _
                        ws, _
                        CurrentRow, _
                        LeftCol, _
                        RiskClassDisplayName(CStr(ClassKey)), _
                        DimensionKey, _
                        CStr(ClassKey), _
                        ExcludeDPM, _
                        Visibility)

            End If

        End If

    Next ClassKey

    If CurrentRow > TopRow + 2 Then

        FormatReportTable _
            ws.Range( _
                ws.Cells(TopRow + 1, LeftCol), _
                ws.Cells(TopRow + 1, LeftCol + 4)), _
            1, _
            ws.Range( _
                ws.Cells(TopRow + 2, LeftCol), _
                ws.Cells(CurrentRow - 1, LeftCol + 4))

    Else

        FormatReportTable _
            ws.Range( _
                ws.Cells(TopRow + 1, LeftCol), _
                ws.Cells(TopRow + 1, LeftCol + 4)), _
            1

    End If

    WriteRiskDimensionTable = CurrentRow

End Function

Private Function NewRiskStageRows() As Collection

    Set NewRiskStageRows = New Collection

End Function

Private Sub AppendRiskStageRow( _
    ByVal StageRows As Collection, _
    ByVal SourceRow As Long, _
    ByVal NDG As String, _
    ByVal ProductISIN As String, _
    ByVal SecurityName As String, _
    ByVal OriginalAssetType As String, _
    ByVal AdditionalComment As String, _
    ByVal RiskAssetClass As String, _
    ByVal ExposureName As String, _
    ByVal ExposureType As String, _
    ByVal AllocationWeight As Double, _
    ByVal PositionValue As Double, _
    ByVal ResolutionSource As String, _
    ByVal AccountScope As String)

    Dim StageRow As Variant

    If StageRows Is Nothing Then Exit Sub

    '
    ' A row with no exposure name cannot be aggregated under anything, so it
    ' is dropped - but dropping every row of a run without a word is how a
    ' staging table comes to look as though there was nothing to report.
    '
    If Trim(ExposureName) = "" Then

        RiskStageRowsDropped = RiskStageRowsDropped + 1

        Exit Sub

    End If

    ReDim StageRow(1 To RISK_STAGE_FIELD_COUNT)

    StageRow(RiskStageSourceRow) = SourceRow
    StageRow(RiskStageNDG) = NDG
    StageRow(RiskStageProductISIN) = UCase(Trim(ProductISIN))
    StageRow(RiskStageSecurityName) = SecurityName
    StageRow(RiskStageOriginalAssetType) = OriginalAssetType
    StageRow(RiskStageAdditionalComment) = AdditionalComment
    StageRow(RiskStageAssetClass) = RiskAssetClass
    StageRow(RiskStageExposureName) = ExposureName
    StageRow(RiskStageExposureType) = ExposureType
    StageRow(RiskStageAllocationWeight) = AllocationWeight
    StageRow(RiskStagePositionValue) = PositionValue
    StageRow(RiskStageAllocatedValue) = _
        PositionValue * AllocationWeight
    StageRow(RiskStageGeography) = ""
    StageRow(RiskStageSector) = ""
    StageRow(RiskStageResolutionSource) = ResolutionSource
    StageRow(RiskStageAccountScope) = AccountScope

    StageRows.Add StageRow

End Sub

Private Sub AppendCertificateRiskStageRows( _
    ByVal StageRows As Collection, _
    ByVal CertificateMap As Object, _
    ByVal FundMap As Object, _
    ByVal UnderlyingAssetClassMap As Object, _
    ByVal ProductISIN As String, _
    ByVal SecurityName As String, _
    ByVal OriginalAssetType As String, _
    ByVal AdditionalComment As String, _
    ByVal NDG As String, _
    ByVal PositionValue As Double, _
    ByVal SourceRow As Long, _
    ByVal AccountScope As String, _
    ByRef UsedTemporaryUnderlying As Boolean)

    Dim Underlyings As Collection
    Dim Underlying As Variant
    Dim Component As Object

    Dim CertificateKey As String
    Dim UnderlyingName As String
    Dim OriginalUnderlyingName As String
    Dim UnderlyingKey As String
    Dim UnderlyingAssetClass As String
    Dim ParentCompany As String
    Dim ExposureType As String
    Dim ResolutionSource As String
    Dim TemporaryUnderlying As String
    Dim AllocationWeight As Double

    UsedTemporaryUnderlying = False
    CertificateKey = UCase(Trim(ProductISIN))

    If CertificateKey <> "" Then

        If CertificateMap.Exists(CertificateKey) Then

            Set Underlyings = CertificateMap(CertificateKey)

        End If

    End If

    If Underlyings Is Nothing Then

        CertificateKey = Trim(SecurityName)

        If CertificateKey <> "" Then

            If CertificateMap.Exists(CertificateKey) Then

                Set Underlyings = CertificateMap(CertificateKey)

            End If

        End If

    End If

    If Underlyings Is Nothing Then GoTo UseTemporaryUnderlying
    If Underlyings.Count = 0 Then GoTo UseTemporaryUnderlying

    For Each Underlying In Underlyings

        Set Component = Underlying
        OriginalUnderlyingName = CStr(Component("Name"))
        UnderlyingName = OriginalUnderlyingName
        UnderlyingAssetClass = CStr(Component("AssetClass"))
        AllocationWeight = CDbl(Component("Weight"))
        ExposureType = "Certificate underlying"
        ResolutionSource = "Certificate reference files"

        UnderlyingKey = _
            NormalizeExactNameKey(OriginalUnderlyingName)

        If UnderlyingAssetClass = "" And _
           Not UnderlyingAssetClassMap Is Nothing Then

            If UnderlyingAssetClassMap.Exists(UnderlyingKey) Then

                UnderlyingAssetClass = _
                    CStr(UnderlyingAssetClassMap(UnderlyingKey))

            End If

        End If

        If Left( _
                UnderlyingName, _
                Len(UNKNOWN_UNDERLYING_PREFIX)) = _
           UNKNOWN_UNDERLYING_PREFIX Then

            UnderlyingName = _
                Mid( _
                    UnderlyingName, _
                    Len(UNKNOWN_UNDERLYING_PREFIX) + 1)

            If Trim(UnderlyingName) = "" Then

                UnderlyingName = "Non-entity component"

            End If

            ExposureType = UNKNOWN_UNDERLYING_TYPE

        ElseIf Left( _
                UnderlyingName, _
                Len(MISSING_CERTIFICATE_RIC_PREFIX)) = _
           MISSING_CERTIFICATE_RIC_PREFIX Then

            UnderlyingName = _
                BuildTemporaryIssuerName( _
                    ProductISIN, _
                    SecurityName, _
                    SourceRow, _
                    PlaceholderFromSecurityName)

            ResolutionSource = "Fallback"
            UsedTemporaryUnderlying = True

        Else

            If UnderlyingAssetClass <> "" Then

                ExposureType = _
                    ExposureType & " - " & UnderlyingAssetClass

            End If

            If StrComp( _
                    UnderlyingAssetClass, _
                    "Fund", _
                    vbTextCompare) = 0 Or _
               StrComp( _
                    UnderlyingAssetClass, _
                    "ETF", _
                    vbTextCompare) = 0 Then

                ParentCompany = _
                    ResolveExactName( _
                        UnderlyingName, _
                        FundMap)

                If ParentCompany <> "" Then

                    UnderlyingName = ParentCompany
                    ExposureType = _
                        "Certificate underlying fund parent company"
                    ResolutionSource = _
                        "Certificate reference files + Funds"

                End If

            End If

            If CBool(Component("Temporary")) Then

                UsedTemporaryUnderlying = True

            End If

        End If

        AppendRiskStageRow _
            StageRows, _
            SourceRow, _
            NDG, _
            ProductISIN, _
            SecurityName, _
            OriginalAssetType, _
            AdditionalComment, _
            "Certificates (Excl. Protected)", _
            UnderlyingName, _
            ExposureType, _
            AllocationWeight, _
            PositionValue, _
            ResolutionSource, _
            AccountScope

    Next Underlying

    Exit Sub

UseTemporaryUnderlying:

    TemporaryUnderlying = _
        BuildTemporaryIssuerName( _
            ProductISIN, _
            SecurityName, _
            SourceRow, _
            PlaceholderFromSecurityName)

    AppendRiskStageRow _
        StageRows, _
        SourceRow, _
        NDG, _
        ProductISIN, _
        SecurityName, _
        OriginalAssetType, _
        AdditionalComment, _
        "Certificates (Excl. Protected)", _
        TemporaryUnderlying, _
        "Certificate underlying", _
        1, _
        PositionValue, _
        "Fallback", _
        AccountScope

    UsedTemporaryUnderlying = True

End Sub

Private Function FinalizeRiskStageData( _
    ByVal StageRows As Collection, _
    ByVal CanonicalNameMap As Object, _
    ByVal CompaniesByName As Object, _
    ByVal CompaniesByVariant As Object, _
    ByVal CountryNameMap As Object, _
    ByVal ResolutionCache As Object) As Variant

    Dim StageData() As Variant
    Dim StageRow As Variant
    Dim CompanyEntry As Object

    Dim RowNo As Long
    Dim ColNo As Long
    Dim RawName As String
    Dim CanonicalName As String
    Dim ExposureType As String
    Dim GeographyCode As String
    Dim GeographyName As String
    Dim SectorName As String
    Dim MappedCountryName As String
    Dim CacheKey As String
    Dim CachedValues As Variant
    Dim IsEntityExposure As Boolean
    Dim MappingWasCached As Boolean

    If StageRows Is Nothing Then Exit Function
    If StageRows.Count = 0 Then Exit Function

    ReDim StageData( _
        1 To StageRows.Count, _
        1 To RISK_STAGE_FIELD_COUNT)

    For Each StageRow In StageRows

        RowNo = RowNo + 1
        RawName = CStr(StageRow(RiskStageExposureName))
        CanonicalName = RawName
        ExposureType = _
            CStr(StageRow(RiskStageExposureType))
        GeographyCode = ""
        GeographyName = ""
        SectorName = ""
        CacheKey = ""
        MappingWasCached = False
        Set CompanyEntry = Nothing

        IsEntityExposure = _
            (Not IsUnknownUnderlyingType(ExposureType) And _
             Left( _
                RawName, _
                Len(UNKNOWN_UNDERLYING_PREFIX)) <> _
             UNKNOWN_UNDERLYING_PREFIX)

        If IsEntityExposure Then

            CacheKey = NormalizeExactNameKey(RawName)

            If CacheKey <> "" And _
               Not ResolutionCache Is Nothing Then

                If ResolutionCache.Exists(CacheKey) Then

                    CachedValues = ResolutionCache(CacheKey)
                    CanonicalName = CStr(CachedValues(0))
                    GeographyName = CStr(CachedValues(1))
                    SectorName = CStr(CachedValues(2))
                    MappingWasCached = True

                End If

            End If

        End If

        If IsEntityExposure And Not MappingWasCached Then

            CanonicalName = _
                ResolveCanonicalEntityName( _
                    RawName, _
                    CanonicalNameMap)

            Set CompanyEntry = _
                ResolveCompanyEntry( _
                    CanonicalName, _
                    CompaniesByName, _
                    CompaniesByVariant)

            If Not CompanyEntry Is Nothing Then

                CanonicalName = CStr(CompanyEntry("Name"))

            End If

            GeographyCode = _
                CompanyDimensionValue( _
                    CompanyEntry, _
                    "Country")

            GeographyName = GeographyCode
            MappedCountryName = _
                ResolveExactName( _
                    GeographyCode, _
                    CountryNameMap)

            If MappedCountryName <> "" Then

                GeographyName = MappedCountryName

            End If

            SectorName = _
                CompanyDimensionValue( _
                    CompanyEntry, _
                    "Sector")

            If CacheKey <> "" And _
               Not ResolutionCache Is Nothing Then

                ResolutionCache(CacheKey) = _
                    Array( _
                        CanonicalName, _
                        GeographyName, _
                        SectorName)

            End If

        End If

        StageRow(RiskStageExposureName) = CanonicalName
        StageRow(RiskStageGeography) = GeographyName
        StageRow(RiskStageSector) = SectorName

        For ColNo = 1 To RISK_STAGE_FIELD_COUNT

            StageData(RowNo, ColNo) = StageRow(ColNo)

        Next ColNo

    Next StageRow


    FinalizeRiskStageData = StageData

End Function

Private Function RiskStageRowCount( _
    ByRef StageData As Variant) As Long

    On Error GoTo NoRows

    If Not IsArray(StageData) Then Exit Function

    RiskStageRowCount = UBound(StageData, 1)

NoRows:

End Function

Private Function RiskStageHeaders() As Variant

    RiskStageHeaders = Array( _
        "Source Row", _
        "NDG", _
        "Product ISIN", _
        "Security Name", _
        "Original Asset Type", _
        "Additional Comment", _
        "Risk Asset Class", _
        "Exposure Name", _
        "Exposure Type", _
        "Allocation Weight", _
        "Position Value", _
        "Allocated Collateral Value", _
        "Geography", _
        "Sector", _
        "Resolution Source", _
        "Account Scope")

End Function

Private Function GetRiskStageTable( _
    ByVal WorksheetName As String, _
    ByVal TableName As String) As ListObject

    Dim wsStage As Worksheet

    Set wsStage = GetOptionalWorksheet(WorksheetName)

    If wsStage Is Nothing Then Exit Function

    On Error Resume Next

    Set GetRiskStageTable = wsStage.ListObjects(TableName)

    On Error GoTo 0

End Function

Private Function RiskStageTableSchemaIsValid( _
    ByVal StageTable As ListObject) As Boolean

    Dim Headers As Variant
    Dim FieldNo As Long

    If StageTable Is Nothing Then Exit Function

    Headers = RiskStageHeaders()

    If StageTable.ListColumns.Count < _
       RISK_STAGE_FIELD_COUNT Then Exit Function

    For FieldNo = 1 To RISK_STAGE_FIELD_COUNT

        If GetTableColumnIndex( _
                StageTable, _
                CStr(Headers(FieldNo - 1))) = 0 Then

            Exit Function

        End If

    Next FieldNo

    RiskStageTableSchemaIsValid = True

End Function

Private Function RiskStageTableHasData( _
    ByVal StageTable As ListObject) As Boolean

    Dim ExposureNameColumn As Long
    Dim AssetClassColumn As Long
    Dim r As Long

    If Not RiskStageTableSchemaIsValid(StageTable) Then Exit Function
    If StageTable.DataBodyRange Is Nothing Then Exit Function

    ExposureNameColumn = _
        GetTableColumnIndex( _
            StageTable, _
            "Exposure Name")
    AssetClassColumn = _
        GetTableColumnIndex( _
            StageTable, _
            "Risk Asset Class")

    For r = 1 To StageTable.ListRows.Count

        If SafeText( _
                StageTable.DataBodyRange.Cells( _
                    r, _
                    ExposureNameColumn).Value) <> "" Or _
           SafeText( _
                StageTable.DataBodyRange.Cells( _
                    r, _
                    AssetClassColumn).Value) <> "" Then

            RiskStageTableHasData = True

            Exit Function

        End If

    Next r

End Function

Private Function RiskStageTableCanBeReused() As Boolean

    Dim StageTable As ListObject

    Set StageTable = _
        GetRiskStageTable( _
            RISK_STAGE_SHEET, _
            RISK_STAGE_TABLE)

    If Not RiskStageTableSchemaIsValid(StageTable) Then Exit Function
    If Not RiskStageTableHasData(StageTable) Then Exit Function

    RiskStageTableCanBeReused = True

End Function

Private Function CurrentRiskStageAnalysisDate() As Date

    Dim RawDate As Variant

    On Error Resume Next

    RawDate = _
        ThisWorkbook.Worksheets("Home") _
        .Range("WeeklyEndDate").Value

    On Error GoTo 0

    If IsDate(RawDate) Then

        CurrentRiskStageAnalysisDate = CDate(RawDate)

    End If

End Function

Private Function StoredRiskStageDate( _
    ByVal WorksheetName As String) As Date

    Dim wsStage As Worksheet
    Dim RawDate As Variant

    Set wsStage = GetOptionalWorksheet(WorksheetName)

    If wsStage Is Nothing Then Exit Function

    RawDate = wsStage.Cells(1, 3).Value

    If IsDate(RawDate) Then

        StoredRiskStageDate = CDate(RawDate)

    End If

End Function

Private Function RiskStageReuseDescription( _
    ByVal AnalysisDate As Date) As String

    Dim StageDate As Date

    StageDate = StoredRiskStageDate(RISK_STAGE_SHEET)

    If StageDate > 0 Then

        If AnalysisDate > 0 And _
           StageDate <> AnalysisDate Then

            RiskStageReuseDescription = _
                "risk detail table (existing as of " & _
                Format(StageDate, "dd/mm/yyyy") & _
                "; requested " & _
                Format(AnalysisDate, "dd/mm/yyyy") & _
                ")"

        Else

            RiskStageReuseDescription = _
                "risk detail table (as of " & _
                Format(StageDate, "dd/mm/yyyy") & _
                ")"

        End If

    Else

        RiskStageReuseDescription = _
            "risk detail table (source date not recorded)"

    End If

End Function

Private Function ShouldRebuildRiskStageTables( _
    ByVal AnalysisDate As Date) As Boolean

    If Not RiskStageTableCanBeReused() Then

        ShouldRebuildRiskStageTables = True

        Exit Function

    End If

    ShouldRebuildRiskStageTables = _
        ShouldOverwriteExistingSheets( _
            RiskStageReuseDescription(AnalysisDate))

End Function

Private Function LoadRiskStageTableData( _
    ByVal WorksheetName As String, _
    ByVal TableName As String) As Variant

    Dim StageTable As ListObject
    Dim Headers As Variant
    Dim RawData As Variant
    Dim StageData() As Variant
    Dim ColumnIndexes(1 To RISK_STAGE_FIELD_COUNT) As Long

    Dim SourceRow As Long
    Dim TargetRow As Long
    Dim FieldNo As Long
    Dim ValidRowCount As Long
    Dim HasRiskData As Boolean

    Set StageTable = _
        GetRiskStageTable( _
            WorksheetName, _
            TableName)

    If Not RiskStageTableSchemaIsValid(StageTable) Then

        Err.Raise _
            vbObjectError + 9180, _
            "LoadRiskStageTableData", _
            "Risk staging table '" & TableName & _
            "' is missing or has an invalid schema."

    End If

    If StageTable.DataBodyRange Is Nothing Then Exit Function

    Headers = RiskStageHeaders()

    For FieldNo = 1 To RISK_STAGE_FIELD_COUNT

        ColumnIndexes(FieldNo) = _
            GetTableColumnIndex( _
                StageTable, _
                CStr(Headers(FieldNo - 1)))

    Next FieldNo

    RawData = StageTable.DataBodyRange.Value2

    For SourceRow = 1 To UBound(RawData, 1)

        HasRiskData = _
            (SafeText( _
                RawData( _
                    SourceRow, _
                    ColumnIndexes( _
                        RiskStageExposureName))) <> "" Or _
             SafeText( _
                RawData( _
                    SourceRow, _
                    ColumnIndexes( _
                        RiskStageAssetClass))) <> "")

        If HasRiskData Then

            ValidRowCount = ValidRowCount + 1

        End If

    Next SourceRow

    If ValidRowCount = 0 Then Exit Function

    ReDim StageData( _
        1 To ValidRowCount, _
        1 To RISK_STAGE_FIELD_COUNT)

    For SourceRow = 1 To UBound(RawData, 1)

        HasRiskData = _
            (SafeText( _
                RawData( _
                    SourceRow, _
                    ColumnIndexes( _
                        RiskStageExposureName))) <> "" Or _
             SafeText( _
                RawData( _
                    SourceRow, _
                    ColumnIndexes( _
                        RiskStageAssetClass))) <> "")

        If HasRiskData Then

            TargetRow = TargetRow + 1

            For FieldNo = 1 To RISK_STAGE_FIELD_COUNT

                StageData(TargetRow, FieldNo) = _
                    RawData( _
                        SourceRow, _
                        ColumnIndexes(FieldNo))

            Next FieldNo

        End If

    Next SourceRow

    LoadRiskStageTableData = StageData

End Function

Private Function RiskStageRowIsDPM( _
    ByRef StageData As Variant, _
    ByVal RowNo As Long) As Boolean

    Dim ScopeText As String

    ScopeText = _
        UCase$( _
            Trim$( _
                SafeText( _
                    StageData( _
                        RowNo, _
                        RiskStageAccountScope))))

    Select Case ScopeText

        Case "DPM"

            RiskStageRowIsDPM = True

        Case "NON-DPM"

            RiskStageRowIsDPM = False

        Case Else

            Err.Raise _
                vbObjectError + 9181, _
                "RiskStageRowIsDPM", _
                "Risk staging row " & CStr(RowNo) & _
                " has an invalid Account Scope: '" & _
                SafeText( _
                    StageData( _
                        RowNo, _
                        RiskStageAccountScope)) & _
                "'."

    End Select

End Function

Private Sub DeleteLegacyRiskStageWorksheets()

    Dim SheetName As Variant
    Dim wsLegacy As Worksheet
    Dim PreviousDisplayAlerts As Boolean

    PreviousDisplayAlerts = Application.DisplayAlerts

    On Error Resume Next

    Application.DisplayAlerts = False

    For Each SheetName In Array( _
        LEGACY_RISK_STAGE_NON_DPM_SHEET, _
        LEGACY_RISK_STAGE_DPM_SHEET)

        Set wsLegacy = Nothing
        Set wsLegacy = _
            GetOptionalWorksheet(CStr(SheetName))

        If Not wsLegacy Is Nothing Then wsLegacy.Delete

    Next SheetName

    Application.DisplayAlerts = PreviousDisplayAlerts

    On Error GoTo 0

End Sub

Private Function EnsureRiskStageWorksheet( _
    ByVal WorksheetName As String) As Worksheet

    Dim StageWorksheet As Worksheet

    Set StageWorksheet = GetOptionalWorksheet(WorksheetName)

    If StageWorksheet Is Nothing Then

        Set StageWorksheet = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        StageWorksheet.name = WorksheetName

    End If

    Set EnsureRiskStageWorksheet = StageWorksheet

End Function

Private Sub WriteRiskStageWorksheet( _
    ByVal WorksheetName As String, _
    ByVal TableName As String, _
    ByRef StageData As Variant, _
    ByVal SnapshotDate As Date)

    Dim wsStage As Worksheet
    Dim StageTable As ListObject
    Dim ExistingDataRange As Range
    Dim Headers As Variant
    Dim RowCount As Long
    Dim LastRow As Long
    Dim ColNo As Long
    Dim ExistingLastRow As Long
    Dim ExistingLastCol As Long

    Set wsStage = EnsureRiskStageWorksheet(WorksheetName)
    Headers = RiskStageHeaders()
    RowCount = RiskStageRowCount(StageData)
    ' Keep one blank data row when the stage is empty. This is safer than
    ' resizing an existing ListObject to a header-only range.
    LastRow = 3

    If RowCount > 0 Then LastRow = RowCount + 2

    On Error Resume Next
    Set StageTable = wsStage.ListObjects(TableName)
    On Error GoTo 0

    If StageTable Is Nothing Then

        Do While wsStage.ListObjects.Count > 0

            wsStage.ListObjects(1).Unlist

        Loop

        wsStage.Cells.Clear

    Else

        ' Preserve the ListObject identity so pivots or formulas linked to the
        ' staging table do not lose their source between weekly runs.
        ' Clear the existing body before shrinking the table. Clearing the
        ' worksheet UsedRange while retaining StageTable can invalidate the
        ' live ListObject reference and raise Error 424 (Object required).
        On Error Resume Next

        Set ExistingDataRange = StageTable.DataBodyRange
        ExistingLastRow = _
            StageTable.Range.Row + _
            StageTable.Range.Rows.Count - 1
        ExistingLastCol = _
            StageTable.Range.Column + _
            StageTable.Range.Columns.Count - 1

        On Error GoTo 0

        If Not ExistingDataRange Is Nothing Then

            ExistingDataRange.ClearContents

        End If

        StageTable.Resize _
            wsStage.Range( _
                wsStage.Cells(2, 1), _
                wsStage.Cells(LastRow, RISK_STAGE_FIELD_COUNT))

        If ExistingLastCol > RISK_STAGE_FIELD_COUNT Then

            wsStage.Range( _
                wsStage.Cells(2, RISK_STAGE_FIELD_COUNT + 1), _
                wsStage.Cells(ExistingLastRow, ExistingLastCol)).Clear

        End If

        If ExistingLastRow > LastRow Then

            wsStage.Range( _
                wsStage.Cells(LastRow + 1, 1), _
                wsStage.Cells(ExistingLastRow, RISK_STAGE_FIELD_COUNT)).Clear

        End If

    End If

    wsStage.Range( _
        wsStage.Cells(1, 1), _
        wsStage.Cells(1, 3)).ClearContents

    wsStage.Range( _
        wsStage.Cells(3, 1), _
        wsStage.Cells(LastRow, RISK_STAGE_FIELD_COUNT)).ClearContents

    wsStage.Cells(1, 1).Value = WorksheetName
    wsStage.Cells(1, 2).Value = "As of"

    If SnapshotDate > 0 Then

        wsStage.Cells(1, 3).Value = SnapshotDate

    End If

    For ColNo = 1 To RISK_STAGE_FIELD_COUNT

        wsStage.Cells(2, ColNo).Value = Headers(ColNo - 1)

    Next ColNo

    If RowCount > 0 Then

        wsStage.Range( _
            wsStage.Cells(3, 1), _
            wsStage.Cells(LastRow, RISK_STAGE_FIELD_COUNT)).Value = _
                StageData

    End If

    If StageTable Is Nothing Then

        Set StageTable = _
            wsStage.ListObjects.Add( _
                xlSrcRange, _
                wsStage.Range( _
                    wsStage.Cells(2, 1), _
                    wsStage.Cells(LastRow, RISK_STAGE_FIELD_COUNT)), _
                , _
                xlYes)

        StageTable.name = TableName

    End If

    ' The staging sheet is an audit trail and calculation source. Leave its
    ' existing presentation untouched instead of reformatting up to
    ' thirty thousand rows on every rebuild.

End Sub


Private Sub BuildRiskGranularitySection( _
    ByVal ws As Worksheet, _
    ByRef PositionData As Variant)

    Dim StageRows As Collection
    Dim StageData As Variant
    Dim RiskPositionData As Variant

    Dim CertificateMap As Object
    Dim EquityNameMap As Object
    Dim BondIssuerMap As Object
    Dim BondIssuerTypeMap As Object
    Dim FundMap As Object
    Dim CertificateUnderlyingReferenceIsinMap As Object
    Dim CertificateUnderlyingAssetClassMap As Object
    Dim GeographyEntries As Object
    Dim GeographyObservationSet As Object
    Dim CanonicalNameMap As Object
    Dim CompaniesByName As Object
    Dim CompaniesByVariant As Object
    Dim CountryNameMap As Object
    Dim StageFinalizationCache As Object
    Dim RiskSubtableVisibility As Object


    Dim CertificateMapReady As Boolean
    Dim EquityNameMapReady As Boolean
    Dim BondIssuerMapReady As Boolean
    Dim FundMapReady As Boolean
    Dim CompaniesReady As Boolean
    Dim CertificateNameSameAsLeft As Boolean
    Dim CertificateGeographySameAsLeft As Boolean
    Dim CertificateSectorSameAsLeft As Boolean
    Dim CertificateMappingIssue As String

    Dim r As Long
    Dim SourceRow As Long
    Dim ReportingAssetClass As String
    Dim AssetType As String
    Dim SecurityName As String
    Dim IssuerName As String
    Dim IssuerTicker As String
    Dim BondIssuerType As String
    Dim BondAssetClass As String
    Dim ISIN As String
    Dim AdditionalComment As String
    Dim NDG As String
    Dim PositionValue As Double
    Dim IsDPMPosition As Boolean
    Dim AccountScope As String
    Dim IsExcludedUniCreditBond As Boolean
    Dim ResolvedValue As String
    Dim ResolutionSource As String

    Dim UsedTemporaryCertificateUnderlying As Boolean
    Dim CertificatePositionCount As Long
    Dim TemporaryCertificateCount As Long
    Dim TemporaryCertificateValue As Double

    Dim RiskNextRow As Long
    Dim RiskExDPMNextRow As Long
    Dim GeographyNextRow As Long
    Dim GeographyExDPMNextRow As Long
    Dim SectorNextRow As Long
    Dim SectorExDPMNextRow As Long
    Dim AnalysisDate As Date
    Dim RebuildRiskStage As Boolean

    If ws Is Nothing Then Exit Sub

    Set RiskSubtableVisibility = _
        BuildRiskSubtableVisibility()

    AnalysisDate = CurrentRiskStageAnalysisDate()
    RebuildRiskStage = _
        ShouldRebuildRiskStageTables(AnalysisDate)

    If Not RebuildRiskStage Then

        StageData = _
            LoadRiskStageTableData( _
                RISK_STAGE_SHEET, _
                RISK_STAGE_TABLE)


        GoTo AggregateUnifiedRiskStageDataLabel

    End If

    If Not WeeklyDataHasRows(PositionData) Then Exit Sub

    Set StageRows = NewRiskStageRows()

    RiskStagePositionsScanned = 0
    RiskStageRowsDropped = 0

    Set GeographyEntries = NewExactNameMap()
    Set GeographyObservationSet = NewExactNameMap()

    RiskPositionData = BuildRiskPositionCache(PositionData)


    ' The maintained reference tables are append-only. Their formulas are
    ' calculated before the maps below are loaded for this same run.
    UpdateRiskReferenceDatabases _
        PositionData, _
        RiskPositionData, _
        ThisWorkbook


    Set CertificateMap = _
        LoadCertificateUnderlyingMap( _
            CertificateMapReady, _
            ThisWorkbook, _
            CertificateUnderlyingReferenceIsinMap, _
            CertificateUnderlyingAssetClassMap, _
            CertificateMappingIssue)

    Set EquityNameMap = _
        LoadEquityNameMap( _
            EquityNameMapReady, _
            ThisWorkbook)

    LoadBondIssuerMaps _
        BondIssuerMap, _
        BondIssuerTypeMap, _
        BondIssuerMapReady, _
        ThisWorkbook

    Set FundMap = _
        LoadFundParentCompanyMap( _
            FundMapReady, _
            ThisWorkbook)


    For r = _
        LBound(PositionData, 1) To _
        UBound(PositionData, 1)

        RiskStagePositionsScanned = RiskStagePositionsScanned + 1

        SourceRow = r + 1
        NDG = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosNDG))
        ISIN = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosISIN))
        SecurityName = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosSecurityName))
        AssetType = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosAssetType))
        IssuerName = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosIssuer))
        AdditionalComment = _
            CStr( _
                PositionData( _
                    r, _
                    WeeklyPosAdditionalComment))
        PositionValue = _
            CDbl( _
                PositionData( _
                    r, _
                    WeeklyPosPositionValue))
        IsDPMPosition = _
            CBool( _
                RiskPositionData( _
                    r, _
                    RiskPositionCacheIsDPM))

        If IsDPMPosition Then

            AccountScope = "DPM"

        Else

            AccountScope = "Non-DPM"

        End If

        ReportingAssetClass = _
            CStr( _
                RiskPositionData( _
                    r, _
                    RiskPositionCacheReportingAssetClass))

        Select Case ReportingAssetClass

            Case "Equity"

                IssuerName = SecurityName
                ResolutionSource = "Position: Security Name"

                If IssuerName = "" Then

                    IssuerName = _
                        ResolveExactName( _
                            ISIN, _
                            EquityNameMap)

                    If IssuerName <> "" Then

                        ResolutionSource = "UnmappedEquities: ISIN"

                    Else

                        IssuerName = _
                            BuildTemporaryIssuerName( _
                                ISIN, _
                                SecurityName, _
                                SourceRow, _
                                PlaceholderFromISIN)

                        ResolutionSource = "Fallback"

                    End If

                End If

                AddGeographyLookupEntry _
                    GeographyEntries, _
                    IssuerName, _
                    "Equity issuer", _
                    ISIN, _
                    "Issued security", _
                    GEO_ISIN_PRIORITY_ISSUED_SECURITY, _
                    GeographyObservationSet

                AppendRiskStageRow _
                    StageRows, _
                    SourceRow, _
                    NDG, _
                    ISIN, _
                    SecurityName, _
                    AssetType, _
                    AdditionalComment, _
                    "Equity", _
                    IssuerName, _
                    "Equity issuer", _
                    1, _
                    PositionValue, _
                    ResolutionSource, _
                    AccountScope

            Case "Bonds"

                IssuerTicker = BondIssuerTicker(SecurityName)
                ResolvedValue = _
                    ResolveExactName( _
                        IssuerTicker, _
                        BondIssuerMap)

                If ResolvedValue <> "" Then

                    IssuerName = ResolvedValue
                    ResolutionSource = "BondIssuers: Issuer Ticker"

                Else

                    ResolutionSource = "Position: Issuer"

                End If

                BondIssuerType = ClassifyBondIssuerType(AssetType)

                If BondIssuerType = "" Then

                    BondIssuerType = _
                        ResolveExactName( _
                            IssuerTicker, _
                            BondIssuerTypeMap)

                End If

                BondAssetClass = _
                    BondRiskAssetClass(BondIssuerType)
                IsExcludedUniCreditBond = _
                    IsUniCreditBondIssuer(IssuerName) Or _
                    IsUniCreditBondIssuer(SecurityName)

                If Not IsExcludedUniCreditBond And _
                   IssuerName = "" Then

                    IssuerName = _
                        BuildTemporaryIssuerName( _
                            ISIN, _
                            SecurityName, _
                            SourceRow, _
                            PlaceholderFromSecurityName)

                    ResolutionSource = "Fallback"

                End If

                If Not IsExcludedUniCreditBond Then

                    IsExcludedUniCreditBond = _
                        IsUniCreditBondIssuer(IssuerName)

                End If

                ' Unknown issuer type remains in BondIssuers for completion
                ' and is intentionally outside both bond denominators.
                If Not IsExcludedUniCreditBond And _
                   BondAssetClass <> "" Then

                    AddGeographyLookupEntry _
                        GeographyEntries, _
                        IssuerName, _
                        "Bond issuer", _
                        ISIN, _
                        "Issued security", _
                        GEO_ISIN_PRIORITY_ISSUED_SECURITY, _
                        GeographyObservationSet

                    AppendRiskStageRow _
                        StageRows, _
                        SourceRow, _
                        NDG, _
                        ISIN, _
                        SecurityName, _
                        AssetType, _
                        AdditionalComment, _
                        BondAssetClass, _
                        IssuerName, _
                        "Bond issuer", _
                        1, _
                        PositionValue, _
                        ResolutionSource, _
                        AccountScope

                End If

            Case "Certificates"

                If IsRiskRelevantCertificateAssetType(AssetType) Then

                    CertificatePositionCount = _
                        CertificatePositionCount + 1

                    AppendCertificateRiskStageRows _
                        StageRows, _
                        CertificateMap, _
                        FundMap, _
                        CertificateUnderlyingAssetClassMap, _
                        ISIN, _
                        SecurityName, _
                        AssetType, _
                        AdditionalComment, _
                        NDG, _
                        PositionValue, _
                        SourceRow, _
                        AccountScope, _
                        UsedTemporaryCertificateUnderlying

                    AddCertificateGeographyEntries _
                        GeographyEntries, _
                        CertificateMap, _
                        CertificateUnderlyingReferenceIsinMap, _
                        CertificateUnderlyingAssetClassMap, _
                        FundMap, _
                        ISIN, _
                        SecurityName, _
                        GeographyObservationSet

                    If UsedTemporaryCertificateUnderlying Then

                        TemporaryCertificateCount = _
                            TemporaryCertificateCount + 1
                        TemporaryCertificateValue = _
                            TemporaryCertificateValue + PositionValue

                    End If

                End If

            Case "Funds"

                IssuerName = _
                    ResolveExactName( _
                        SecurityName, _
                        FundMap)

                If IssuerName <> "" Then

                    ResolutionSource = "Funds: Fund Name"

                Else

                    IssuerName = _
                        ResolveFundParentCompany( _
                            FundMap, _
                            ISIN, _
                            SecurityName, _
                            SourceRow)

                    ResolutionSource = "Fallback"

                End If

                AddGeographyLookupEntry _
                    GeographyEntries, _
                    IssuerName, _
                    "Fund parent company", _
                    ISIN, _
                    "Managed fund", _
                    GEO_ISIN_PRIORITY_MANAGED_FUND, _
                    GeographyObservationSet

                AppendRiskStageRow _
                    StageRows, _
                    SourceRow, _
                    NDG, _
                    ISIN, _
                    SecurityName, _
                    AssetType, _
                    AdditionalComment, _
                    "Funds", _
                    IssuerName, _
                    "Fund parent company", _
                    1, _
                    PositionValue, _
                    ResolutionSource, _
                    AccountScope

        End Select

    Next r


    Set CanonicalNameMap = _
        BuildCanonicalEntityNameMap(GeographyEntries)


    UpdateNameVariantsWorksheet _
        GeographyEntries, _
        CanonicalNameMap


    Set GeographyEntries = _
        CanonicalizeGeographyEntries( _
            GeographyEntries, _
            CanonicalNameMap)


    LoadCompaniesLookup _
        CompaniesByName, _
        CompaniesByVariant, _
        CompaniesReady


    Set GeographyEntries = _
        CanonicalizeGeographyUsingCompanies( _
            GeographyEntries, _
            CompaniesByName, _
            CompaniesByVariant)


    Set CountryNameMap = LoadCountryNameMap(ThisWorkbook)
    Set StageFinalizationCache = NewExactNameMap()


    AddRiskStagingNote StageRows

    StageData = _
        FinalizeRiskStageData( _
            StageRows, _
            CanonicalNameMap, _
            CompaniesByName, _
            CompaniesByVariant, _
            CountryNameMap, _
            StageFinalizationCache)


    ' The unified table is the audit trail and the sole source for all risk
    ' tables below. Each certificate component is already allocated to its
    ' final weight; Account Scope separates Non-DPM and DPM rows.
    WriteRiskStageWorksheet _
        RISK_STAGE_SHEET, _
        RISK_STAGE_TABLE, _
        StageData, _
        AnalysisDate


    DeleteLegacyRiskStageWorksheets


AggregateUnifiedRiskStageDataLabel:

    '
    ' Reported whether the staging table was rebuilt or reused, because the
    ' condition belongs to the data being reported, not to how it got here.
    '
    AddUnknownUnderlyingNote StageData

    '
    ' Every table below reads the staging table through its own formulas, so
    ' the only thing to settle first is whether a right-hand Certificate
    ' subtable would repeat the left-hand one - the case
    ' WriteSameAsLeftExposureGroup answers with a single row.
    '
    CertificateNameSameAsLeft = _
        CertificateResultsMatch("Issuer", RiskSubtableVisibility)

    CertificateGeographySameAsLeft = _
        CertificateResultsMatch("Country", RiskSubtableVisibility)

    CertificateSectorSameAsLeft = _
        CertificateResultsMatch("Sector", RiskSubtableVisibility)


    RiskNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.RiskRow, _
            Layout.RiskCol, _
            "Name Concentration - Top 10", _
            "Name", _
            "Issuer", _
            False, _
            RiskSubtableVisibility)

    RiskExDPMNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.RiskExSegRow, _
            Layout.RiskExSegCol, _
            "Name Concentration - Top 10 " & _
            "(Excl. DPM)", _
            "Name", _
            "Issuer", _
            True, _
            RiskSubtableVisibility, _
            CertificateNameSameAsLeft)

    Layout.CountryRiskRow = _
        NextRiskSectionRow( _
            RiskNextRow, _
            RiskExDPMNextRow)
    Layout.CountryRiskExSegRow = Layout.CountryRiskRow

    GeographyNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.CountryRiskRow, _
            Layout.CountryRiskCol, _
            "Geographic Concentration - Top 10", _
            "Geography", _
            "Country", _
            False, _
            RiskSubtableVisibility)

    GeographyExDPMNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.CountryRiskExSegRow, _
            Layout.CountryRiskExSegCol, _
            "Geographic Concentration - Top 10 " & _
            "(Excl. DPM)", _
            "Geography", _
            "Country", _
            True, _
            RiskSubtableVisibility, _
            CertificateGeographySameAsLeft)

    Layout.SectorRiskRow = _
        NextRiskSectionRow( _
            GeographyNextRow, _
            GeographyExDPMNextRow)
    Layout.SectorRiskExSegRow = Layout.SectorRiskRow

    SectorNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.SectorRiskRow, _
            Layout.SectorRiskCol, _
            "Sector Concentration - Top 10", _
            "Sector", _
            "Sector", _
            False, _
            RiskSubtableVisibility)

    SectorExDPMNextRow = _
        WriteRiskDimensionTable( _
            ws, _
            Layout.SectorRiskExSegRow, _
            Layout.SectorRiskExSegCol, _
            "Sector Concentration - Top 10 " & _
            "(Excl. DPM)", _
            "Sector", _
            "Sector", _
            True, _
            RiskSubtableVisibility, _
            CertificateSectorSameAsLeft)


    If RebuildRiskStage Then

        WriteNewGeoSecLookupWorksheet _
            GeographyEntries, _
            CompaniesByName, _
            CompaniesByVariant

        If Not CompaniesReady Then

            WriteNoteWeekly _
                "The '" & COMPANIES_SHEET & _
                "' master-data worksheet was not available. " & _
                "Current entities were added to '" & _
                GEO_SEC_LOOKUP_SHEET & _
                "'; geography and sector exposures are shown under " & _
                OTHER_RISK_DIMENSION & "."

        End If

        AddCertificateMappingNotes _
            CertificatePositionCount, _
            CertificateMapReady, _
            CertificateMappingIssue, _
            TemporaryCertificateCount, _
            TemporaryCertificateValue


    End If


End Sub

'
' An exposure the certificate expansion could not name.
'
Private Function IsUnknownUnderlyingType( _
    ByVal ExposureType As String) As Boolean

    IsUnknownUnderlyingType = _
        StrComp( _
            Trim$(ExposureType), _
            UNKNOWN_UNDERLYING_TYPE, _
            vbTextCompare) = 0

End Function

'
' A certificate whose underlying could not be resolved to anything nameable
' carries its whole value into the Issuer share denominator under a name no
' one can act on, and used to say nothing about it.  This names the
' certificates, so the reference data behind them can be fixed.
'
Private Sub AddUnknownUnderlyingNote( _
    ByRef StageData As Variant)

    Dim Certificates As Object
    Dim CertificateKey As Variant

    Dim RowCount As Long
    Dim TotalValue As Double
    Dim Listed As Long
    Dim Detail As String

    Dim RowNo As Long

    If Not IsArray(StageData) Then Exit Sub
    If UBound(StageData, 1) < 1 Then Exit Sub

    Set Certificates = CreateObject("Scripting.Dictionary")
    Certificates.CompareMode = vbTextCompare

    For RowNo = 1 To UBound(StageData, 1)

        If IsUnknownUnderlyingType( _
                SafeText(StageData(RowNo, RiskStageExposureType))) Then

            RowCount = RowCount + 1

            TotalValue = _
                TotalValue + _
                CDbl(StageData(RowNo, RiskStageAllocatedValue))

            Certificates( _
                SafeText(StageData(RowNo, RiskStageSecurityName)) & _
                " (" & _
                SafeText(StageData(RowNo, RiskStageProductISIN)) & ")") = True

        End If

    Next RowNo

    If RowCount = 0 Then Exit Sub

    For Each CertificateKey In Certificates.Keys

        If Listed >= UNKNOWN_UNDERLYING_NOTE_LIMIT Then

            Detail = _
                Detail & vbLf & "and " & _
                CStr(Certificates.Count - Listed) & " more"

            Exit For

        End If

        Detail = Detail & vbLf & CStr(CertificateKey)
        Listed = Listed + 1

    Next CertificateKey

    WriteNoteWeekly _
        CStr(Certificates.Count) & _
        " Certificate position(s) (excl. Protected) have an underlying " & _
        "that could not be identified at all. Their collateral value is " & _
        "counted in the Certificates total and in the Issuer share " & _
        "denominator, but they can never appear in a Top 10 by name." & _
        vbLf & _
        "Collateral value: " & Format(TotalValue, EuroNumberFormat()) & _
        " over " & CStr(RowCount) & " staging row(s)." & _
        Detail

End Sub

'
' What the staging pass made of the positions it read.  Nothing is said when
' every position produced an exposure, which is the normal case; the report
' hears about it only when rows were dropped for want of an exposure name, or
' when the pass produced no rows at all.  Those two look identical afterwards
' - an empty staging table and a report of Nones - and neither used to say
' anything at all.
'
Private Sub AddRiskStagingNote( _
    ByVal StageRows As Collection)

    Dim StagedRows As Long

    If Not StageRows Is Nothing Then StagedRows = StageRows.Count

    If StagedRows > 0 And RiskStageRowsDropped = 0 Then Exit Sub

    WriteNoteWeekly _
        "Risk staging read " & CStr(RiskStagePositionsScanned) & _
        " position(s) and produced " & CStr(StagedRows) & " row(s), " & _
        "dropping " & CStr(RiskStageRowsDropped) & _
        " for which no exposure name could be resolved." & vbLf & _
        "Exposure names come from the maintained reference tables, so a run " & _
        "made before those tables have finished calculating resolves none " & _
        "of them."

End Sub

Private Sub AddCertificateMappingNotes( _
    ByVal CertificatePositionCount As Long, _
    ByVal CertificateMapReady As Boolean, _
    ByVal MappingIssue As String, _
    ByVal TemporaryCertificateCount As Long, _
    ByVal TemporaryCertificateValue As Double)

    If CertificatePositionCount = 0 Then Exit Sub

    If Not CertificateMapReady Then

        WriteNoteWeekly _
            "Certificate underlying mapping is unavailable " & _
            "(excl. Protected)." & _
            vbLf & _
            MappingIssue

    End If

    If TemporaryCertificateCount > 0 Then

        WriteNoteWeekly _
            CStr(TemporaryCertificateCount) & _
            " Certificate position(s) (excl. Protected) could not be " & _
            "resolved " & _
            "to underlying " & _
            "company names and are temporarily shown as [Security Name] " & _
            "or [ISIN]." & vbLf & _
            "Collateral value: " & _
            Format( _
                TemporaryCertificateValue, _
                EuroNumberFormat())

    End If

    If CertificateMapReady And MappingIssue <> "" Then

        WriteNoteWeekly _
            "Certificate reference data is incomplete " & _
            "(excl. Protected)." & _
            vbLf & _
            MappingIssue

    End If

End Sub


Private Function WriteTopExposureGroup( _
    ByVal ws As Worksheet, _
    ByVal StartRow As Long, _
    ByVal LeftCol As Long, _
    ByVal DisplayName As String, _
    ByVal DimensionKey As String, _
    ByVal AssetClassKey As String, _
    ByVal ExcludeDPM As Boolean, _
    ByVal Visibility As Object) As Long

    Dim AnchorCell As Range

    Dim OutputCount As Long
    Dim FirstDataRow As Long
    Dim TotalRow As Long
    Dim i As Long

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4))

        .Merge
        .Value = DisplayName

    End With

    FirstDataRow = StartRow + 1

    '
    ' One formula produces the whole ranked table - name, value, share and
    ' distinct NDG count - and spills it across the four columns beside the
    ' rank.  Its height is however many names the class has, up to ten, so
    ' it is calculated here to find out where the total row goes.
    '
    Set AnchorCell = ws.Cells(FirstDataRow, LeftCol + 1)

    AnchorCell.Formula2 = _
        RiskRankedFormula( _
            DimensionKey, _
            AssetClassKey, _
            ExcludeDPM, _
            Visibility)

    AnchorCell.Calculate

    OutputCount = RiskSpilledRowCount(AnchorCell)

    If OutputCount = 0 Then

        AnchorCell.ClearContents

        OutputCount = 1
        ws.Cells(FirstDataRow, LeftCol + 1).Value = "None"
        ws.Cells(FirstDataRow, LeftCol + 4).Value = 0

    Else

        For i = 1 To OutputCount

            ws.Cells(FirstDataRow + i - 1, LeftCol).Value = i

        Next i

    End If

    TotalRow = FirstDataRow + OutputCount

    ws.Cells(TotalRow, LeftCol + 2).Formula2 = _
        "=SUM(" & _
        ws.Range( _
            ws.Cells(FirstDataRow, LeftCol + 2), _
            ws.Cells(TotalRow - 1, LeftCol + 2)).Address(True, True) & ")"

    ws.Cells(TotalRow, LeftCol + 4).Formula2 = _
        RiskUnionNdgFormula( _
            DimensionKey, _
            AssetClassKey, _
            ExcludeDPM, _
            Visibility, _
            ws.Range( _
                ws.Cells(FirstDataRow, LeftCol + 1), _
                ws.Cells(TotalRow - 1, LeftCol + 1)).Address(True, True))

    '
    ' The share the report has always shown: the top ten against the whole
    ' category.  IFERROR keeps an empty category at nought rather than
    ' #DIV/0!, which is what the division used to be guarded for.
    '
    ws.Cells(TotalRow, LeftCol + 3).Formula2 = _
        "=IFERROR(" & _
        ws.Cells(TotalRow, LeftCol + 2).Address(True, True) & " / " & _
        RiskCategoryTotalFormula( _
            DimensionKey, _
            AssetClassKey, _
            ExcludeDPM, _
            Visibility) & ", 0)"

    FormatReportTable _
        ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(TotalRow, LeftCol + 4)), _
        1

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4))

        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(111, 38, 61)

    End With

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4)) _
            .Borders(xlEdgeBottom)

        .LineStyle = xlNone

    End With

    With ws.Range( _
            ws.Cells(StartRow, LeftCol), _
            ws.Cells(StartRow, LeftCol + 4)) _
            .Borders(xlEdgeTop)

        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(111, 38, 61)

    End With

    ws.Range( _
        ws.Cells(FirstDataRow, LeftCol + 1), _
        ws.Cells(TotalRow, LeftCol + 1)).HorizontalAlignment = _
        xlLeft

    With ws.Range( _
            ws.Cells(TotalRow, LeftCol), _
            ws.Cells(TotalRow, LeftCol + 4))

        .Interior.Color = RGB(255, 255, 255)

    End With

    FormatTotalRow _
        ws, _
        LeftCol, _
        LeftCol + 4, _
        TotalRow

    ws.Range( _
        ws.Cells(FirstDataRow, LeftCol + 2), _
        ws.Cells(TotalRow, LeftCol + 2)).NumberFormat = _
        EuroNumberFormat()

    ws.Range( _
        ws.Cells(FirstDataRow, LeftCol + 3), _
        ws.Cells(TotalRow, LeftCol + 3)).NumberFormat = _
        "0.00%"

    ws.Range( _
        ws.Cells(FirstDataRow, LeftCol + 4), _
        ws.Cells(TotalRow, LeftCol + 4)).NumberFormat = _
        "0"

    WriteTopExposureGroup = TotalRow + 1

End Function


Private Sub CreateCollateralPieChart( _
    ByVal ws As Worksheet, _
    ByVal DictCurrent As Object, _
    ByVal ReportDate As Date)

    Dim ChartObj As ChartObject

    Dim TotalCollateral As Double

    Dim TopPos As Double
    Dim LeftPos As Double
    Dim RightPos As Double

    Dim BreakdownRow As Long
    Dim BreakdownCol As Long

    Dim SliceColors As Variant
    Dim i As Long

    BreakdownRow = Layout.BreakdownRow
    BreakdownCol = Layout.BreakdownCol

    '
    ' Total Collateral
    '

    TotalCollateral = CollateralTotal(DictCurrent)


    '
    ' Delete old chart
    '

    On Error Resume Next

    ws.ChartObjects("CollateralPie").Delete

    On Error GoTo 0

    '
    ' Position
    '

    TopPos = ws.Rows(Layout.PieRow).Top

    LeftPos = ws.Columns(Layout.PieCol).Left

    RightPos = _
        ws.Columns(Layout.EnteredCol + 7).Left + _
        ws.Columns(Layout.EnteredCol + 7).Width

    '
    ' Create Chart
    '
    
    Dim PieHeight As Double

    PieHeight = 0
    
    For i = Layout.PieRow To _
             Layout.PieRow + Layout.PieHeightRows - 1
    
        PieHeight = PieHeight + _
                    ws.Rows(i).Height
    
    Next i

    Set ChartObj = ws.ChartObjects.Add( _
        Left:=LeftPos, _
        Top:=TopPos, _
        Width:=RightPos - LeftPos, _
        Height:=PieHeight)

    ChartObj.name = "CollateralPie"

    With ChartObj.Chart

        .ChartType = xlPie

        Do While .SeriesCollection.Count > 0

            .SeriesCollection(1).Delete

        Loop

        .SeriesCollection.NewSeries

        '
        ' Categories
        '

        .SeriesCollection(1).XValues = _
            ws.Range( _
                ws.Cells(BreakdownRow + 1, BreakdownCol + 1), _
                ws.Cells( _
                    BreakdownRow + 1, _
                    BreakdownCol + CollateralCategoryCount()))

        '
        ' Current Amounts
        '

        .SeriesCollection(1).Values = _
            ws.Range( _
                ws.Cells(BreakdownRow + 5, BreakdownCol + 1), _
                ws.Cells( _
                    BreakdownRow + 5, _
                    BreakdownCol + CollateralCategoryCount()))

        '
        ' Clean look
        '

        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse

        '
        ' Title
        '

        .HasTitle = True

        .ChartTitle.Text = _
            "Collateral Breakdown" & _
            vbLf & _
            "As of " & _
            Format(ReportDate, "dd/mm/yyyy")

        With .ChartTitle.Format.TextFrame2.TextRange.Font

            .Size = 14
            .Bold = msoTrue
            .name = "Aptos Display"

        End With

        '
        ' Move pie slightly down
        '

        On Error Resume Next

        .PlotArea.Top = .PlotArea.Top + 15

        On Error GoTo 0

        '
        ' Legend
        '

        .Legend.Position = xlLegendPositionRight

        '
        ' Labels
        '

        .ApplyDataLabels

        With .SeriesCollection(1)

            .DataLabels.ShowPercentage = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowValue = False

            .DataLabels.NumberFormat = "0.0%"

            On Error Resume Next

            .DataLabels.Position = _
                xlLabelPositionBestFit

            On Error GoTo 0

        End With
        
        With .SeriesCollection(1).DataLabels

            .Font.name = "Aptos Display"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(40, 40, 40)
        
        End With
        
        With .Legend.Font
        
            .name = "Aptos Display"
            .Size = 9
        
        End With

    End With

    SliceColors = CollateralSliceColors()

    For i = 1 To _
        ChartObj.Chart.SeriesCollection(1).Points.Count

        With ChartObj.Chart.SeriesCollection(1).Points(i)

            .Format.Fill.ForeColor.RGB = _
                SliceColors(i - 1)

            '
            ' White separators
            '
            
            .Format.Line.Visible = msoTrue
            
            .Format.Line.ForeColor.RGB = _
                RGB(255, 255, 255)
            
            .Format.Line.Weight = 0.75
            
            .Format.Line.DashStyle = _
                msoLineDash
                
        End With

    Next i

    '
    ' Total Collateral Card
    '

    With ChartObj.Chart.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            ChartObj.Width - 230, _
            ChartObj.Height - 55, _
            210, _
            35)
    
        With .TextFrame
    
            .Characters.Text = _
                "Total Collateral Value:" & vbLf & _
                Application.WorksheetFunction.Text( _
                    TotalCollateral, _
                    EuroNumberFormat())
    
            .HorizontalAlignment = xlRight
            
            .Characters(1, Len("Total Collateral Value:")). _
                Font.Bold = True
                
            .Characters( _
                Len("Total Collateral Value:") + 2). _
                Font.Bold = False
    
            .Characters.Font.name = "Aptos Display"
    
            .Characters.Font.Size = 11
    
        End With
    
        .Line.Visible = msoFalse
    
        .Fill.Visible = msoTrue
    
        .Fill.ForeColor.RGB = RGB(245, 245, 245)
    
    End With

End Sub

'
' Public because nothing calls it by name: GenerateWeeklyAnalysis puts
' "WriteNoteWeekly" into the NoteHandler global and Note reaches it
' through Application.Run, which cannot see a Private procedure.
'
Public Sub WriteNoteWeekly( _
    ByVal Message As String)

    If ReportNotes Is Nothing Then

        Set ReportNotes = New Collection

    End If

    If Trim(Message) = "" Then Exit Sub

    ReportNotes.Add Message

End Sub


Private Sub BuildNotes( _
    ByVal ws As Worksheet)

    Dim FirstRow As Long
    Dim LastRow As Long
    Dim FirstCol As Long
    Dim LastCol As Long
    Dim Item As Variant
    Dim Parts() As String
    Dim NotesText As String
    Dim NoteText As String
    Dim i As Long

    FirstRow = Layout.CommentRow + 1
    LastRow = Layout.PieRow + _
              Layout.PieHeightRows - 1
    FirstCol = Layout.CommentCol
    LastCol = Layout.CommentCol + 3

    WriteSectionTitle _
        ws, _
        Layout.CommentRow, _
        FirstCol, _
        4, _
        "Notes"

    ws.Range( _
        ws.Cells(Layout.CommentRow, FirstCol), _
        ws.Cells(Layout.CommentRow, LastCol)).Font.name = _
            "Aptos Display"

    If ReportNotes Is Nothing Then

        NotesText = "None"

    ElseIf ReportNotes.Count = 0 Then

        NotesText = "None"

    Else

        For Each Item In ReportNotes

            Parts = Split(CStr(Item), vbLf)
            NoteText = ChrW(8226) & " " & Parts(0)

            For i = 1 To UBound(Parts)

                NoteText = _
                    NoteText & vbLf & _
                    "    " & Parts(i)

            Next i

            If NotesText <> "" Then

                NotesText = NotesText & vbLf & vbLf

            End If

            NotesText = NotesText & NoteText

        Next Item

    End If

    With ws.Range( _
            ws.Cells(FirstRow, FirstCol), _
            ws.Cells(LastRow, LastCol))

        .UnMerge
        .ClearContents
        .Merge
        .Value = NotesText
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Bold = False
        .Font.name = "Aptos Display"
        .Font.Size = 10

    End With

End Sub

Private Sub FormatNotesBox( _
    ByVal ws As Worksheet)

    Dim FirstRow As Long
    Dim LastRow As Long

    Dim FirstCol As Long
    Dim LastCol As Long

    FirstRow = Layout.CommentRow + 1

    LastRow = Layout.PieRow + _
              Layout.PieHeightRows - 1

    FirstCol = Layout.CommentCol

    LastCol = Layout.CommentCol + 3

    ws.Range( _
        ws.Cells(FirstRow, 1), _
        ws.Cells(LastRow, 1)).EntireRow.RowHeight = 16

    With ws.Range( _
        ws.Cells(FirstRow, FirstCol), _
        ws.Cells(LastRow, LastCol))

        If Not .MergeCells Then .Merge

        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.name = "Aptos Display"
        .Font.Size = 10

        .BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlMedium, _
            Color:=RGB(60, 60, 60)

    End With

End Sub

Private Sub CreateWeeklyEmailButton(ByVal ws As Worksheet)

    Dim btn As Button
    Dim Btn2 As Button

    On Error Resume Next
    ws.Buttons("btnWeeklyEmail").Delete
    ws.Buttons("btnWeeklyRerun").Delete
    On Error GoTo 0

    Set btn = ws.Buttons.Add(345, 16, 100, 26)
    btn.name = "btnWeeklyEmail"
    btn.Characters.Text = "Generate Email"
    btn.OnAction = "CreateWeeklyEmail"

    Set Btn2 = ws.Buttons.Add(455, 16, 50, 26)
    Btn2.name = "btnWeeklyRerun"
    Btn2.Characters.Text = "Rerun"
    Btn2.OnAction = "GenerateWeeklyAnalysis"

End Sub


