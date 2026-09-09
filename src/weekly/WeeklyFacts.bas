Attribute VB_Name = "WeeklyFacts"
Option Explicit

'
' Weekly Facts: one sheet of figures and superlatives about the book as of
' one date - the largest and smallest clients, the biggest positions and
' who holds them, what moved and by how much, and what the whole run of
' snapshots on file says.  Its own button on Home; it goes into no email.
'
' The one input is the end date, Home!FactsEndDate, or Home!WeeklyEndDate
' when that name does not exist.  A date with no snapshot is read as the
' last snapshot on or before it.  Every comparison is made three ways,
' each to a snapshot on file - the previous snapshot, the month-earlier
' one and the year's first, which are the finest window the files allow
' and the report's own two - and the history section walks every
' Accounts snapshot up to the end date.
'
' Reads the snapshots itself, through the weekly module's field cleaner
' and number parser, so its figures agree with the report's.
'

Private Const FACTS_SHEET As String = "Weekly Facts"
Private Const FACTS_DATE_NAME As String = "FactsEndDate"
Private Const REPORT_DATE_NAME As String = "WeeklyEndDate"

Private Const ACCOUNT_FILE_SUFFIX As String = "_Lombard_Loans_ITA_Accounts.csv"
Private Const POSITION_FILE_SUFFIX As String = "_Lombard_Loans_ITA_Positions.csv"

Private Const FACT_DATE_FORMAT As String = "dd/mm/yyyy"
Private Const FACT_TOLERANCE As Double = 0.005
Private Const FIRST_COL As Long = 2

'
' One Accounts row.  The two flags come in as numbers, one for set.
'
Private Enum FactAccountField
    FactAccNDG = 1
    FactAccApproved = 2
    FactAccDrawn = 3
    FactAccMTM = 4
    FactAccHCV = 5
    FactAccMarginCall = 6
    FactAccShortfall = 7
End Enum

Private Const FACT_ACCOUNT_FIELDS As Long = 7

'
' One Positions row, with the class GetAssetClass gives its asset type.
'
Private Enum FactPositionField
    FactPosNDG = 1
    FactPosISIN = 2
    FactPosName = 3
    FactPosAssetType = 4
    FactPosClass = 5
    FactPosCurrency = 6
    FactPosIssuer = 7
    FactPosValue = 8
    FactPosHCV = 9
    FactPosAboveLimit = 10
End Enum

Private Const FACT_POSITION_FIELDS As Long = 10

'
' A date's two snapshots and the sums drawn from its positions.
'
Private Type FactSnapshot
    AsOfDate As Date
    Accounts As Variant
    Positions As Variant
    Aggregates As Object
End Type

Private FactsSheet As Worksheet
Private FactsRow As Long
Private FactsNotes As Collection

'====================================================================
' Entry
'====================================================================

Public Sub GenerateWeeklyFacts()

    Dim RequestedDate As Date
    Dim EndDate As Date
    Dim Dates As Variant
    Dim EndSnap As FactSnapshot

    On Error GoTo ErrorHandler

    If Not ReadFactsEndDate(RequestedDate) Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    MissingFiles = ""
    Set FactsNotes = New Collection

    Dates = AccountSnapshotDates(RequestedDate)

    If IsEmpty(Dates) Then
        Fatal _
            "No Accounts snapshot on file on or before " & _
            Format(RequestedDate, FACT_DATE_FORMAT) & "."
    End If

    EndDate = Dates(UBound(Dates))

    If EndDate <> RequestedDate Then
        AddFactsNote _
            "No snapshot for " & Format(RequestedDate, FACT_DATE_FORMAT) & _
            ": the last one on or before it, " & _
            Format(EndDate, FACT_DATE_FORMAT) & ", is used."
    End If

    If Not SourceFileExists(EndDate, "POSITIONS") Then
        Fatal _
            "No Positions snapshot on file for " & _
            Format(EndDate, FACT_DATE_FORMAT) & "."
    End If

    Set FactsSheet = CreateOrReplaceSheet(FACTS_SHEET)

    If FactsSheet Is Nothing Then GoTo ExitRoutine

    EndSnap = LoadFactSnapshot(EndDate)

    WriteFactsHeader EndDate
    WritePortfolioSection EndSnap
    WriteClientSection EndSnap
    WritePositionSection EndSnap

    WriteHorizon EndSnap, PreviousSnapshotDate(Dates, EndDate), _
        "Since the previous snapshot"
    WriteHorizon EndSnap, _
        SnapshotOnOrAfter(Dates, GetComparisonDate(EndDate), EndDate), _
        "Since a month earlier"
    WriteHorizon EndSnap, _
        SnapshotOnOrAfter(Dates, GetYTDDate(EndDate), EndDate), _
        "Since the year started"

    WriteHistorySection Dates, EndSnap

    FinishFactsSheet

    If MissingFiles <> "" Then
        MsgBox _
            "Weekly Facts completed with warnings." & vbCrLf & vbCrLf & _
            "Missing source files:" & vbCrLf & MissingFiles, _
            vbExclamation, "Weekly Facts"
    End If

ExitRoutine:

    ResetExcel

    If Not FactsSheet Is Nothing Then FactsSheet.Activate

    Set FactsSheet = Nothing
    Set FactsNotes = Nothing

    Exit Sub

ErrorHandler:

    ResetExcel

    MsgBox Err.Description, vbCritical, "Weekly Facts"

    GoTo ExitRoutine

End Sub

'
' The end date from Home: FactsEndDate, or the report's own date when
' that name has not been created.  False, with a message, when neither
' holds a date.
'
Private Function ReadFactsEndDate( _
    ByRef EndDate As Date) As Boolean

    Dim RawValue As Variant
    Dim NameUsed As String

    NameUsed = FACTS_DATE_NAME

    On Error Resume Next

    RawValue = ThisWorkbook.Worksheets("Home").Range(FACTS_DATE_NAME).Value

    If Err.Number <> 0 Then

        Err.Clear
        NameUsed = REPORT_DATE_NAME
        RawValue = ThisWorkbook.Worksheets("Home").Range(REPORT_DATE_NAME).Value

    End If

    On Error GoTo 0

    If Not IsDate(RawValue) Then

        MsgBox _
            "Home!" & NameUsed & " does not hold a valid date." & _
            vbCrLf & vbCrLf & _
            "Name a cell on Home " & FACTS_DATE_NAME & _
            " with the date the facts are wanted for.", _
            vbExclamation, "Weekly Facts"

        Exit Function

    End If

    EndDate = CDate(RawValue)
    ReadFactsEndDate = True

End Function

Private Sub AddFactsNote( _
    ByVal Message As String)

    FactsNotes.Add Message

End Sub

'====================================================================
' The snapshots on file
'====================================================================

'
' Every Accounts snapshot in the source folder dated on or before the
' date given, ascending; Empty when there is none.
'
Private Function AccountSnapshotDates( _
    ByVal LastDate As Date) As Variant

    Dim BasePath As String
    Dim FileName As String
    Dim Found As Collection
    Dim Dates() As Date
    Dim Candidate As Date
    Dim Swap As Date
    Dim i As Long
    Dim j As Long

    BasePath = PathSelection()
    Set Found = New Collection

    FileName = Dir(BasePath & "*" & ACCOUNT_FILE_SUFFIX)

    Do While FileName <> ""

        If SnapshotDateFromFileName(FileName, Candidate) Then
            If Candidate <= LastDate Then Found.Add Candidate
        End If

        FileName = Dir

    Loop

    If Found.Count = 0 Then Exit Function

    ReDim Dates(0 To Found.Count - 1)

    For i = 1 To Found.Count
        Dates(i - 1) = Found(i)
    Next i

    For i = 0 To UBound(Dates) - 1
        For j = i + 1 To UBound(Dates)
            If Dates(j) < Dates(i) Then
                Swap = Dates(i)
                Dates(i) = Dates(j)
                Dates(j) = Swap
            End If
        Next j
    Next i

    AccountSnapshotDates = Dates

End Function

'
' The yyyymmdd a source file starts with, as a date; False when the name
' does not start that way.
'
Private Function SnapshotDateFromFileName( _
    ByVal FileName As String, _
    ByRef SnapshotDate As Date) As Boolean

    Dim Code As String

    Code = Left$(FileName, 8)

    If Len(Code) < 8 Then Exit Function
    If Not IsNumeric(Code) Then Exit Function

    On Error GoTo BadDate

    SnapshotDate = _
        DateSerial( _
            CLng(Left$(Code, 4)), _
            CLng(Mid$(Code, 5, 2)), _
            CLng(Mid$(Code, 7, 2)))

    SnapshotDateFromFileName = True

    Exit Function

BadDate:

    Err.Clear

End Function

'
' The last snapshot before the end date that has positions too; zero when
' there is none.
'
Private Function PreviousSnapshotDate( _
    ByRef Dates As Variant, _
    ByVal EndDate As Date) As Date

    Dim i As Long

    For i = UBound(Dates) To LBound(Dates) Step -1

        If Dates(i) < EndDate Then
            If SourceFileExists(Dates(i), "POSITIONS") Then
                PreviousSnapshotDate = Dates(i)
                Exit Function
            End If
        End If

    Next i

End Function

'
' The first snapshot on or after the target and before the end date that
' has positions too - the way the report resolves its month and year
' dates; zero when there is none.
'
Private Function SnapshotOnOrAfter( _
    ByRef Dates As Variant, _
    ByVal Target As Date, _
    ByVal EndDate As Date) As Date

    Dim i As Long

    For i = LBound(Dates) To UBound(Dates)

        If Dates(i) >= Target And Dates(i) < EndDate Then
            If SourceFileExists(Dates(i), "POSITIONS") Then
                SnapshotOnOrAfter = Dates(i)
                Exit Function
            End If
        End If

    Next i

End Function

'====================================================================
' Loading a snapshot
'====================================================================

Private Function LoadFactSnapshot( _
    ByVal SnapshotDate As Date) As FactSnapshot

    Dim Snap As FactSnapshot

    Snap.AsOfDate = SnapshotDate
    Snap.Accounts = LoadFactAccounts(SnapshotDate)
    Snap.Positions = LoadFactPositions(SnapshotDate)
    Set Snap.Aggregates = BuildPositionAggregates(Snap.Positions)

    LoadFactSnapshot = Snap

End Function

Private Function FactSourcePath( _
    ByVal SnapshotDate As Date, _
    ByVal FileSuffix As String) As String

    FactSourcePath = PathSelection() & GetDateCode(SnapshotDate) & FileSuffix

End Function

'
' A header's position among the file's, matched with spaces and
' punctuation ignored, trying each candidate in turn; -1 when none is
' there, which raises for a required column.
'
Private Function FactHeaderIndex( _
    ByRef HeaderFields As Variant, _
    ByVal Candidates As Variant, _
    ByVal Required As Boolean, _
    ByVal FilePath As String) As Long

    Dim Candidate As Variant
    Dim i As Long

    FactHeaderIndex = -1

    For Each Candidate In Candidates

        For i = LBound(HeaderFields) To UBound(HeaderFields)

            If NormalizeFactHeader(CleanWeeklyCsvField(HeaderFields(i))) = _
               NormalizeFactHeader(CStr(Candidate)) Then

                FactHeaderIndex = i

                Exit Function

            End If

        Next i

    Next Candidate

    If Required Then
        Err.Raise _
            vbObjectError + 1950, _
            "FactHeaderIndex", _
            "Column '" & CStr(Candidates(LBound(Candidates))) & _
            "' could not be found in:" & vbCrLf & FilePath
    End If

End Function

Private Function NormalizeFactHeader( _
    ByVal HeaderText As String) As String

    Dim Result As String
    Dim Ch As String
    Dim i As Long

    HeaderText = UCase$(HeaderText)

    For i = 1 To Len(HeaderText)

        Ch = Mid$(HeaderText, i, 1)

        If (Ch >= "A" And Ch <= "Z") Or (Ch >= "0" And Ch <= "9") Then
            Result = Result & Ch
        End If

    Next i

    NormalizeFactHeader = Result

End Function

Private Function FactField( _
    ByRef Fields As Variant, _
    ByVal FieldIndex As Long) As String

    If FieldIndex < LBound(Fields) Or FieldIndex > UBound(Fields) Then Exit Function

    FactField = CleanWeeklyCsvField(Fields(FieldIndex))

End Function

'
' A flag column's value: the words for yes, or any number above zero.
'
Private Function FactFlag( _
    ByVal FieldText As String) As Boolean

    Select Case UCase$(Trim$(FieldText))

        Case "1", "TRUE", "VERO", "Y", "YES", "SI", "X"
            FactFlag = True

        Case ""
            FactFlag = False

        Case Else
            FactFlag = (WeeklyCsvDouble(FieldText) > 0)

    End Select

End Function

'
' One date's Accounts, one row per NDG - the first row when a file has
' two - as a 1-based array over FactAccountField.  Empty when the file is
' missing, which is also noted in MissingFiles.
'
Private Function LoadFactAccounts( _
    ByVal SnapshotDate As Date) As Variant

    Dim FilePath As String
    Dim Lines As Variant
    Dim HeaderFields As Variant
    Dim Fields As Variant
    Dim Data() As Variant
    Dim Seen As Object

    Dim ColNDG As Long
    Dim ColApproved As Long
    Dim ColDrawn As Long
    Dim ColMTM As Long
    Dim ColHCV As Long
    Dim ColMC As Long
    Dim ColSF As Long

    Dim NDG As String
    Dim RowCount As Long
    Dim OutputRow As Long
    Dim r As Long

    FilePath = FactSourcePath(SnapshotDate, ACCOUNT_FILE_SUFFIX)

    If Dir(FilePath) = "" Then
        MissingFiles = MissingFiles & vbCrLf & FilePath
        Exit Function
    End If

    Lines = ReadAllLines(FilePath)

    If UBound(Lines) < 1 Then Exit Function

    HeaderFields = Split(CStr(Lines(0)), ";")

    ColNDG = FactHeaderIndex(HeaderFields, Array("NDG"), True, FilePath)
    ColApproved = FactHeaderIndex( _
        HeaderFields, Array("Max Approved Loan", "Approved"), True, FilePath)
    ColDrawn = FactHeaderIndex( _
        HeaderFields, Array("Drawn Amount", "Drawn"), True, FilePath)
    ColMTM = FactHeaderIndex( _
        HeaderFields, _
        Array("MTM Collateral (MTM_t)", "MTM Collateral", "MTM"), _
        False, FilePath)
    ColHCV = FactHeaderIndex( _
        HeaderFields, _
        Array("Haircut Collateral Value (HCV_t)", "Haircut Collateral Value", "HCV"), _
        False, FilePath)
    ColMC = FactHeaderIndex( _
        HeaderFields, _
        Array("Margin Call (HCV_t <= Max Approved Loan)", "Margin Call", "MC"), _
        False, FilePath)
    ColSF = FactHeaderIndex( _
        HeaderFields, _
        Array("Shortfall (HTM_t <= Max Approved Loan)", "Shortfall", "SF"), _
        False, FilePath)

    For r = 1 To UBound(Lines)
        If Len(Trim$(CStr(Lines(r)))) > 0 Then RowCount = RowCount + 1
    Next r

    If RowCount = 0 Then Exit Function

    ReDim Data(1 To RowCount, 1 To FACT_ACCOUNT_FIELDS)

    Set Seen = CreateObject("Scripting.Dictionary")
    Seen.CompareMode = vbTextCompare

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            Fields = Split(CStr(Lines(r)), ";")
            NDG = FactField(Fields, ColNDG)

            If NDG <> "" Then

                If Not Seen.Exists(NDG) Then

                    Seen(NDG) = True
                    OutputRow = OutputRow + 1

                    Data(OutputRow, FactAccNDG) = NDG
                    Data(OutputRow, FactAccApproved) = _
                        WeeklyCsvDouble(FactField(Fields, ColApproved))
                    Data(OutputRow, FactAccDrawn) = _
                        WeeklyCsvDouble(FactField(Fields, ColDrawn))
                    Data(OutputRow, FactAccMTM) = _
                        WeeklyCsvDouble(FactField(Fields, ColMTM))
                    Data(OutputRow, FactAccHCV) = _
                        WeeklyCsvDouble(FactField(Fields, ColHCV))
                    Data(OutputRow, FactAccMarginCall) = _
                        FactFlag(FactField(Fields, ColMC))
                    Data(OutputRow, FactAccShortfall) = _
                        FactFlag(FactField(Fields, ColSF))

                End If

            End If

        End If

    Next r

    If OutputRow = 0 Then Exit Function

    LoadFactAccounts = TrimFactRows(Data, OutputRow, FACT_ACCOUNT_FIELDS)

End Function

'
' One date's Positions as a 1-based array over FactPositionField, every
' row kept.  Empty when the file is missing.
'
Private Function LoadFactPositions( _
    ByVal SnapshotDate As Date) As Variant

    Dim FilePath As String
    Dim Lines As Variant
    Dim HeaderFields As Variant
    Dim Fields As Variant
    Dim Data() As Variant

    Dim ColNDG As Long
    Dim ColISIN As Long
    Dim ColName As Long
    Dim ColAssetType As Long
    Dim ColCurrency As Long
    Dim ColIssuer As Long
    Dim ColValue As Long
    Dim ColHCV As Long
    Dim ColAboveLimit As Long

    Dim AssetType As String
    Dim RowCount As Long
    Dim OutputRow As Long
    Dim r As Long

    FilePath = FactSourcePath(SnapshotDate, POSITION_FILE_SUFFIX)

    If Dir(FilePath) = "" Then
        MissingFiles = MissingFiles & vbCrLf & FilePath
        Exit Function
    End If

    Lines = ReadAllLines(FilePath)

    If UBound(Lines) < 1 Then Exit Function

    HeaderFields = Split(CStr(Lines(0)), ";")

    ColNDG = FactHeaderIndex(HeaderFields, Array("NDG"), True, FilePath)
    ColISIN = FactHeaderIndex(HeaderFields, Array("ISIN"), True, FilePath)
    ColName = FactHeaderIndex( _
        HeaderFields, Array("Security Name"), True, FilePath)
    ColAssetType = FactHeaderIndex( _
        HeaderFields, _
        Array("Asset Type / Classification", "Asset Type", "Classification"), _
        True, FilePath)
    ColValue = FactHeaderIndex( _
        HeaderFields, Array("Position Value"), True, FilePath)
    ColCurrency = FactHeaderIndex( _
        HeaderFields, Array("Pricing Currency", "Currency", "CCY"), False, FilePath)
    ColIssuer = FactHeaderIndex(HeaderFields, Array("Issuer"), False, FilePath)
    ColHCV = FactHeaderIndex( _
        HeaderFields, Array("Max LTV Value", "HCV"), False, FilePath)
    ColAboveLimit = FactHeaderIndex( _
        HeaderFields, _
        Array("Position MV above limit (Haircut to Zero)", "Position MV above limit"), _
        False, FilePath)

    For r = 1 To UBound(Lines)
        If Len(Trim$(CStr(Lines(r)))) > 0 Then RowCount = RowCount + 1
    Next r

    If RowCount = 0 Then Exit Function

    ReDim Data(1 To RowCount, 1 To FACT_POSITION_FIELDS)

    For r = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(r)))) > 0 Then

            Fields = Split(CStr(Lines(r)), ";")
            OutputRow = OutputRow + 1

            AssetType = FactField(Fields, ColAssetType)

            Data(OutputRow, FactPosNDG) = FactField(Fields, ColNDG)
            Data(OutputRow, FactPosISIN) = FactField(Fields, ColISIN)
            Data(OutputRow, FactPosName) = FactField(Fields, ColName)
            Data(OutputRow, FactPosAssetType) = AssetType
            Data(OutputRow, FactPosClass) = GetAssetClass(AssetType)
            Data(OutputRow, FactPosCurrency) = FactField(Fields, ColCurrency)
            Data(OutputRow, FactPosIssuer) = FactField(Fields, ColIssuer)
            Data(OutputRow, FactPosValue) = _
                WeeklyCsvDouble(FactField(Fields, ColValue))
            Data(OutputRow, FactPosHCV) = _
                WeeklyCsvDouble(FactField(Fields, ColHCV))
            Data(OutputRow, FactPosAboveLimit) = _
                WeeklyCsvDouble(FactField(Fields, ColAboveLimit))

        End If

    Next r

    LoadFactPositions = Data

End Function

'
' The first rows of a 1-based array, when fewer were filled than made.
'
Private Function TrimFactRows( _
    ByRef Data As Variant, _
    ByVal RowsFilled As Long, _
    ByVal FieldCount As Long) As Variant

    Dim Trimmed() As Variant
    Dim r As Long
    Dim c As Long

    If RowsFilled >= UBound(Data, 1) Then
        TrimFactRows = Data
        Exit Function
    End If

    ReDim Trimmed(1 To RowsFilled, 1 To FieldCount)

    For r = 1 To RowsFilled
        For c = 1 To FieldCount
            Trimmed(r, c) = Data(r, c)
        Next c
    Next r

    TrimFactRows = Trimmed

End Function

Private Function FactDataHasRows( _
    ByRef Data As Variant) As Boolean

    If Not IsArray(Data) Then Exit Function

    On Error GoTo NoRows

    FactDataHasRows = (UBound(Data, 1) >= LBound(Data, 1))

    Exit Function

NoRows:

    Err.Clear

End Function

'====================================================================
' Sums over a snapshot
'====================================================================

Private Function NewTextDictionary() As Object

    Dim Dict As Object

    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare

    Set NewTextDictionary = Dict

End Function

Private Sub AddAmount( _
    ByVal Dict As Object, _
    ByVal Key As String, _
    ByVal Amount As Double)

    If Dict.Exists(Key) Then
        Dict(Key) = Dict(Key) + Amount
    Else
        Dict(Key) = Amount
    End If

End Sub

'
' The dictionary of dictionaries: Outer(Key) is itself a dictionary, made
' the first time the key is seen.
'
Private Function InnerDictionary( _
    ByVal Outer As Object, _
    ByVal Key As String) As Object

    If Not Outer.Exists(Key) Then Outer.Add Key, NewTextDictionary()

    Set InnerDictionary = Outer(Key)

End Function

'
' Outer(Key)(Member) = True, the inner dictionary made when needed.
'
Private Sub MarkInner( _
    ByVal Outer As Object, _
    ByVal Key As String, _
    ByVal Member As String)

    Dim Inner As Object

    Set Inner = InnerDictionary(Outer, Key)
    Inner(Member) = True

End Sub

'
' The key holding the largest or the smallest amount; with PositiveOnly,
' among amounts above zero.  Empty when there is none.  Ties go to the
' key seen first.
'
Private Function BestKey( _
    ByVal Dict As Object, _
    ByVal Largest As Boolean, _
    Optional ByVal PositiveOnly As Boolean = False) As String

    Dim Key As Variant
    Dim Best As Double
    Dim Found As Boolean

    For Each Key In Dict.Keys

        If Not PositiveOnly Or Dict(Key) > 0 Then

            If Not Found Then
                Found = True
                Best = Dict(Key)
                BestKey = CStr(Key)
            ElseIf Largest And Dict(Key) > Best Then
                Best = Dict(Key)
                BestKey = CStr(Key)
            ElseIf Not Largest And Dict(Key) < Best Then
                Best = Dict(Key)
                BestKey = CStr(Key)
            End If

        End If

    Next Key

End Function

Private Function SecurityKey( _
    ByVal ISIN As String, _
    ByVal Name As String) As String

    If ISIN <> "" Then
        SecurityKey = ISIN
    Else
        SecurityKey = Name
    End If

End Function

'
' Everything the sections read off a positions snapshot, summed once:
' totals, per client, per client and security, per security, issuer,
' currency and class.  Cash has no ISIN, so a security is keyed by its
' ISIN or, failing that, its name.
'
Private Function BuildPositionAggregates( _
    ByRef Positions As Variant) As Object

    Dim Agg As Object

    Dim Collateral As Object
    Dim PosCount As Object
    Dim NdgIsin As Object
    Dim NdgIsinClass As Object
    Dim NdgIsinName As Object
    Dim NdgClass As Object
    Dim NdgCurrency As Object
    Dim NdgSecurities As Object
    Dim NdgHcv As Object
    Dim NdgAboveLimit As Object
    Dim Security As Object
    Dim SecurityHolders As Object
    Dim SecurityName As Object
    Dim SecurityClass As Object
    Dim Issuers As Object
    Dim IssuerHolders As Object
    Dim Currencies As Object
    Dim Classes As Object

    Dim NDG As String
    Dim ISIN As String
    Dim Name As String
    Dim Class As String
    Dim Currency As String
    Dim Issuer As String
    Dim Value As Double
    Dim SecKey As String
    Dim PosKey As String

    Dim r As Long

    Set Agg = NewTextDictionary()

    Agg("Total") = 0#
    Agg("PositionCount") = 0
    Agg("HcvTotal") = 0#
    Agg("AboveLimitTotal") = 0#

    Set Collateral = NewTextDictionary()
    Set PosCount = NewTextDictionary()
    Set NdgIsin = NewTextDictionary()
    Set NdgIsinClass = NewTextDictionary()
    Set NdgIsinName = NewTextDictionary()
    Set NdgClass = NewTextDictionary()
    Set NdgCurrency = NewTextDictionary()
    Set NdgSecurities = NewTextDictionary()
    Set NdgHcv = NewTextDictionary()
    Set NdgAboveLimit = NewTextDictionary()
    Set Security = NewTextDictionary()
    Set SecurityHolders = NewTextDictionary()
    Set SecurityName = NewTextDictionary()
    Set SecurityClass = NewTextDictionary()
    Set Issuers = NewTextDictionary()
    Set IssuerHolders = NewTextDictionary()
    Set Currencies = NewTextDictionary()
    Set Classes = NewTextDictionary()

    Agg.Add "Collateral", Collateral
    Agg.Add "PosCount", PosCount
    Agg.Add "NdgIsin", NdgIsin
    Agg.Add "NdgIsinClass", NdgIsinClass
    Agg.Add "NdgIsinName", NdgIsinName
    Agg.Add "NdgClass", NdgClass
    Agg.Add "NdgCurrency", NdgCurrency
    Agg.Add "NdgSecurities", NdgSecurities
    Agg.Add "NdgHcv", NdgHcv
    Agg.Add "NdgAboveLimit", NdgAboveLimit
    Agg.Add "Security", Security
    Agg.Add "SecurityHolders", SecurityHolders
    Agg.Add "SecurityName", SecurityName
    Agg.Add "SecurityClass", SecurityClass
    Agg.Add "Issuer", Issuers
    Agg.Add "IssuerHolders", IssuerHolders
    Agg.Add "Currency", Currencies
    Agg.Add "Class", Classes

    Set BuildPositionAggregates = Agg

    If Not FactDataHasRows(Positions) Then Exit Function

    For r = LBound(Positions, 1) To UBound(Positions, 1)

        NDG = CStr(Positions(r, FactPosNDG))

        If NDG <> "" Then

            ISIN = CStr(Positions(r, FactPosISIN))
            Name = CStr(Positions(r, FactPosName))
            Class = CStr(Positions(r, FactPosClass))
            Currency = UCase$(CStr(Positions(r, FactPosCurrency)))
            Issuer = CStr(Positions(r, FactPosIssuer))
            Value = CDbl(Positions(r, FactPosValue))

            SecKey = SecurityKey(ISIN, Name)
            PosKey = NDG & vbTab & SecKey

            Agg("Total") = Agg("Total") + Value
            Agg("PositionCount") = Agg("PositionCount") + 1
            Agg("HcvTotal") = Agg("HcvTotal") + CDbl(Positions(r, FactPosHCV))
            Agg("AboveLimitTotal") = _
                Agg("AboveLimitTotal") + CDbl(Positions(r, FactPosAboveLimit))

            AddAmount Collateral, NDG, Value
            AddAmount PosCount, NDG, 1
            AddAmount NdgIsin, PosKey, Value
            AddAmount NdgHcv, NDG, CDbl(Positions(r, FactPosHCV))
            AddAmount NdgAboveLimit, NDG, CDbl(Positions(r, FactPosAboveLimit))

            If Not NdgIsinClass.Exists(PosKey) Then
                NdgIsinClass(PosKey) = Class
                NdgIsinName(PosKey) = Name
            End If

            AddAmount InnerDictionary(NdgClass, NDG), Class, Value
            AddAmount InnerDictionary(NdgCurrency, NDG), Currency, Value
            MarkInner NdgSecurities, NDG, SecKey

            AddAmount Security, SecKey, Value
            MarkInner SecurityHolders, SecKey, NDG

            If Not SecurityName.Exists(SecKey) Then
                SecurityName(SecKey) = Name
                SecurityClass(SecKey) = Class
            End If

            If Issuer <> "" Then
                AddAmount Issuers, Issuer, Value
                MarkInner IssuerHolders, Issuer, NDG
            End If

            AddAmount Currencies, Currency, Value
            AddAmount Classes, Class, Value

        End If

    Next r

End Function

'
' The NDGs of an Accounts snapshot, each to its row.
'
Private Function AccountIndex( _
    ByRef Accounts As Variant) As Object

    Dim Index As Object
    Dim r As Long

    Set Index = NewTextDictionary()
    Set AccountIndex = Index

    If Not FactDataHasRows(Accounts) Then Exit Function

    For r = LBound(Accounts, 1) To UBound(Accounts, 1)
        If Not Index.Exists(CStr(Accounts(r, FactAccNDG))) Then
            Index.Add CStr(Accounts(r, FactAccNDG)), r
        End If
    Next r

End Function

'
' One Accounts column as a dictionary by NDG.
'
Private Function AccountColumn( _
    ByRef Accounts As Variant, _
    ByVal Field As FactAccountField) As Object

    Dim Dict As Object
    Dim r As Long

    Set Dict = NewTextDictionary()
    Set AccountColumn = Dict

    If Not FactDataHasRows(Accounts) Then Exit Function

    For r = LBound(Accounts, 1) To UBound(Accounts, 1)
        If Not Dict.Exists(CStr(Accounts(r, FactAccNDG))) Then
            Dict.Add CStr(Accounts(r, FactAccNDG)), CDbl(Accounts(r, Field))
        End If
    Next r

End Function

Private Function AccountSum( _
    ByRef Accounts As Variant, _
    ByVal Field As FactAccountField) As Double

    Dim r As Long

    If Not FactDataHasRows(Accounts) Then Exit Function

    For r = LBound(Accounts, 1) To UBound(Accounts, 1)
        AccountSum = AccountSum + CDbl(Accounts(r, Field))
    Next r

End Function

Private Function AccountFlagCount( _
    ByRef Accounts As Variant, _
    ByVal Field As FactAccountField) As Long

    Dim r As Long

    If Not FactDataHasRows(Accounts) Then Exit Function

    For r = LBound(Accounts, 1) To UBound(Accounts, 1)
        If CBool(Accounts(r, Field)) Then AccountFlagCount = AccountFlagCount + 1
    Next r

End Function

Private Function AccountRowCount( _
    ByRef Accounts As Variant) As Long

    If FactDataHasRows(Accounts) Then
        AccountRowCount = UBound(Accounts, 1) - LBound(Accounts, 1) + 1
    End If

End Function

Private Function DictAmount( _
    ByVal Dict As Object, _
    ByVal Key As String) As Double

    If Dict.Exists(Key) Then DictAmount = CDbl(Dict(Key))

End Function

Private Function SafeShare( _
    ByVal Part As Double, _
    ByVal Whole As Double) As Double

    If Abs(Whole) >= FACT_TOLERANCE Then SafeShare = Part / Whole

End Function

'
' The values of a dictionary, sorted ascending, as a 0-based array of
' doubles; Empty for an empty dictionary.
'
Private Function SortedAmounts( _
    ByVal Dict As Object) As Variant

    Dim Values() As Double
    Dim Key As Variant
    Dim n As Long
    Dim Gap As Long
    Dim i As Long
    Dim j As Long
    Dim Temp As Double

    If Dict.Count = 0 Then Exit Function

    ReDim Values(0 To Dict.Count - 1)

    For Each Key In Dict.Keys
        Values(n) = CDbl(Dict(Key))
        n = n + 1
    Next Key

    Gap = n \ 2

    Do While Gap > 0

        For i = Gap To n - 1

            Temp = Values(i)
            j = i

            Do While j >= Gap
                If Values(j - Gap) <= Temp Then Exit Do
                Values(j) = Values(j - Gap)
                j = j - Gap
            Loop

            Values(j) = Temp

        Next i

        Gap = Gap \ 2

    Loop

    SortedAmounts = Values

End Function

Private Function MedianOf( _
    ByRef Sorted As Variant) As Double

    Dim n As Long

    If IsEmpty(Sorted) Then Exit Function

    n = UBound(Sorted) - LBound(Sorted) + 1

    If n Mod 2 = 1 Then
        MedianOf = Sorted(LBound(Sorted) + n \ 2)
    Else
        MedianOf = _
            (Sorted(LBound(Sorted) + n \ 2 - 1) + Sorted(LBound(Sorted) + n \ 2)) / 2
    End If

End Function

'====================================================================
' Text
'====================================================================

Private Function EuroText( _
    ByVal Amount As Double) As String

    EuroText = ChrW(&H20AC) & Format(Amount, "#,##0")

End Function

Private Function SignedEuroText( _
    ByVal Amount As Double) As String

    If Amount < 0 Then
        SignedEuroText = ChrW(&H2212) & EuroText(-Amount)
    Else
        SignedEuroText = "+" & EuroText(Amount)
    End If

End Function

Private Function PctText( _
    ByVal Share As Double) As String

    PctText = Format(Share, "0.0%")

End Function

Private Function SignedPctText( _
    ByVal Share As Double) As String

    If Share < 0 Then
        SignedPctText = ChrW(&H2212) & Format(-Share, "0.0%")
    Else
        SignedPctText = "+" & Format(Share, "0.0%")
    End If

End Function

Private Function NdgText( _
    ByVal NDG As String) As String

    NdgText = "NDG " & NDG

End Function

'
' A security as "name (ISIN)"; the name alone when it has no ISIN, the
' key alone when the name is not on file.
'
Private Function SecurityText( _
    ByVal Agg As Object, _
    ByVal SecKey As String) As String

    Dim Name As String

    Name = ""
    If Agg("SecurityName").Exists(SecKey) Then Name = CStr(Agg("SecurityName")(SecKey))

    If Name = "" Then
        SecurityText = SecKey
    ElseIf Name = SecKey Then
        SecurityText = Name
    Else
        SecurityText = Name & " (" & SecKey & ")"
    End If

End Function

Private Function PositionText( _
    ByVal Agg As Object, _
    ByVal PosKey As String) As String

    Dim Parts() As String

    Parts = Split(PosKey, vbTab)

    PositionText = NdgText(Parts(0)) & " - " & SecurityText(Agg, Parts(1))

End Function

Private Function PosKeyNdg( _
    ByVal PosKey As String) As String

    PosKeyNdg = Split(PosKey, vbTab)(0)

End Function

Private Function PosKeySecurity( _
    ByVal PosKey As String) As String

    PosKeySecurity = Split(PosKey, vbTab)(1)

End Function

Private Function CategoryLabel( _
    ByVal Class As String) As String

    Dim Category As Variant

    For Each Category In CollateralCategories()
        If StrComp(Category(0), Class, vbTextCompare) = 0 Then
            CategoryLabel = Category(1)
            Exit Function
        End If
    Next Category

    CategoryLabel = Class

End Function

Private Function DateText( _
    ByVal Value As Date) As String

    DateText = Format(Value, FACT_DATE_FORMAT)

End Function

Private Function Plural( _
    ByVal Count As Long, _
    ByVal Singular As String, _
    ByVal PluralForm As String) As String

    Plural = Format(Count, "#,##0") & " " & IIf(Count = 1, Singular, PluralForm)

End Function

'====================================================================
' Writing the sheet
'====================================================================

Private Sub WriteFactsHeader( _
    ByVal EndDate As Date)

    With FactsSheet

        .Cells(2, FIRST_COL).Value = "Weekly Facts"
        .Cells(2, FIRST_COL).Font.Size = 16
        .Cells(2, FIRST_COL).Font.Bold = True

        .Cells(3, FIRST_COL).Value = "As of " & DateText(EndDate)
        .Cells(3, FIRST_COL).Font.Size = 11

        .Cells(5, FIRST_COL).Value = "Fact"
        .Cells(5, FIRST_COL + 1).Value = "Value"
        .Cells(5, FIRST_COL + 2).Value = "Who / What"
        .Cells(5, FIRST_COL + 3).Value = "Detail"

        With .Range(.Cells(5, FIRST_COL), .Cells(5, FIRST_COL + 3))
            .Font.Bold = True
            .Interior.Color = RGB(212, 212, 212)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With

    End With

    FactsRow = 6

End Sub

'
' A section title on its own row, with a line under it that says what
' the section is read against.
'
Private Sub StartSection( _
    ByVal Title As String, _
    Optional ByVal Basis As String = "")

    FactsRow = FactsRow + 1

    With FactsSheet.Cells(FactsRow, FIRST_COL)
        .Value = Title
        .Font.Bold = True
        .Font.Size = 12
    End With

    If Basis <> "" Then
        With FactsSheet.Cells(FactsRow, FIRST_COL + 2)
            .Value = Basis
            .Font.Italic = True
            .Font.Color = RGB(90, 90, 90)
        End With
    End If

    With FactsSheet.Range( _
        FactsSheet.Cells(FactsRow, FIRST_COL), _
        FactsSheet.Cells(FactsRow, FIRST_COL + 3))
        .Interior.Color = RGB(235, 235, 235)
    End With

    FactsRow = FactsRow + 1

End Sub

'
' One fact on one row: what it is, its value in the format its kind asks
' for - eur, pct, int, x (a ratio), date, or text - who or what it is
' about, and anything worth adding.
'
Private Sub WriteFact( _
    ByVal Label As String, _
    ByVal Value As Variant, _
    Optional ByVal Kind As String = "", _
    Optional ByVal Who As String = "", _
    Optional ByVal Detail As String = "")

    With FactsSheet

        .Cells(FactsRow, FIRST_COL).Value = Label
        .Cells(FactsRow, FIRST_COL + 1).Value = Value
        .Cells(FactsRow, FIRST_COL + 2).Value = Who
        .Cells(FactsRow, FIRST_COL + 3).Value = Detail

        With .Cells(FactsRow, FIRST_COL + 1)

            Select Case Kind
                Case "eur"
                    .NumberFormat = ChrW(&H20AC) & "#,##0;" & ChrW(&H2212) & ChrW(&H20AC) & "#,##0"
                Case "pct"
                    .NumberFormat = "0.0%;" & ChrW(&H2212) & "0.0%"
                Case "int"
                    .NumberFormat = "#,##0;" & ChrW(&H2212) & "#,##0"
                Case "x"
                    .NumberFormat = "0.00""x"""
                Case "ratio"
                    .NumberFormat = "0.000"
                Case "date"
                    .NumberFormat = FACT_DATE_FORMAT
                Case Else
                    .NumberFormat = "@"
                    .HorizontalAlignment = xlLeft
            End Select

        End With

    End With

    FactsRow = FactsRow + 1

End Sub

Private Sub WriteNoFact( _
    ByVal Message As String)

    With FactsSheet.Cells(FactsRow, FIRST_COL)
        .Value = Message
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With

    FactsRow = FactsRow + 1

End Sub

Private Sub FinishFactsSheet()

    Dim Item As Variant

    If FactsNotes.Count > 0 Then

        StartSection "Notes"

        For Each Item In FactsNotes
            WriteNoFact ChrW(8226) & " " & CStr(Item)
        Next Item

    End If

    With FactsSheet

        .Cells.Font.name = "Aptos Display"
        .Cells.Font.Size = 10
        .Cells(2, FIRST_COL).Font.Size = 16

        .Columns(FIRST_COL).ColumnWidth = 44
        .Columns(FIRST_COL + 1).ColumnWidth = 18
        .Columns(FIRST_COL + 2).ColumnWidth = 48
        .Columns(FIRST_COL + 3).ColumnWidth = 80
        .Columns(1).ColumnWidth = 2

        .Range(.Cells(6, FIRST_COL + 2), .Cells(FactsRow, FIRST_COL + 3)).WrapText = False
        .Range(.Cells(6, FIRST_COL + 1), .Cells(FactsRow, FIRST_COL + 1)).HorizontalAlignment = xlRight

        .Activate
        .Range("A1").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitRow = 0
        ActiveWindow.SplitColumn = 0
        .Range("A6").Select
        ActiveWindow.FreezePanes = True

    End With

End Sub

'====================================================================
' The book as of the end date
'====================================================================

Private Sub WritePortfolioSection( _
    ByRef Snap As FactSnapshot)

    Dim Agg As Object
    Dim Sorted As Variant
    Dim Category As Variant
    Dim Key As Variant

    Dim Loans As Long
    Dim Approved As Double
    Dim Drawn As Double
    Dim Mtm As Double
    Dim Hcv As Double
    Dim Collateral As Double
    Dim MarginCalls As Long
    Dim Shortfalls As Long

    Dim Securities As Long
    Dim EuroShare As Double
    Dim Share As Double
    Dim Hhi As Double
    Dim TopFive As Double
    Dim TopTen As Double
    Dim Untouched As Long
    Dim UnusedApproved As Double
    Dim HighUse As Long
    Dim SingleSecurity As Long
    Dim FullHouse As Long
    Dim CategoryCount As Long
    Dim Held As Long

    Dim i As Long
    Dim r As Long

    Set Agg = Snap.Aggregates

    StartSection "The book", "as of " & DateText(Snap.AsOfDate)

    Loans = AccountRowCount(Snap.Accounts)
    Approved = AccountSum(Snap.Accounts, FactAccApproved)
    Drawn = AccountSum(Snap.Accounts, FactAccDrawn)
    Mtm = AccountSum(Snap.Accounts, FactAccMTM)
    Hcv = AccountSum(Snap.Accounts, FactAccHCV)
    Collateral = CDbl(Agg("Total"))
    MarginCalls = AccountFlagCount(Snap.Accounts, FactAccMarginCall)
    Shortfalls = AccountFlagCount(Snap.Accounts, FactAccShortfall)

    WriteFact "Active loans", Loans, "int"
    WriteFact "Approved lines", Approved, "eur"
    WriteFact "Drawn", Drawn, "eur"
    WriteFact "Utilisation", SafeShare(Drawn, Approved), "pct", "", _
        "drawn over approved"
    WriteFact "Collateral", Collateral, "eur", "", _
        Plural(CLng(Agg("PositionCount")), "position", "positions") & _
        " held by " & Plural(Agg("Collateral").Count, "client", "clients")

    If Mtm > 0 Then
        WriteFact "Collateral as Accounts carry it (MTM)", Mtm, "eur", "", _
            "against the positions' " & EuroText(Collateral)
    End If

    If Hcv > 0 Then
        WriteFact "Haircut collateral value", Hcv, "eur", "", _
            "an advance rate of " & PctText(SafeShare(Hcv, Collateral)) & _
            " on the collateral; headroom over drawn " & SignedEuroText(Hcv - Drawn)
    End If

    WriteFact "Loan to value", SafeShare(Drawn, Collateral), "pct", "", _
        "drawn over collateral"
    WriteFact "Cover", SafeShare(Collateral, Drawn), "x", "", _
        "collateral over drawn"

    For Each Key In Agg("Security").Keys
        If StrComp(CStr(Agg("SecurityClass")(Key)), "Cash", vbTextCompare) <> 0 Then
            Securities = Securities + 1
        End If
    Next Key

    WriteFact "Securities held", Securities, "int", "", "cash aside"
    WriteFact "Issuers", Agg("Issuer").Count, "int"

    EuroShare = SafeShare(DictAmount(Agg("Currency"), "EUR"), Collateral)
    WriteFact "Currencies", Agg("Currency").Count, "int", "", _
        PctText(EuroShare) & " of the collateral is priced in euro"

    WriteFact "Clients in margin call", MarginCalls, "int", "", _
        PctText(SafeShare(MarginCalls, Loans)) & " of loans"
    WriteFact "Clients in shortfall", Shortfalls, "int", "", _
        PctText(SafeShare(Shortfalls, Loans)) & " of loans"

    Sorted = SortedAmounts(Agg("Collateral"))

    If Not IsEmpty(Sorted) Then

        WriteFact "Median client", MedianOf(Sorted), "eur", "", _
            "the average client holds " & _
            EuroText(SafeShare(Collateral, Agg("Collateral").Count))

        For i = UBound(Sorted) To LBound(Sorted) Step -1

            Share = SafeShare(Sorted(i), Collateral)
            Hhi = Hhi + Share * Share

            If UBound(Sorted) - i < 5 Then TopFive = TopFive + Share
            If UBound(Sorted) - i < 10 Then TopTen = TopTen + Share

        Next i

        WriteFact "Top 5 clients' share", TopFive, "pct", "", "of the collateral"
        WriteFact "Top 10 clients' share", TopTen, "pct", "", "of the collateral"
        WriteFact "Concentration (Herfindahl index)", Hhi, "ratio", "", _
            "as concentrated as " & Format(SafeShare(1, Hhi), "0.0") & _
            " equal clients would be"

    End If

    If FactDataHasRows(Snap.Accounts) Then

        For r = LBound(Snap.Accounts, 1) To UBound(Snap.Accounts, 1)

            If CDbl(Snap.Accounts(r, FactAccApproved)) > 0 Then

                If Abs(CDbl(Snap.Accounts(r, FactAccDrawn))) < FACT_TOLERANCE Then
                    Untouched = Untouched + 1
                    UnusedApproved = UnusedApproved + CDbl(Snap.Accounts(r, FactAccApproved))
                ElseIf CDbl(Snap.Accounts(r, FactAccDrawn)) >= _
                       0.9 * CDbl(Snap.Accounts(r, FactAccApproved)) Then
                    HighUse = HighUse + 1
                End If

            End If

        Next r

    End If

    WriteFact "Untouched lines", Untouched, "int", "", _
        EuroText(UnusedApproved) & " approved and not drawn at all"
    WriteFact "Lines drawn above 90%", HighUse, "int"

    CategoryCount = UBound(CollateralCategories()) + 1

    For Each Key In Agg("NdgSecurities").Keys

        If Agg("NdgSecurities")(Key).Count = 1 Then SingleSecurity = SingleSecurity + 1

        Held = 0

        For Each Category In CollateralCategories()
            If Agg("NdgClass")(Key).Exists(Category(0)) Then Held = Held + 1
        Next Category

        If Held = CategoryCount Then FullHouse = FullHouse + 1

    Next Key

    WriteFact "Clients with a single security", SingleSecurity, "int"
    WriteFact "Clients holding every category", FullHouse, "int", "", _
        "all " & CategoryCount & " collateral categories at once"

End Sub

'====================================================================
' Clients
'====================================================================

Private Sub WriteClientSection( _
    ByRef Snap As FactSnapshot)

    Dim Agg As Object
    Dim Collateral As Object
    Dim DrawnBy As Object
    Dim ApprovedBy As Object
    Dim HcvBy As Object
    Dim Utilisation As Object
    Dim UtilisationDrawn As Object
    Dim Cover As Object
    Dim Headroom As Object
    Dim Depth As Object
    Dim SecurityCount As Object
    Dim CurrencyCount As Object
    Dim CategoryCount As Object
    Dim TopShare As Object
    Dim TopKey As Object
    Dim TopValue As Object
    Dim CashBy As Object
    Dim NonEligibleBy As Object

    Dim Key As Variant
    Dim Category As Variant
    Dim NDG As String
    Dim Total As Double
    Dim Held As Long
    Dim r As Long

    Set Agg = Snap.Aggregates
    Set Collateral = Agg("Collateral")
    Total = CDbl(Agg("Total"))

    Set DrawnBy = AccountColumn(Snap.Accounts, FactAccDrawn)
    Set ApprovedBy = AccountColumn(Snap.Accounts, FactAccApproved)
    Set HcvBy = AccountColumn(Snap.Accounts, FactAccHCV)

    StartSection "Clients", "as of " & DateText(Snap.AsOfDate)

    If Collateral.Count = 0 And DrawnBy.Count = 0 Then
        WriteNoFact "No clients on file."
        Exit Sub
    End If

    '
    ' Size
    '

    NDG = BestKey(Collateral, True)
    If NDG <> "" Then
        WriteFact "Largest client by collateral", Collateral(NDG), "eur", NdgText(NDG), _
            PctText(SafeShare(Collateral(NDG), Total)) & " of the book; drawn " & _
            EuroText(DictAmount(DrawnBy, NDG)) & "; " & _
            Plural(CLng(DictAmount(Agg("PosCount"), NDG)), "position", "positions")
    End If

    NDG = BestKey(Collateral, False, True)
    If NDG <> "" Then
        WriteFact "Smallest client by collateral", Collateral(NDG), "eur", NdgText(NDG), _
            "drawn " & EuroText(DictAmount(DrawnBy, NDG)) & "; approved " & _
            EuroText(DictAmount(ApprovedBy, NDG))
    End If

    NDG = BestKey(DrawnBy, True)
    If NDG <> "" Then
        WriteFact "Largest client by drawn amount", DrawnBy(NDG), "eur", NdgText(NDG), _
            PctText(SafeShare(DrawnBy(NDG), DictAmount(ApprovedBy, NDG))) & _
            " of a " & EuroText(DictAmount(ApprovedBy, NDG)) & " line; collateral " & _
            EuroText(DictAmount(Collateral, NDG))
    End If

    NDG = BestKey(DrawnBy, False, True)
    If NDG <> "" Then
        WriteFact "Smallest drawn amount", DrawnBy(NDG), "eur", NdgText(NDG), _
            "on a " & EuroText(DictAmount(ApprovedBy, NDG)) & " line"
    End If

    NDG = BestKey(ApprovedBy, True)
    If NDG <> "" Then
        WriteFact "Largest approved line", ApprovedBy(NDG), "eur", NdgText(NDG), _
            "drawn " & EuroText(DictAmount(DrawnBy, NDG)) & " (" & _
            PctText(SafeShare(DictAmount(DrawnBy, NDG), ApprovedBy(NDG))) & ")"
    End If

    NDG = BestKey(ApprovedBy, False, True)
    If NDG <> "" Then
        WriteFact "Smallest approved line", ApprovedBy(NDG), "eur", NdgText(NDG), _
            "drawn " & EuroText(DictAmount(DrawnBy, NDG))
    End If

    '
    ' Use of the line, cover and headroom
    '

    Set Utilisation = NewTextDictionary()
    Set UtilisationDrawn = NewTextDictionary()
    Set Cover = NewTextDictionary()
    Set Headroom = NewTextDictionary()
    Set Depth = NewTextDictionary()

    If FactDataHasRows(Snap.Accounts) Then

        For r = LBound(Snap.Accounts, 1) To UBound(Snap.Accounts, 1)

            NDG = CStr(Snap.Accounts(r, FactAccNDG))

            If CDbl(Snap.Accounts(r, FactAccApproved)) > 0 Then
                Utilisation(NDG) = _
                    CDbl(Snap.Accounts(r, FactAccDrawn)) / _
                    CDbl(Snap.Accounts(r, FactAccApproved))
                If CDbl(Snap.Accounts(r, FactAccDrawn)) >= FACT_TOLERANCE Then
                    UtilisationDrawn(NDG) = Utilisation(NDG)
                End If
            End If

            If CDbl(Snap.Accounts(r, FactAccDrawn)) >= FACT_TOLERANCE Then

                Cover(NDG) = _
                    DictAmount(Collateral, NDG) / CDbl(Snap.Accounts(r, FactAccDrawn))

                If CDbl(Snap.Accounts(r, FactAccHCV)) > 0 Then
                    If CBool(Snap.Accounts(r, FactAccMarginCall)) Then
                        Depth(NDG) = _
                            CDbl(Snap.Accounts(r, FactAccDrawn)) - _
                            CDbl(Snap.Accounts(r, FactAccHCV))
                    Else
                        Headroom(NDG) = _
                            CDbl(Snap.Accounts(r, FactAccHCV)) - _
                            CDbl(Snap.Accounts(r, FactAccDrawn))
                    End If
                End If

            End If

        Next r

    End If

    NDG = BestKey(Utilisation, True)
    If NDG <> "" Then
        WriteFact "Highest utilisation", Utilisation(NDG), "pct", NdgText(NDG), _
            EuroText(DictAmount(DrawnBy, NDG)) & " drawn of " & _
            EuroText(DictAmount(ApprovedBy, NDG))
    End If

    NDG = BestKey(UtilisationDrawn, False)
    If NDG <> "" Then
        WriteFact "Lowest utilisation among drawn lines", UtilisationDrawn(NDG), "pct", _
            NdgText(NDG), EuroText(DictAmount(DrawnBy, NDG)) & " drawn of " & _
            EuroText(DictAmount(ApprovedBy, NDG))
    End If

    NDG = BestKey(Cover, False)
    If NDG <> "" Then
        WriteFact "Thinnest cover", Cover(NDG), "x", NdgText(NDG), _
            EuroText(DictAmount(Collateral, NDG)) & " of collateral over " & _
            EuroText(DictAmount(DrawnBy, NDG)) & " drawn; loan to value " & _
            PctText(SafeShare(DictAmount(DrawnBy, NDG), DictAmount(Collateral, NDG)))
    End If

    NDG = BestKey(Cover, True)
    If NDG <> "" Then
        WriteFact "Thickest cover", Cover(NDG), "x", NdgText(NDG), _
            EuroText(DictAmount(Collateral, NDG)) & " of collateral over " & _
            EuroText(DictAmount(DrawnBy, NDG)) & " drawn"
    End If

    NDG = BestKey(Headroom, False)
    If NDG <> "" Then
        WriteFact "Closest to a margin call", Headroom(NDG), "eur", NdgText(NDG), _
            "haircut collateral value " & EuroText(DictAmount(HcvBy, NDG)) & _
            " over " & EuroText(DictAmount(DrawnBy, NDG)) & " drawn; " & _
            PctText(SafeShare(Headroom(NDG), DictAmount(DrawnBy, NDG))) & " of headroom"
    End If

    NDG = BestKey(Depth, True)
    If NDG <> "" Then
        WriteFact "Deepest margin call", Depth(NDG), "eur", NdgText(NDG), _
            EuroText(DictAmount(DrawnBy, NDG)) & " drawn against " & _
            EuroText(DictAmount(HcvBy, NDG)) & " of haircut collateral value"
    End If

    '
    ' How the collateral is held
    '

    Set SecurityCount = NewTextDictionary()
    Set CurrencyCount = NewTextDictionary()
    Set CategoryCount = NewTextDictionary()
    Set TopShare = NewTextDictionary()
    Set TopKey = NewTextDictionary()
    Set TopValue = NewTextDictionary()
    Set CashBy = NewTextDictionary()
    Set NonEligibleBy = NewTextDictionary()

    For Each Key In Agg("NdgSecurities").Keys

        SecurityCount(Key) = Agg("NdgSecurities")(Key).Count
        CurrencyCount(Key) = Agg("NdgCurrency")(Key).Count

        Held = 0
        For Each Category In CollateralCategories()
            If Agg("NdgClass")(Key).Exists(Category(0)) Then Held = Held + 1
        Next Category
        CategoryCount(Key) = Held

        If Agg("NdgClass")(Key).Exists("Cash") Then
            CashBy(Key) = Agg("NdgClass")(Key)("Cash")
        End If

        If Agg("NdgClass")(Key).Exists("Non Eligible Asset") Then
            NonEligibleBy(Key) = Agg("NdgClass")(Key)("Non Eligible Asset")
        End If

    Next Key

    For Each Key In Agg("NdgIsin").Keys

        NDG = PosKeyNdg(CStr(Key))

        If Not TopValue.Exists(NDG) Then
            TopValue(NDG) = Agg("NdgIsin")(Key)
            TopKey(NDG) = Key
        ElseIf CDbl(Agg("NdgIsin")(Key)) > CDbl(TopValue(NDG)) Then
            TopValue(NDG) = Agg("NdgIsin")(Key)
            TopKey(NDG) = Key
        End If

    Next Key

    For Each Key In TopValue.Keys
        If DictAmount(SecurityCount, CStr(Key)) >= 2 And DictAmount(Collateral, CStr(Key)) > 0 Then
            TopShare(Key) = CDbl(TopValue(Key)) / CDbl(Collateral(Key))
        End If
    Next Key

    NDG = BestKey(Agg("PosCount"), True)
    If NDG <> "" Then
        WriteFact "Most positions", Agg("PosCount")(NDG), "int", NdgText(NDG), _
            Plural(CLng(DictAmount(SecurityCount, NDG)), "security", "securities") & _
            " worth " & EuroText(DictAmount(Collateral, NDG))
    End If

    NDG = BestKey(SecurityCount, True)
    If NDG <> "" Then
        WriteFact "Most securities", SecurityCount(NDG), "int", NdgText(NDG), _
            "collateral " & EuroText(DictAmount(Collateral, NDG))
    End If

    NDG = BestKey(CurrencyCount, True)
    If NDG <> "" Then
        If CDbl(CurrencyCount(NDG)) > 1 Then
            WriteFact "Most currencies", CurrencyCount(NDG), "int", NdgText(NDG), _
                Join(Agg("NdgCurrency")(NDG).Keys, ", ")
        End If
    End If

    NDG = BestKey(CategoryCount, True)
    If NDG <> "" Then
        WriteFact "Most categories held", CategoryCount(NDG), "int", NdgText(NDG), _
            "of " & (UBound(CollateralCategories()) + 1)
    End If

    NDG = BestKey(TopShare, True)
    If NDG <> "" Then
        WriteFact "Most concentrated client", TopShare(NDG), "pct", NdgText(NDG), _
            "of its " & EuroText(DictAmount(Collateral, NDG)) & " in " & _
            SecurityText(Agg, PosKeySecurity(CStr(TopKey(NDG)))) & _
            "; among clients with more than one security"
    End If

    NDG = BestKey(TopShare, False)
    If NDG <> "" Then
        WriteFact "Most evenly spread client", TopShare(NDG), "pct", NdgText(NDG), _
            "its largest holding, " & _
            SecurityText(Agg, PosKeySecurity(CStr(TopKey(NDG)))) & _
            ", out of " & Plural(CLng(DictAmount(SecurityCount, NDG)), "security", "securities")
    End If

    NDG = BestKey(CashBy, True, True)
    If NDG <> "" Then
        WriteFact "Largest cash holder", CashBy(NDG), "eur", NdgText(NDG), _
            PctText(SafeShare(CashBy(NDG), DictAmount(Collateral, NDG))) & _
            " of its collateral"
    End If

    NDG = BestKey(NonEligibleBy, True, True)
    If NDG <> "" Then
        WriteFact "Most non-eligible collateral", NonEligibleBy(NDG), "eur", NdgText(NDG), _
            PctText(SafeShare(NonEligibleBy(NDG), DictAmount(Collateral, NDG))) & _
            " of its collateral"
    End If

    NDG = BestKey(Agg("NdgAboveLimit"), True, True)
    If NDG <> "" Then
        WriteFact "Most collateral above a concentration limit", _
            Agg("NdgAboveLimit")(NDG), "eur", NdgText(NDG), _
            PctText(SafeShare(Agg("NdgAboveLimit")(NDG), DictAmount(Collateral, NDG))) & _
            " of its collateral counts for nothing"
    End If

End Sub

'====================================================================
' Positions and securities
'====================================================================

Private Sub WritePositionSection( _
    ByRef Snap As FactSnapshot)

    Dim Agg As Object
    Dim ClassBest As Object
    Dim ClassBestValue As Object
    Dim HolderCount As Object
    Dim NonCash As Object
    Dim IssuerHolderCount As Object
    Dim NonEuro As Object

    Dim Key As Variant
    Dim Category As Variant
    Dim PosKey As String
    Dim SecKey As String
    Dim Class As String
    Dim Total As Double
    Dim Value As Double
    Dim LongestName As String
    Dim ShortestName As String
    Dim Name As String
    Dim Singletons As Long
    Dim SingletonValue As Double

    Set Agg = Snap.Aggregates
    Total = CDbl(Agg("Total"))

    StartSection "Positions and securities", "as of " & DateText(Snap.AsOfDate)

    If Agg("NdgIsin").Count = 0 Then
        WriteNoFact "No positions on file."
        Exit Sub
    End If

    '
    ' The largest position, over the book and in each category
    '

    Set ClassBest = NewTextDictionary()
    Set ClassBestValue = NewTextDictionary()

    For Each Key In Agg("NdgIsin").Keys

        Class = CStr(Agg("NdgIsinClass")(Key))
        Value = CDbl(Agg("NdgIsin")(Key))

        If Not ClassBestValue.Exists(Class) Then
            ClassBestValue(Class) = Value
            ClassBest(Class) = Key
        ElseIf Value > CDbl(ClassBestValue(Class)) Then
            ClassBestValue(Class) = Value
            ClassBest(Class) = Key
        End If

    Next Key

    PosKey = BestKey(Agg("NdgIsin"), True)
    WriteFact "Largest position", Agg("NdgIsin")(PosKey), "eur", PositionText(Agg, PosKey), _
        PositionDetail(Agg, PosKey)

    For Each Category In CollateralCategories()

        If ClassBest.Exists(Category(0)) Then

            PosKey = CStr(ClassBest(Category(0)))

            WriteFact "Largest " & Category(1) & " position", Agg("NdgIsin")(PosKey), _
                "eur", PositionText(Agg, PosKey), PositionDetail(Agg, PosKey)

        End If

    Next Category

    '
    ' Securities across the book, cash aside
    '

    Set HolderCount = NewTextDictionary()
    Set NonCash = NewTextDictionary()

    For Each Key In Agg("Security").Keys

        If StrComp(CStr(Agg("SecurityClass")(Key)), "Cash", vbTextCompare) <> 0 Then

            NonCash(Key) = Agg("Security")(Key)
            HolderCount(Key) = Agg("SecurityHolders")(Key).Count

            If Agg("SecurityHolders")(Key).Count = 1 Then
                Singletons = Singletons + 1
                SingletonValue = SingletonValue + CDbl(Agg("Security")(Key))
            End If

        End If

    Next Key

    SecKey = BestKey(HolderCount, True)
    If SecKey <> "" Then
        WriteFact "Most widely held security", HolderCount(SecKey), "int", _
            SecurityText(Agg, SecKey), _
            "clients hold it; " & EuroText(DictAmount(NonCash, SecKey)) & " across the book"
    End If

    SecKey = BestKey(NonCash, True)
    If SecKey <> "" Then
        WriteFact "Largest security across the book", NonCash(SecKey), "eur", _
            SecurityText(Agg, SecKey), _
            PctText(SafeShare(NonCash(SecKey), Total)) & " of the collateral, held by " & _
            Plural(CLng(DictAmount(HolderCount, SecKey)), "client", "clients")
    End If

    WriteFact "Securities held by one client only", Singletons, "int", "", _
        EuroText(SingletonValue) & " of collateral rests on securities nobody else pledged"

    Set IssuerHolderCount = NewTextDictionary()

    For Each Key In Agg("Issuer").Keys
        IssuerHolderCount(Key) = Agg("IssuerHolders")(Key).Count
    Next Key

    SecKey = BestKey(Agg("Issuer"), True)
    If SecKey <> "" Then
        WriteFact "Largest issuer", Agg("Issuer")(SecKey), "eur", SecKey, _
            PctText(SafeShare(Agg("Issuer")(SecKey), Total)) & " of the collateral, held by " & _
            Plural(CLng(DictAmount(IssuerHolderCount, SecKey)), "client", "clients")
    End If

    SecKey = BestKey(IssuerHolderCount, True)
    If SecKey <> "" Then
        WriteFact "Most common issuer", IssuerHolderCount(SecKey), "int", SecKey, _
            "clients hold its paper; " & EuroText(DictAmount(Agg("Issuer"), SecKey)) & _
            " across the book"
    End If

    '
    ' Currencies and categories
    '

    Set NonEuro = NewTextDictionary()

    For Each Key In Agg("Currency").Keys
        If CStr(Key) <> "EUR" And CStr(Key) <> "" Then NonEuro(Key) = Agg("Currency")(Key)
    Next Key

    If NonEuro.Count > 0 Then

        Value = 0
        For Each Key In NonEuro.Keys
            Value = Value + CDbl(NonEuro(Key))
        Next Key

        WriteFact "Collateral not priced in euro", Value, "eur", "", _
            PctText(SafeShare(Value, Total)) & " of the book, in " & _
            Plural(NonEuro.Count, "currency", "currencies")

        SecKey = BestKey(NonEuro, True)
        WriteFact "Largest foreign currency", NonEuro(SecKey), "eur", SecKey, _
            PctText(SafeShare(NonEuro(SecKey), Total)) & " of the book"

    End If

    WriteFact "Cash", DictAmount(Agg("Class"), "Cash"), "eur", "", _
        PctText(SafeShare(DictAmount(Agg("Class"), "Cash"), Total)) & " of the collateral"
    WriteFact "Non-eligible collateral", DictAmount(Agg("Class"), "Non Eligible Asset"), "eur", "", _
        PctText(SafeShare(DictAmount(Agg("Class"), "Non Eligible Asset"), Total)) & " of the collateral"

    If DictAmount(Agg("Class"), "UNKNOWN") > 0 Then
        WriteFact "Collateral of an unmapped asset type", DictAmount(Agg("Class"), "UNKNOWN"), _
            "eur", "", "asset types GetAssetClass could not place; outside every category"
    End If

    If CDbl(Agg("AboveLimitTotal")) > 0 Then
        WriteFact "Collateral above concentration limits", CDbl(Agg("AboveLimitTotal")), "eur", "", _
            PctText(SafeShare(CDbl(Agg("AboveLimitTotal")), Total)) & _
            " of the collateral is haircut to zero"
    End If

    '
    ' Names
    '

    For Each Key In Agg("SecurityName").Keys

        Name = CStr(Agg("SecurityName")(Key))

        If Name <> "" Then
            If Len(Name) > Len(LongestName) Then LongestName = Name
            If ShortestName = "" Or Len(Name) < Len(ShortestName) Then ShortestName = Name
        End If

    Next Key

    If LongestName <> "" Then
        WriteFact "Longest security name", Len(LongestName), "int", LongestName, "characters"
        WriteFact "Shortest security name", Len(ShortestName), "int", ShortestName, "characters"
    End If

End Sub

'
' What a position is to its client and to the book.
'
Private Function PositionDetail( _
    ByVal Agg As Object, _
    ByVal PosKey As String) As String

    Dim NDG As String
    Dim Value As Double

    NDG = PosKeyNdg(PosKey)
    Value = CDbl(Agg("NdgIsin")(PosKey))

    PositionDetail = _
        PctText(SafeShare(Value, DictAmount(Agg("Collateral"), NDG))) & _
        " of the client's collateral, " & _
        PctText(SafeShare(Value, CDbl(Agg("Total")))) & " of the book; " & _
        CategoryLabel(CStr(Agg("NdgIsinClass")(PosKey)))

End Function

'====================================================================
' What moved, read against an earlier snapshot
'====================================================================

Private Sub WriteHorizon( _
    ByRef EndSnap As FactSnapshot, _
    ByVal BaseDate As Date, _
    ByVal Title As String)

    Dim BaseSnap As FactSnapshot

    If BaseDate = 0 Then
        StartSection Title
        WriteNoFact "No earlier snapshot on file to read against."
        Exit Sub
    End If

    StartSection Title, DateText(BaseDate) & " to " & DateText(EndSnap.AsOfDate)

    BaseSnap = LoadFactSnapshot(BaseDate)

    WriteMovementFacts EndSnap, BaseSnap

End Sub

Private Sub WriteMovementFacts( _
    ByRef EndSnap As FactSnapshot, _
    ByRef BaseSnap As FactSnapshot)

    Dim EndAgg As Object
    Dim BaseAgg As Object
    Dim EndIndex As Object
    Dim BaseIndex As Object
    Dim EndDrawn As Object
    Dim BaseDrawn As Object
    Dim EndApproved As Object
    Dim BaseApproved As Object
    Dim EndMc As Object
    Dim BaseMc As Object

    Dim NewCollateral As Object
    Dim EndedCollateral As Object
    Dim Delta As Object
    Dim DeltaDrawn As Object
    Dim DeltaApproved As Object
    Dim PosDelta As Object
    Dim Turnover As Object
    Dim ClassDelta As Object
    Dim NewSecurities As Object

    Dim Key As Variant
    Dim Category As Variant
    Dim NDG As String
    Dim SecKey As String
    Dim EndTotal As Double
    Dim BaseTotal As Double
    Dim Change As Double
    Dim BaseValue As Double
    Dim EndValue As Double
    Dim NewApproved As Double
    Dim EndedApproved As Double
    Dim GoneCount As Long
    Dim GoneValue As Double
    Dim NewMcCount As Long
    Dim ClearedMcCount As Long
    Dim Days As Long

    Set EndAgg = EndSnap.Aggregates
    Set BaseAgg = BaseSnap.Aggregates
    Set EndIndex = AccountIndex(EndSnap.Accounts)
    Set BaseIndex = AccountIndex(BaseSnap.Accounts)
    Set EndDrawn = AccountColumn(EndSnap.Accounts, FactAccDrawn)
    Set BaseDrawn = AccountColumn(BaseSnap.Accounts, FactAccDrawn)
    Set EndApproved = AccountColumn(EndSnap.Accounts, FactAccApproved)
    Set BaseApproved = AccountColumn(BaseSnap.Accounts, FactAccApproved)
    Set EndMc = AccountColumn(EndSnap.Accounts, FactAccMarginCall)
    Set BaseMc = AccountColumn(BaseSnap.Accounts, FactAccMarginCall)

    EndTotal = CDbl(EndAgg("Total"))
    BaseTotal = CDbl(BaseAgg("Total"))
    Days = CLng(EndSnap.AsOfDate - BaseSnap.AsOfDate)

    '
    ' The book as a whole
    '

    Change = EndTotal - BaseTotal
    WriteFact "Collateral", Change, "eur", "", _
        SignedPctText(SafeShare(Change, BaseTotal)) & " over " & _
        Plural(Days, "day", "days") & "; " & EuroText(BaseTotal) & " to " & EuroText(EndTotal)

    Change = AccountSum(EndSnap.Accounts, FactAccDrawn) - AccountSum(BaseSnap.Accounts, FactAccDrawn)
    WriteFact "Drawn", Change, "eur", "", _
        SignedPctText(SafeShare(Change, AccountSum(BaseSnap.Accounts, FactAccDrawn))) & _
        "; " & EuroText(AccountSum(BaseSnap.Accounts, FactAccDrawn)) & " to " & _
        EuroText(AccountSum(EndSnap.Accounts, FactAccDrawn))

    Change = AccountSum(EndSnap.Accounts, FactAccApproved) - AccountSum(BaseSnap.Accounts, FactAccApproved)
    WriteFact "Approved lines", Change, "eur", "", _
        SignedPctText(SafeShare(Change, AccountSum(BaseSnap.Accounts, FactAccApproved)))

    WriteFact "Active loans", AccountRowCount(EndSnap.Accounts) - AccountRowCount(BaseSnap.Accounts), _
        "int", "", AccountRowCount(BaseSnap.Accounts) & " to " & AccountRowCount(EndSnap.Accounts)
    WriteFact "Positions", CLng(EndAgg("PositionCount")) - CLng(BaseAgg("PositionCount")), _
        "int", "", EndAgg("PositionCount") & " now"

    '
    ' Loans that came and went
    '

    Set NewCollateral = NewTextDictionary()
    Set EndedCollateral = NewTextDictionary()

    For Each Key In EndIndex.Keys
        If Not BaseIndex.Exists(Key) Then
            NewCollateral(Key) = DictAmount(EndAgg("Collateral"), CStr(Key))
            NewApproved = NewApproved + DictAmount(EndApproved, CStr(Key))
        End If
    Next Key

    For Each Key In BaseIndex.Keys
        If Not EndIndex.Exists(Key) Then
            EndedCollateral(Key) = DictAmount(BaseAgg("Collateral"), CStr(Key))
            EndedApproved = EndedApproved + DictAmount(BaseApproved, CStr(Key))
        End If
    Next Key

    WriteFact "New loans", NewCollateral.Count, "int", "", _
        EuroText(SumOf(NewCollateral)) & " of collateral on " & EuroText(NewApproved) & " of new lines"

    NDG = BestKey(NewCollateral, True)
    If NDG <> "" Then
        WriteFact "Largest new loan", NewCollateral(NDG), "eur", NdgText(NDG), _
            "collateral; line " & EuroText(DictAmount(EndApproved, NDG)) & ", drawn " & _
            EuroText(DictAmount(EndDrawn, NDG))
    End If

    WriteFact "Ended loans", EndedCollateral.Count, "int", "", _
        EuroText(SumOf(EndedCollateral)) & " of collateral on " & EuroText(EndedApproved) & _
        " of lines, as they last stood"

    NDG = BestKey(EndedCollateral, True)
    If NDG <> "" Then
        WriteFact "Largest ended loan", EndedCollateral(NDG), "eur", NdgText(NDG), _
            "collateral as it last stood; line " & EuroText(DictAmount(BaseApproved, NDG)) & _
            ", drawn " & EuroText(DictAmount(BaseDrawn, NDG))
    End If

    '
    ' The clients that stayed
    '

    Set Delta = NewTextDictionary()
    Set DeltaDrawn = NewTextDictionary()
    Set DeltaApproved = NewTextDictionary()

    For Each Key In EndIndex.Keys

        If BaseIndex.Exists(Key) Then

            NDG = CStr(Key)
            Delta(NDG) = DictAmount(EndAgg("Collateral"), NDG) - DictAmount(BaseAgg("Collateral"), NDG)
            DeltaDrawn(NDG) = DictAmount(EndDrawn, NDG) - DictAmount(BaseDrawn, NDG)
            DeltaApproved(NDG) = DictAmount(EndApproved, NDG) - DictAmount(BaseApproved, NDG)

            If CBool(EndMc(NDG)) And Not CBool(BaseMc(NDG)) Then NewMcCount = NewMcCount + 1
            If CBool(BaseMc(NDG)) And Not CBool(EndMc(NDG)) Then ClearedMcCount = ClearedMcCount + 1

        End If

    Next Key

    NDG = BestKey(Delta, True)
    If NDG <> "" Then
        If Delta(NDG) > 0 Then
            WriteFact "Biggest riser", Delta(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseAgg("Collateral"), NDG), DictAmount(EndAgg("Collateral"), NDG))
        End If
    End If

    NDG = BestKey(Delta, False)
    If NDG <> "" Then
        If Delta(NDG) < 0 Then
            WriteFact "Biggest faller", Delta(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseAgg("Collateral"), NDG), DictAmount(EndAgg("Collateral"), NDG))
        End If
    End If

    '
    ' Positions of the clients that stayed
    '

    Set PosDelta = NewTextDictionary()
    Set Turnover = NewTextDictionary()

    For Each Key In EndAgg("NdgIsin").Keys
        NDG = PosKeyNdg(CStr(Key))
        If Delta.Exists(NDG) Then
            PosDelta(Key) = CDbl(EndAgg("NdgIsin")(Key)) - DictAmount(BaseAgg("NdgIsin"), CStr(Key))
        End If
    Next Key

    For Each Key In BaseAgg("NdgIsin").Keys
        NDG = PosKeyNdg(CStr(Key))
        If Delta.Exists(NDG) And Not PosDelta.Exists(Key) Then
            PosDelta(Key) = -CDbl(BaseAgg("NdgIsin")(Key))
        End If
    Next Key

    For Each Key In PosDelta.Keys
        AddAmount Turnover, PosKeyNdg(CStr(Key)), Abs(CDbl(PosDelta(Key)))
    Next Key

    Key = BestKey(PosDelta, True)
    If CStr(Key) <> "" Then
        If PosDelta(Key) > 0 Then
            WriteFact "Largest position increase", PosDelta(Key), "eur", _
                PositionTextEither(EndAgg, BaseAgg, CStr(Key)), _
                RangeDetail(DictAmount(BaseAgg("NdgIsin"), CStr(Key)), DictAmount(EndAgg("NdgIsin"), CStr(Key)))
        End If
    End If

    Key = BestKey(PosDelta, False)
    If CStr(Key) <> "" Then
        If PosDelta(Key) < 0 Then
            WriteFact "Largest position decrease", PosDelta(Key), "eur", _
                PositionTextEither(EndAgg, BaseAgg, CStr(Key)), _
                RangeDetail(DictAmount(BaseAgg("NdgIsin"), CStr(Key)), DictAmount(EndAgg("NdgIsin"), CStr(Key)))
        End If
    End If

    NDG = BestKey(Turnover, True, True)
    If NDG <> "" Then
        WriteFact "Most active repositioner", Turnover(NDG), "eur", NdgText(NDG), _
            "moved across its positions, " & _
            PctText(SafeShare(Turnover(NDG), DictAmount(BaseAgg("Collateral"), NDG))) & _
            " of the collateral it started with"
    End If

    '
    ' Lines and drawings
    '

    NDG = BestKey(DeltaApproved, True)
    If NDG <> "" Then
        If DeltaApproved(NDG) > 0 Then
            WriteFact "Biggest line increase", DeltaApproved(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseApproved, NDG), DictAmount(EndApproved, NDG))
        End If
    End If

    NDG = BestKey(DeltaApproved, False)
    If NDG <> "" Then
        If DeltaApproved(NDG) < 0 Then
            WriteFact "Biggest line cut", DeltaApproved(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseApproved, NDG), DictAmount(EndApproved, NDG))
        End If
    End If

    NDG = BestKey(DeltaDrawn, True)
    If NDG <> "" Then
        If DeltaDrawn(NDG) > 0 Then
            WriteFact "Biggest drawdown", DeltaDrawn(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseDrawn, NDG), DictAmount(EndDrawn, NDG))
        End If
    End If

    NDG = BestKey(DeltaDrawn, False)
    If NDG <> "" Then
        If DeltaDrawn(NDG) < 0 Then
            WriteFact "Biggest repayment", DeltaDrawn(NDG), "eur", NdgText(NDG), _
                RangeDetail(DictAmount(BaseDrawn, NDG), DictAmount(EndDrawn, NDG))
        End If
    End If

    '
    ' Categories
    '

    Set ClassDelta = NewTextDictionary()

    For Each Category In CollateralCategories()
        ClassDelta(Category(0)) = _
            DictAmount(EndAgg("Class"), CStr(Category(0))) - _
            DictAmount(BaseAgg("Class"), CStr(Category(0)))
    Next Category

    Key = BestKey(ClassDelta, True)
    If CStr(Key) <> "" Then
        If ClassDelta(Key) > 0 Then
            WriteFact "Category gaining most", ClassDelta(Key), "eur", CategoryLabel(CStr(Key)), _
                RangeDetail(DictAmount(BaseAgg("Class"), CStr(Key)), DictAmount(EndAgg("Class"), CStr(Key)))
        End If
    End If

    Key = BestKey(ClassDelta, False)
    If CStr(Key) <> "" Then
        If ClassDelta(Key) < 0 Then
            WriteFact "Category losing most", ClassDelta(Key), "eur", CategoryLabel(CStr(Key)), _
                RangeDetail(DictAmount(BaseAgg("Class"), CStr(Key)), DictAmount(EndAgg("Class"), CStr(Key)))
        End If
    End If

    '
    ' Securities that arrived and left, cash aside
    '

    Set NewSecurities = NewTextDictionary()

    For Each Key In EndAgg("Security").Keys
        If StrComp(CStr(EndAgg("SecurityClass")(Key)), "Cash", vbTextCompare) <> 0 Then
            If Not BaseAgg("Security").Exists(Key) Then NewSecurities(Key) = EndAgg("Security")(Key)
        End If
    Next Key

    For Each Key In BaseAgg("Security").Keys
        If StrComp(CStr(BaseAgg("SecurityClass")(Key)), "Cash", vbTextCompare) <> 0 Then
            If Not EndAgg("Security").Exists(Key) Then
                GoneCount = GoneCount + 1
                GoneValue = GoneValue + CDbl(BaseAgg("Security")(Key))
            End If
        End If
    Next Key

    WriteFact "Securities new to the book", NewSecurities.Count, "int", "", _
        EuroText(SumOf(NewSecurities)) & " of collateral in securities nobody held before"

    SecKey = BestKey(NewSecurities, True)
    If SecKey <> "" Then
        WriteFact "Largest newcomer", NewSecurities(SecKey), "eur", SecurityText(EndAgg, SecKey), _
            "held by " & Plural(EndAgg("SecurityHolders")(SecKey).Count, "client", "clients")
    End If

    WriteFact "Securities gone from the book", GoneCount, "int", "", _
        EuroText(GoneValue) & " of collateral as they last stood"

    '
    ' Margin calls
    '

    WriteFact "New margin calls", NewMcCount, "int", "", "clients that stayed, flagged now and not before"
    WriteFact "Margin calls cleared", ClearedMcCount, "int", "", "clients that stayed, flagged before and not now"

End Sub

Private Function SumOf( _
    ByVal Dict As Object) As Double

    Dim Key As Variant

    For Each Key In Dict.Keys
        SumOf = SumOf + CDbl(Dict(Key))
    Next Key

End Function

'
' "from X to Y (+z%)"; a start of nothing reads as new, an end of nothing
' as gone.
'
Private Function RangeDetail( _
    ByVal StartValue As Double, _
    ByVal EndValue As Double) As String

    If Abs(StartValue) < FACT_TOLERANCE Then
        RangeDetail = "new; " & EuroText(EndValue) & " now"
    ElseIf Abs(EndValue) < FACT_TOLERANCE Then
        RangeDetail = "gone; " & EuroText(StartValue) & " before"
    Else
        RangeDetail = _
            EuroText(StartValue) & " to " & EuroText(EndValue) & " (" & _
            SignedPctText(SafeShare(EndValue - StartValue, StartValue)) & ")"
    End If

End Function

'
' A position's text from whichever snapshot still names its security.
'
Private Function PositionTextEither( _
    ByVal EndAgg As Object, _
    ByVal BaseAgg As Object, _
    ByVal PosKey As String) As String

    If EndAgg("SecurityName").Exists(PosKeySecurity(PosKey)) Then
        PositionTextEither = PositionText(EndAgg, PosKey)
    Else
        PositionTextEither = PositionText(BaseAgg, PosKey)
    End If

End Function

'====================================================================
' The whole run of snapshots on file
'====================================================================

Private Sub WriteHistorySection( _
    ByRef Dates As Variant, _
    ByRef EndSnap As FactSnapshot)

    Dim Accounts As Variant
    Dim Present As Object
    Dim WasPresent As Object
    Dim FirstSeen As Object
    Dim LastSeen As Object
    Dim Spells As Object
    Dim McCount As Object
    Dim McStreak As Object
    Dim McBest As Object
    Dim McBestEnd As Object
    Dim SfEver As Object
    Dim EndIndex As Object
    Dim Ages As Object
    Dim Lives As Object
    Dim Sorted As Variant

    Dim SnapshotDate As Date
    Dim NDG As String
    Dim Key As Variant
    Dim r As Long
    Dim i As Long

    Dim Mtm As Double
    Dim Drawn As Double
    Dim Loans As Long
    Dim RecordMtm As Double
    Dim RecordMtmDate As Date
    Dim LowMtm As Double
    Dim LowMtmDate As Date
    Dim RecordDrawn As Double
    Dim RecordDrawnDate As Date
    Dim LowDrawn As Double
    Dim LowDrawnDate As Date
    Dim RecordLoans As Long
    Dim RecordLoansDate As Date
    Dim LowLoans As Long
    Dim LowLoansDate As Date
    Dim RecordLine As Double
    Dim RecordLineNdg As String
    Dim RecordLineDate As Date

    Dim NewToday As Long
    Dim EndedToday As Long
    Dim BusiestNew As Long
    Dim BusiestNewDate As Date
    Dim BusiestEnded As Long
    Dim BusiestEndedDate As Date
    Dim LastNewDate As Date
    Dim LastEndedDate As Date
    Dim SnapshotsRead As Long
    Dim EndedCount As Long
    Dim ReturnedCount As Long
    Dim EverMc As Long
    Dim AgeSum As Double

    Set Present = NewTextDictionary()
    Set FirstSeen = NewTextDictionary()
    Set LastSeen = NewTextDictionary()
    Set Spells = NewTextDictionary()
    Set McCount = NewTextDictionary()
    Set McStreak = NewTextDictionary()
    Set McBest = NewTextDictionary()
    Set McBestEnd = NewTextDictionary()
    Set SfEver = NewTextDictionary()

    StartSection "The whole run on file", _
        DateText(Dates(LBound(Dates))) & " to " & DateText(EndSnap.AsOfDate) & ", " & _
        Plural(UBound(Dates) - LBound(Dates) + 1, "Accounts snapshot", "Accounts snapshots")

    For i = LBound(Dates) To UBound(Dates)

        SnapshotDate = Dates(i)

        If SnapshotDate = EndSnap.AsOfDate Then
            Accounts = EndSnap.Accounts
        Else
            Accounts = LoadFactAccounts(SnapshotDate)
        End If

        If FactDataHasRows(Accounts) Then

            SnapshotsRead = SnapshotsRead + 1
            Set WasPresent = Present
            Set Present = NewTextDictionary()

            Mtm = 0
            Drawn = 0
            Loans = 0
            NewToday = 0
            EndedToday = 0

            For r = LBound(Accounts, 1) To UBound(Accounts, 1)

                NDG = CStr(Accounts(r, FactAccNDG))
                Present(NDG) = True
                Loans = Loans + 1
                Mtm = Mtm + CDbl(Accounts(r, FactAccMTM))
                Drawn = Drawn + CDbl(Accounts(r, FactAccDrawn))

                If Not FirstSeen.Exists(NDG) Then FirstSeen(NDG) = SnapshotDate
                LastSeen(NDG) = SnapshotDate

                If Not WasPresent.Exists(NDG) Then
                    AddAmount Spells, NDG, 1
                    If SnapshotsRead > 1 Then NewToday = NewToday + 1
                End If

                If CBool(Accounts(r, FactAccMarginCall)) Then

                    AddAmount McCount, NDG, 1
                    AddAmount McStreak, NDG, 1

                    If CDbl(McStreak(NDG)) > DictAmount(McBest, NDG) Then
                        McBest(NDG) = McStreak(NDG)
                        McBestEnd(NDG) = SnapshotDate
                    End If

                Else

                    McStreak(NDG) = 0

                End If

                If CBool(Accounts(r, FactAccShortfall)) Then SfEver(NDG) = True

                If CDbl(Accounts(r, FactAccApproved)) > RecordLine Then
                    RecordLine = CDbl(Accounts(r, FactAccApproved))
                    RecordLineNdg = NDG
                    RecordLineDate = SnapshotDate
                End If

            Next r

            For Each Key In WasPresent.Keys
                If Not Present.Exists(Key) Then
                    EndedToday = EndedToday + 1
                    McStreak(Key) = 0
                End If
            Next Key

            If SnapshotsRead = 1 Or Mtm > RecordMtm Then
                RecordMtm = Mtm
                RecordMtmDate = SnapshotDate
            End If
            If SnapshotsRead = 1 Or Mtm < LowMtm Then
                LowMtm = Mtm
                LowMtmDate = SnapshotDate
            End If
            If SnapshotsRead = 1 Or Drawn > RecordDrawn Then
                RecordDrawn = Drawn
                RecordDrawnDate = SnapshotDate
            End If
            If SnapshotsRead = 1 Or Drawn < LowDrawn Then
                LowDrawn = Drawn
                LowDrawnDate = SnapshotDate
            End If
            If SnapshotsRead = 1 Or Loans > RecordLoans Then
                RecordLoans = Loans
                RecordLoansDate = SnapshotDate
            End If
            If SnapshotsRead = 1 Or Loans < LowLoans Then
                LowLoans = Loans
                LowLoansDate = SnapshotDate
            End If

            If NewToday > 0 Then
                LastNewDate = SnapshotDate
                If NewToday > BusiestNew Then
                    BusiestNew = NewToday
                    BusiestNewDate = SnapshotDate
                End If
            End If

            If EndedToday > 0 Then
                LastEndedDate = SnapshotDate
                If EndedToday > BusiestEnded Then
                    BusiestEnded = EndedToday
                    BusiestEndedDate = SnapshotDate
                End If
            End If

        End If

    Next i

    If SnapshotsRead = 0 Then
        WriteNoFact "No Accounts snapshot could be read."
        Exit Sub
    End If

    WriteFact "Snapshots read", SnapshotsRead, "int", "", _
        "Accounts only; the positions are read for the dates above alone"

    '
    ' The loans on the book today, by how long they have been there
    '

    Set EndIndex = AccountIndex(EndSnap.Accounts)
    Set Ages = NewTextDictionary()
    Set Lives = NewTextDictionary()

    For Each Key In FirstSeen.Keys

        If EndIndex.Exists(Key) Then
            Ages(Key) = CDbl(EndSnap.AsOfDate - CDate(FirstSeen(Key)))
            AgeSum = AgeSum + CDbl(Ages(Key))
        Else
            EndedCount = EndedCount + 1
            Lives(Key) = CDbl(CDate(LastSeen(Key)) - CDate(FirstSeen(Key)))
        End If

        If DictAmount(Spells, CStr(Key)) > 1 Then ReturnedCount = ReturnedCount + 1
        If DictAmount(McCount, CStr(Key)) > 0 Then EverMc = EverMc + 1

    Next Key

    NDG = BestKey(Ages, True)
    If NDG <> "" Then
        WriteFact "Oldest active loan", CDate(FirstSeen(NDG)), "date", NdgText(NDG), _
            "first seen; " & Plural(CLng(Ages(NDG)), "day", "days") & " on the book" & _
            SpellsText(Spells, NDG)
    End If

    NDG = BestKey(Ages, False)
    If NDG <> "" Then
        WriteFact "Youngest active loan", CDate(FirstSeen(NDG)), "date", NdgText(NDG), _
            "first seen; " & Plural(CLng(Ages(NDG)), "day", "days") & " on the book" & _
            SpellsText(Spells, NDG)
    End If

    If Ages.Count > 0 Then
        Sorted = SortedAmounts(Ages)
        WriteFact "Average age of an active loan", AgeSum / Ages.Count, "int", "", _
            "days since first seen; the median is " & Plural(CLng(MedianOf(Sorted)), "day", "days")
    End If

    '
    ' Comings and goings
    '

    WriteFact "Clients ever on the book", FirstSeen.Count, "int", "", _
        EndIndex.Count & " on it today"
    WriteFact "Loans ended since the first snapshot", EndedCount, "int"
    WriteFact "Loans that came back", ReturnedCount, "int", "", _
        "clients that left and reappeared"

    NDG = BestKey(Spells, True)
    If NDG <> "" Then
        If DictAmount(Spells, NDG) > 1 Then
            WriteFact "Most spells on the book", Spells(NDG), "int", NdgText(NDG), _
                "first seen " & DateText(CDate(FirstSeen(NDG))) & ", last " & _
                DateText(CDate(LastSeen(NDG)))
        End If
    End If

    NDG = BestKey(Lives, True)
    If NDG <> "" Then
        WriteFact "Longest-lived ended loan", Lives(NDG), "int", NdgText(NDG), _
            "days from " & DateText(CDate(FirstSeen(NDG))) & " to " & DateText(CDate(LastSeen(NDG)))
    End If

    NDG = BestKey(Lives, False)
    If NDG <> "" Then
        WriteFact "Shortest-lived ended loan", Lives(NDG), "int", NdgText(NDG), _
            IIf(CDbl(Lives(NDG)) = 0, "seen in one snapshot only, ", "days; ") & _
            DateText(CDate(FirstSeen(NDG))) & " to " & DateText(CDate(LastSeen(NDG)))
    End If

    WriteFact "Busiest snapshot for new loans", BusiestNew, "int", _
        IIf(BusiestNew > 0, DateText(BusiestNewDate), ""), _
        "new loans in one snapshot, the first snapshot aside"
    WriteFact "Busiest snapshot for ended loans", BusiestEnded, "int", _
        IIf(BusiestEnded > 0, DateText(BusiestEndedDate), ""), "loans ended in one snapshot"

    If LastNewDate > 0 Then
        WriteFact "Days since the last new loan", CLng(EndSnap.AsOfDate - LastNewDate), "int", _
            DateText(LastNewDate), "the last snapshot with a new NDG"
    End If

    If LastEndedDate > 0 Then
        WriteFact "Days since the last ended loan", CLng(EndSnap.AsOfDate - LastEndedDate), "int", _
            DateText(LastEndedDate), "the last snapshot an NDG went missing from"
    End If

    '
    ' Records
    '

    If RecordMtm > 0 Then

        WriteFact "Record collateral (MTM in Accounts)", RecordMtm, "eur", DateText(RecordMtmDate), _
            "today's " & EuroText(AccountSum(EndSnap.Accounts, FactAccMTM)) & " is " & _
            SignedPctText(SafeShare(AccountSum(EndSnap.Accounts, FactAccMTM), RecordMtm) - 1) & _
            " against it"
        WriteFact "Lowest collateral (MTM in Accounts)", LowMtm, "eur", DateText(LowMtmDate)

    End If

    WriteFact "Record drawn", RecordDrawn, "eur", DateText(RecordDrawnDate), _
        "today's " & EuroText(AccountSum(EndSnap.Accounts, FactAccDrawn)) & " is " & _
        SignedPctText(SafeShare(AccountSum(EndSnap.Accounts, FactAccDrawn), RecordDrawn) - 1) & _
        " against it"
    WriteFact "Lowest drawn", LowDrawn, "eur", DateText(LowDrawnDate)
    WriteFact "Most loans at once", RecordLoans, "int", DateText(RecordLoansDate), _
        EndIndex.Count & " today"
    WriteFact "Fewest loans at once", LowLoans, "int", DateText(LowLoansDate)

    If RecordLine > 0 Then
        WriteFact "Largest line ever approved", RecordLine, "eur", NdgText(RecordLineNdg), _
            "as of " & DateText(RecordLineDate) & _
            IIf(EndIndex.Exists(RecordLineNdg), "; still on the book", "; since ended")
    End If

    '
    ' Margin calls over the run
    '

    WriteFact "Clients ever in margin call", EverMc, "int", "", _
        PctText(SafeShare(EverMc, FirstSeen.Count)) & " of all clients ever on the book"
    WriteFact "Clients ever in shortfall", SfEver.Count, "int"

    NDG = BestKey(McCount, True, True)
    If NDG <> "" Then
        WriteFact "Most snapshots in margin call", McCount(NDG), "int", NdgText(NDG), _
            IIf(EndIndex.Exists(NDG), "still on the book", "since ended")
    End If

    NDG = BestKey(McBest, True, True)
    If NDG <> "" Then
        WriteFact "Longest margin call spell", McBest(NDG), "int", NdgText(NDG), _
            "snapshots in a row, ending " & DateText(CDate(McBestEnd(NDG)))
    End If

End Sub

Private Function SpellsText( _
    ByVal Spells As Object, _
    ByVal NDG As String) As String

    If DictAmount(Spells, NDG) > 1 Then
        SpellsText = " over " & Plural(CLng(DictAmount(Spells, NDG)), "spell", "spells")
    End If

End Function
