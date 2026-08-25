Attribute VB_Name = "Journey"
Option Explicit

' RELEASE: EXTERNAL_REPORT_PATH_20260820_V9
' Code!journey_start controls the first included snapshot.
' A blank journey_start includes all available snapshots; common European and
' ISO text-date formats are parsed without relying on the Windows locale.
' Margin-call reasons and comments come only from the external workbook stored
' in Code!report_path; the current workbook's Report sheet is never read.

Private Const JOURNEY_SHEET_NAME As String = "NDG Journey"
Private Const HOME_SHEET_NAME As String = "Home"
Private Const JOURNEY_NDG_RANGE As String = "JourneyNDG"
Private Const JOURNEY_START_RANGE As String = "journey_start"
Private Const REPORT_PATH_RANGE As String = "report_path"
Private Const REPORT_SHEET_NAME As String = "Report"

Private Const HEADER_APPROVED As String = "Max Approved Loan"
Private Const HEADER_DRAWN As String = "Drawn Amount"
Private Const HEADER_MTM As String = "MTM Collateral (MTM_t)"
Private Const HEADER_HCV As String = "Haircut Collateral Value (HCV_t)"
Private Const HEADER_MC As String = "Margin Call (HCV_t <= Max Approved Loan)"
Private Const HEADER_SF As String = "Shortfall (HTM_t <= Max Approved Loan)"
Private Const HEADER_MC_SINCE As String = "Alert since (as of date) - Margin Call"
Private Const HEADER_SF_SINCE As String = "Alert since (as of date) - Shortfall"
Private Const HEADER_MC_COMM As String = "Date of Margin call communication - since last Margin Call"
Private Const HEADER_SF_COMM As String = "Date of Shortfall communication - since last Shortfall"
Private Const HEADER_LIQUIDITY_WARN As String = "Date of Liquidation Warning communication"
Private Const HEADER_MC_SELL As String = "Sell Date - Margin Call"
Private Const HEADER_SF_SELL As String = "Sell Date - Shortfall"

Private Const EVENT_LOAN_ENDED As String = "Loan Ended"
Private Const EVENT_LOAN_RESTARTED As String = "Loan Restarted"

Public Sub ExtractNDGHistory()
    Dim PreviousScreenUpdating As Boolean
    Dim PreviousCalculation As XlCalculation
    Dim PreviousEnableEvents As Boolean
    Dim TargetNDG As String
    Dim JourneyStartDate As Date
    Dim HasJourneyStart As Boolean
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    PreviousScreenUpdating = Application.ScreenUpdating
    PreviousCalculation = Application.Calculation
    PreviousEnableEvents = Application.EnableEvents

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    TargetNDG = Trim$(CStr( _
        ThisWorkbook.Worksheets(HOME_SHEET_NAME) _
        .Range(JOURNEY_NDG_RANGE).Value))
    If TargetNDG = "" Then
        Err.Raise vbObjectError + 1600, "ExtractNDGHistory", _
                  "JourneyNDG is empty."
    End If

    JourneyStartDate = ReadJourneyStartDate(HasJourneyStart)

    BuildNDGHistorySheet _
        TargetNDG, HasJourneyStart, JourneyStartDate
    AddJourneyAnalysisColumns
    AppendMCReasonAndCommentToJourney TargetNDG
    ClearEndedLoanCaseDetails ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)
    RenameJourneyColumns
    FormatJourneyTable
    AddPositionAnalysisButtons ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)
    BuildJourneyDashboardTables True

CleanUp:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = PreviousScreenUpdating
    Application.Calculation = PreviousCalculation
    Application.EnableEvents = PreviousEnableEvents
    On Error GoTo 0

    If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Sub

Private Function ReadJourneyStartDate( _
    ByRef HasJourneyStart As Boolean) As Date

    Dim ValueIn As Variant
    Dim ParsedDate As Date

    ValueIn = ThisWorkbook.Worksheets(HOME_SHEET_NAME) _
        .Range(JOURNEY_START_RANGE).Value

    HasJourneyStart = False

    If IsError(ValueIn) Then
        Err.Raise vbObjectError + 1604, "ReadJourneyStartDate", _
                  "journey_start contains an Excel error."
    End If

    If IsEmpty(ValueIn) Then Exit Function
    If Len(Trim$(CStr(ValueIn))) = 0 Then Exit Function

    If Not TryParseJourneyDate(ValueIn, ParsedDate) Then
        Err.Raise vbObjectError + 1604, "ReadJourneyStartDate", _
                  "journey_start must be a valid Excel date, " & _
                  "dd.mm.yyyy, dd/mm/yyyy, or yyyy-mm-dd."
    End If

    ReadJourneyStartDate = DateSerial( _
        Year(ParsedDate), Month(ParsedDate), Day(ParsedDate))
    HasJourneyStart = True
End Function

Private Function TryParseJourneyDate( _
    ByVal ValueIn As Variant, _
    ByRef ParsedDate As Date) As Boolean

    Dim TextValue As String
    Dim DateParts As Variant
    Dim DayNumber As Long
    Dim MonthNumber As Long
    Dim YearNumber As Long

    On Error GoTo InvalidDate

    If VarType(ValueIn) = vbDate Then
        ParsedDate = CDate(ValueIn)
        TryParseJourneyDate = True
        Exit Function
    End If

    If IsNumeric(ValueIn) Then
        ParsedDate = CDate(CDbl(ValueIn))
        TryParseJourneyDate = True
        Exit Function
    End If

    TextValue = Trim$(CStr(ValueIn))
    TextValue = Replace(TextValue, "/", ".")
    TextValue = Replace(TextValue, "-", ".")
    DateParts = Split(TextValue, ".")

    If UBound(DateParts) - LBound(DateParts) <> 2 Then Exit Function
    If Not IsNumeric(Trim$(CStr(DateParts(0)))) Then Exit Function
    If Not IsNumeric(Trim$(CStr(DateParts(1)))) Then Exit Function
    If Not IsNumeric(Trim$(CStr(DateParts(2)))) Then Exit Function

    If Len(Trim$(CStr(DateParts(0)))) = 4 Then
        YearNumber = CLng(Trim$(CStr(DateParts(0))))
        MonthNumber = CLng(Trim$(CStr(DateParts(1))))
        DayNumber = CLng(Trim$(CStr(DateParts(2))))
    Else
        DayNumber = CLng(Trim$(CStr(DateParts(0))))
        MonthNumber = CLng(Trim$(CStr(DateParts(1))))
        YearNumber = CLng(Trim$(CStr(DateParts(2))))
    End If

    ParsedDate = DateSerial(YearNumber, MonthNumber, DayNumber)

    If Year(ParsedDate) <> YearNumber Then Exit Function
    If Month(ParsedDate) <> MonthNumber Then Exit Function
    If Day(ParsedDate) <> DayNumber Then Exit Function

    TryParseJourneyDate = True
    Exit Function

InvalidDate:
    TryParseJourneyDate = False
End Function

Private Sub BuildNDGHistorySheet( _
    ByVal TargetNDG As String, _
    ByVal HasJourneyStart As Boolean, _
    ByVal JourneyStartDate As Date)

    Dim BasePath As String
    Dim Files As Variant
    Dim FileName As String
    Dim Lines As Variant
    Dim Header As Variant
    Dim Arr As Variant
    Dim MatchedArr As Variant
    Dim SnapshotDate As Date
    Dim ws As Worksheet

    Dim idxNDG As Long
    Dim idxApproved As Long
    Dim idxDrawn As Long
    Dim idxMTM As Long
    Dim idxHCV As Long
    Dim idxMC As Long
    Dim idxSF As Long

    Dim HeaderWritten As Boolean
    Dim FoundNDG As Boolean
    Dim WasPresent As Boolean
    Dim HasSeenActiveLoan As Boolean
    Dim OutputRow As Long
    Dim PreviousOutputRow As Long
    Dim LifecycleEventCol As Long
    Dim SourceLastCol As Long
    Dim i As Long
    Dim r As Long
    Dim c As Long

    BasePath = PathSelection()
    If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"

    Files = GetSortedAccountFiles(BasePath)
    If IsEmpty(Files) Then
        Err.Raise vbObjectError + 1601, "BuildNDGHistorySheet", _
                  "No Account files found."
    End If

    Set ws = CreateOrReplaceSheet(JOURNEY_SHEET_NAME)
    OutputRow = 2

    For i = LBound(Files) To UBound(Files)
        FileName = CStr(Files(i))
        SnapshotDate = DateSerial(CLng(Left$(FileName, 4)), CLng(Mid$(FileName, 5, 2)), CLng(Mid$(FileName, 7, 2)))
        If HasJourneyStart Then
            If SnapshotDate < JourneyStartDate Then GoTo NextFile
        End If

        Lines = ReadAllLines(BasePath & FileName)
        If UBound(Lines) < 1 Then GoTo NextFile

        Header = Split(Lines(0), ";")
        idxNDG = RequiredAccountHeader(Header, "NDG", FileName)

        If Not HeaderWritten Then
            ws.Cells(1, 1).Value = "Snapshot Date"

            For c = LBound(Header) To UBound(Header)
                ws.Cells(1, c + 2).Value = Header(c)
            Next c

            SourceLastCol = UBound(Header) - LBound(Header) + 2
            LifecycleEventCol = SourceLastCol + 1
            ws.Cells(1, LifecycleEventCol).Value = "Lifecycle Event"

            idxApproved = RequiredAccountHeader(Header, HEADER_APPROVED, FileName)
            idxDrawn = RequiredAccountHeader(Header, HEADER_DRAWN, FileName)
            idxMTM = RequiredAccountHeader(Header, HEADER_MTM, FileName)
            idxHCV = RequiredAccountHeader(Header, HEADER_HCV, FileName)
            idxMC = RequiredAccountHeader(Header, HEADER_MC, FileName)
            idxSF = RequiredAccountHeader(Header, HEADER_SF, FileName)
            HeaderWritten = True
        End If

        FoundNDG = False

        For r = 1 To UBound(Lines)
            If Len(Trim$(Lines(r))) > 0 Then
                Arr = Split(Lines(r), ";")

                If UBound(Arr) >= idxNDG Then
                    If CLng(Val(Arr(idxNDG))) = CLng(Val(TargetNDG)) Then
                        MatchedArr = Arr
                        FoundNDG = True
                        Exit For
                    End If
                End If
            End If
        Next r

        If FoundNDG Then
            ws.Cells(OutputRow, 1).Value = SnapshotDate

            For c = LBound(MatchedArr) To UBound(MatchedArr)
                ws.Cells(OutputRow, c + 2).Value = MatchedArr(c)
            Next c

            If HasSeenActiveLoan And Not WasPresent Then
                ws.Cells(OutputRow, LifecycleEventCol).Value = EVENT_LOAN_RESTARTED & "; "
            End If

            PreviousOutputRow = OutputRow
            OutputRow = OutputRow + 1
            WasPresent = True
            HasSeenActiveLoan = True

        ElseIf WasPresent Then
            ws.Range(ws.Cells(OutputRow, 1), ws.Cells(OutputRow, SourceLastCol)).Value = _
                ws.Range(ws.Cells(PreviousOutputRow, 1), ws.Cells(PreviousOutputRow, SourceLastCol)).Value

            ws.Cells(OutputRow, 1).Value = SnapshotDate
            ws.Cells(OutputRow, idxApproved + 2).Value = 0
            ws.Cells(OutputRow, idxDrawn + 2).Value = 0
            ws.Cells(OutputRow, idxMTM + 2).Value = 0
            ws.Cells(OutputRow, idxHCV + 2).Value = 0
            ws.Cells(OutputRow, idxMC + 2).Value = 0
            ws.Cells(OutputRow, idxSF + 2).Value = 0
            ws.Cells(OutputRow, LifecycleEventCol).Value = EVENT_LOAN_ENDED & "; "

            PreviousOutputRow = OutputRow
            OutputRow = OutputRow + 1
            WasPresent = False
        End If

NextFile:
    Next i

    If Not HeaderWritten Then
        If HasJourneyStart Then
            Err.Raise vbObjectError + 1602, "BuildNDGHistorySheet", _
                      "No usable Accounts snapshot was found on or after " & _
                      Format$(JourneyStartDate, "dd/mm/yyyy") & "."
        Else
            Err.Raise vbObjectError + 1602, "BuildNDGHistorySheet", _
                      "No usable Accounts snapshot was found."
        End If
    End If

    If Not HasSeenActiveLoan Then
        Err.Raise vbObjectError + 1603, "BuildNDGHistorySheet", _
                  "NDG " & TargetNDG & " was not found in the available snapshots."
    End If
End Sub

Private Function RequiredAccountHeader(ByRef Header As Variant, _
                                       ByVal HeaderName As String, _
                                       ByVal FileName As String) As Long
    RequiredAccountHeader = FindHeaderIndex(Header, HeaderName)

    If RequiredAccountHeader < 0 Then
        Err.Raise vbObjectError + 1600, "BuildNDGHistorySheet", _
                  "Column '" & HeaderName & "' could not be found in:" & vbCrLf & FileName
    End If
End Function

Private Function GetSortedAccountFiles(ByVal BasePath As String) As Variant
    Dim Files As Collection
    Dim FileArray() As String
    Dim FileName As String
    Dim Temp As String
    Dim i As Long
    Dim j As Long

    Set Files = New Collection
    If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
    FileName = Dir(BasePath & "*_Lombard_Loans_ITA_Accounts.csv")

    Do While FileName <> ""
        Files.Add FileName
        FileName = Dir
    Loop

    If Files.Count = 0 Then Exit Function

    ReDim FileArray(1 To Files.Count)

    For i = 1 To Files.Count
        FileArray(i) = Files(i)
    Next i

    For i = 1 To UBound(FileArray) - 1
        For j = i + 1 To UBound(FileArray)
            If FileArray(i) > FileArray(j) Then
                Temp = FileArray(i)
                FileArray(i) = FileArray(j)
                FileArray(j) = Temp
            End If
        Next j
    Next i

    GetSortedAccountFiles = FileArray
End Function

Private Sub AddJourneyAnalysisColumns()
    Dim ws As Worksheet
    Dim DateCol As Long
    Dim LastRow As Long
    Dim LastSourceCol As Long
    Dim RowCount As Long

    Dim ApprovedCol As Long
    Dim DrawnCol As Long
    Dim MTMCol As Long
    Dim HCVCol As Long
    Dim MarginCallCol As Long
    Dim ShortfallCol As Long
    Dim LifecycleEventCol As Long

    Dim DeltaApprovedCol As Long
    Dim DeltaDrawnCol As Long
    Dim DeltaHCVCol As Long
    Dim LTVCol As Long
    Dim DeltaLTVCol As Long
    Dim MCClearedCol As Long
    Dim EventCol As Long

    Dim SourceData As Variant
    Dim Result() As Variant
    Dim CurrentApproved As Double
    Dim PreviousApproved As Double
    Dim CurrentDrawn As Double
    Dim PreviousDrawn As Double
    Dim CurrentMTM As Double
    Dim PreviousMTM As Double
    Dim CurrentHCV As Double
    Dim PreviousHCV As Double
    Dim CurrentMC As Double
    Dim PreviousMC As Double
    Dim CurrentSF As Double
    Dim PreviousSF As Double
    Dim CurrentLTV As Double
    Dim PreviousLTV As Double
    Dim LifecycleEvent As String
    Dim StandardEvent As String
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)
    DateCol = FindColumnByHeader(ws, "Snapshot Date")

    If DateCol = 0 Then
        Err.Raise vbObjectError + 1700, "AddJourneyAnalysisColumns", _
                  "The Snapshot Date column could not be found."
    End If

    LastRow = ws.Cells(ws.Rows.Count, DateCol).End(xlUp).Row

    If LastRow < 2 Then
        Err.Raise vbObjectError + 1700, "AddJourneyAnalysisColumns", _
                  JOURNEY_SHEET_NAME & " contains no data."
    End If

    ApprovedCol = FindColumnByHeader(ws, HEADER_APPROVED)
    DrawnCol = FindColumnByHeader(ws, HEADER_DRAWN)
    MTMCol = FindColumnByHeader(ws, HEADER_MTM)
    HCVCol = FindColumnByHeader(ws, HEADER_HCV)
    MarginCallCol = FindColumnByHeader(ws, HEADER_MC)
    ShortfallCol = FindColumnByHeader(ws, HEADER_SF)
    LifecycleEventCol = FindColumnByHeader(ws, "Lifecycle Event")

    If ApprovedCol = 0 Or DrawnCol = 0 Or MTMCol = 0 Or HCVCol = 0 Or _
       MarginCallCol = 0 Or ShortfallCol = 0 Or LifecycleEventCol = 0 Then
        Err.Raise vbObjectError + 1701, "AddJourneyAnalysisColumns", _
                  "One or more Journey analysis columns could not be found."
    End If

    LastSourceCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    DeltaApprovedCol = LastSourceCol + 1
    DeltaDrawnCol = DeltaApprovedCol + 1
    DeltaHCVCol = DeltaApprovedCol + 2
    LTVCol = DeltaApprovedCol + 3
    DeltaLTVCol = DeltaApprovedCol + 4
    MCClearedCol = DeltaApprovedCol + 5
    EventCol = DeltaApprovedCol + 6

    ws.Cells(1, DeltaApprovedCol).Value = "Delta Approved"
    ws.Cells(1, DeltaDrawnCol).Value = "Delta Drawn"
    ws.Cells(1, DeltaHCVCol).Value = "Delta HCV"
    ws.Cells(1, LTVCol).Value = "LTV"
    ws.Cells(1, DeltaLTVCol).Value = "Delta LTV"
    ws.Cells(1, MCClearedCol).Value = "MC Cleared"
    ws.Cells(1, EventCol).Value = "Event"

    RowCount = LastRow - 1
    SourceData = ws.Cells(2, 1).Resize(RowCount, LastSourceCol).Value2
    ReDim Result(1 To RowCount, 1 To 7)

    CurrentApproved = WorksheetDouble(SourceData(1, ApprovedCol))
    CurrentMTM = WorksheetDouble(SourceData(1, MTMCol))

    Result(1, 1) = 0
    Result(1, 2) = 0
    Result(1, 3) = 0
    If CurrentMTM > 0 Then Result(1, 4) = CurrentApproved / CurrentMTM Else Result(1, 4) = ""
    Result(1, 5) = ""
    Result(1, 6) = 0
    Result(1, 7) = "Initial Snapshot; " & Trim$(CStr(SourceData(1, LifecycleEventCol)))

    For r = 2 To RowCount
        CurrentApproved = WorksheetDouble(SourceData(r, ApprovedCol))
        PreviousApproved = WorksheetDouble(SourceData(r - 1, ApprovedCol))
        CurrentDrawn = WorksheetDouble(SourceData(r, DrawnCol))
        PreviousDrawn = WorksheetDouble(SourceData(r - 1, DrawnCol))
        CurrentMTM = WorksheetDouble(SourceData(r, MTMCol))
        PreviousMTM = WorksheetDouble(SourceData(r - 1, MTMCol))
        CurrentHCV = WorksheetDouble(SourceData(r, HCVCol))
        PreviousHCV = WorksheetDouble(SourceData(r - 1, HCVCol))
        CurrentMC = WorksheetDouble(SourceData(r, MarginCallCol))
        PreviousMC = WorksheetDouble(SourceData(r - 1, MarginCallCol))
        CurrentSF = WorksheetDouble(SourceData(r, ShortfallCol))
        PreviousSF = WorksheetDouble(SourceData(r - 1, ShortfallCol))
        LifecycleEvent = Trim$(CStr(SourceData(r, LifecycleEventCol)))

        Result(r, 1) = CurrentApproved - PreviousApproved
        Result(r, 2) = CurrentDrawn - PreviousDrawn
        Result(r, 3) = CurrentHCV - PreviousHCV

        If CurrentMTM > 0 Then
            CurrentLTV = CurrentApproved / CurrentMTM
            Result(r, 4) = CurrentLTV
        Else
            CurrentLTV = 0
            Result(r, 4) = ""
        End If

        If CurrentMTM > 0 And PreviousMTM > 0 Then
            PreviousLTV = PreviousApproved / PreviousMTM
            Result(r, 5) = CurrentLTV - PreviousLTV
        Else
            Result(r, 5) = ""
        End If

        Result(r, 6) = IIf(PreviousMC > 0 And CurrentMC = 0, 1, 0)

        StandardEvent = DetectJourneyEvent( _
            CurrentApproved, PreviousApproved, CurrentDrawn, PreviousDrawn, _
            CLng(CurrentMC), CLng(PreviousMC), CLng(CurrentSF), CLng(PreviousSF))

        If IsLoanEndedEvent(LifecycleEvent) Then
            Result(r, 7) = LifecycleEvent
        Else
            Result(r, 7) = LifecycleEvent & StandardEvent
        End If
    Next r

    ws.Cells(2, DeltaApprovedCol).Resize(RowCount, 7).Value = Result
    ws.Cells(2, LTVCol).Resize(RowCount).NumberFormat = "0.0%"
    ws.Cells(2, DeltaLTVCol).Resize(RowCount).NumberFormat = "+0.0%;-0.0%;-"
    ws.Columns(LifecycleEventCol).Delete
End Sub

Private Function DetectJourneyEvent(ByVal Approved As Double, ByVal PrevApproved As Double, _
                                    ByVal Drawn As Double, ByVal PrevDrawn As Double, _
                                    ByVal MarginCall As Long, ByVal PrevMarginCall As Long, _
                                    ByVal Shortfall As Long, ByVal PrevShortfall As Long) As String
    Dim Events As String

    If Approved > PrevApproved Then Events = Events & "Limit Increase; "
    If Approved < PrevApproved Then Events = Events & "Limit Reduction; "
    If Drawn > PrevDrawn Then Events = Events & "Drawn Increased; "
    If Drawn < PrevDrawn Then Events = Events & "Drawn Decreased; "
    If PrevMarginCall = 0 And MarginCall > 0 Then Events = Events & "Margin Call Triggered; "
    If PrevMarginCall > 0 And MarginCall = 0 Then Events = Events & "Margin Call Cleared; "
    If PrevShortfall = 0 And Shortfall > 0 Then Events = Events & "Shortfall Triggered; "
    If PrevShortfall > 0 And Shortfall = 0 Then Events = Events & "Shortfall Cleared; "

    DetectJourneyEvent = Events
End Function

Private Function GetExternalReportWorkbook( _
    ByRef OpenedByCode As Boolean) As Workbook

    Dim PathValue As Variant
    Dim ReportPath As String
    Dim CandidateWorkbook As Workbook

    OpenedByCode = False
    PathValue = ThisWorkbook.Worksheets(HOME_SHEET_NAME) _
        .Range(REPORT_PATH_RANGE).Value

    If IsError(PathValue) Then
        Err.Raise vbObjectError + 1210, "GetExternalReportWorkbook", _
                  "report_path contains an Excel error."
    End If

    ReportPath = Trim$(CStr(PathValue))
    If ReportPath = "" Then
        Err.Raise vbObjectError + 1210, "GetExternalReportWorkbook", _
                  "report_path is empty."
    End If

    If Len(ReportPath) >= 2 Then
        If Left$(ReportPath, 1) = Chr$(34) And _
           Right$(ReportPath, 1) = Chr$(34) Then
            ReportPath = Mid$(ReportPath, 2, Len(ReportPath) - 2)
        End If
    End If

    If StrComp(ReportPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1210, "GetExternalReportWorkbook", _
                  "report_path must point to an external workbook, " & _
                  "not the current workbook."
    End If

    For Each CandidateWorkbook In Application.Workbooks
        If StrComp(CandidateWorkbook.FullName, ReportPath, _
                   vbTextCompare) = 0 Then
            Set GetExternalReportWorkbook = CandidateWorkbook
            Exit Function
        End If
    Next CandidateWorkbook

    If Len(Dir$(ReportPath, _
                vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) = 0 Then
        Err.Raise vbObjectError + 1210, "GetExternalReportWorkbook", _
                  "The workbook specified by report_path was not found:" & _
                  vbCrLf & ReportPath
    End If

    Set GetExternalReportWorkbook = Application.Workbooks.Open( _
        FileName:=ReportPath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        Notify:=False, _
        AddToMru:=False)
    OpenedByCode = True
End Function

Private Sub AppendMCReasonAndCommentToJourney(ByVal TargetNDG As String)
    Const REPORT_DATE_COL As Long = 1
    Const REPORT_NDG_COL As Long = 2
    Const REPORT_START_DATE_COL As Long = 9
    Const REPORT_REASON_COL As Long = 12
    Const REPORT_COMMENT_COL As Long = 13

    Dim wbReport As Workbook
    Dim wsReport As Worksheet
    Dim CandidateSheet As Worksheet
    Dim wsJourney As Worksheet
    Dim ReportWorkbookOpenedByCode As Boolean
    Dim LastReportRow As Long
    Dim LastJourneyRow As Long
    Dim LastJourneyCol As Long
    Dim DateCol As Long
    Dim MCStartCol As Long
    Dim SFStartCol As Long
    Dim ReasonCol As Long
    Dim CommentCol As Long
    Dim JourneyData As Variant
    Dim DictRows As Object
    Dim RowsForCase As Collection
    Dim Key As String
    Dim ReportDataDate As Date
    Dim ReportNDG As String
    Dim ReportStartDate As Variant
    Dim ReportReason As String
    Dim ReportComment As String
    Dim SnapshotDate As Variant
    Dim r As Long
    Dim j As Variant
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    On Error GoTo CleanUp

    Set wbReport = GetExternalReportWorkbook( _
        ReportWorkbookOpenedByCode)

    For Each CandidateSheet In wbReport.Worksheets
        If StrComp(CandidateSheet.name, REPORT_SHEET_NAME, _
                   vbTextCompare) = 0 Then
            Set wsReport = CandidateSheet
            Exit For
        End If
    Next CandidateSheet

    If wsReport Is Nothing Then
        Err.Raise vbObjectError + 1211, _
                  "AppendMCReasonAndCommentToJourney", _
                  "Worksheet '" & REPORT_SHEET_NAME & _
                  "' was not found in:" & vbCrLf & wbReport.FullName
    End If

    Set wsJourney = ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)
    LastReportRow = GetLastRow(wsReport, "A")
    DateCol = FindColumnByHeader(wsJourney, "Snapshot Date")

    If DateCol = 0 Then
        Err.Raise vbObjectError + 1200, "AppendMCReasonAndCommentToJourney", _
                  "The Snapshot Date column could not be found."
    End If

    LastJourneyRow = wsJourney.Cells(wsJourney.Rows.Count, DateCol).End(xlUp).Row
    LastJourneyCol = wsJourney.Cells(1, wsJourney.Columns.Count).End(xlToLeft).Column
    MCStartCol = FindColumnByHeader(wsJourney, HEADER_MC_SINCE)
    SFStartCol = FindColumnByHeader(wsJourney, HEADER_SF_SINCE)
    ReasonCol = FindColumnByHeader(wsJourney, "Reason MC/SF")
    CommentCol = FindColumnByHeader(wsJourney, "Comment")

    If DateCol = 0 Or MCStartCol = 0 Or SFStartCol = 0 Then
        Err.Raise vbObjectError + 1200, "AppendMCReasonAndCommentToJourney", _
                  "Required Journey columns could not be found."
    End If

    If ReasonCol = 0 Then
        ReasonCol = LastJourneyCol + 1
        wsJourney.Cells(1, ReasonCol).Value = "Reason MC/SF"
        LastJourneyCol = ReasonCol
    End If

    If CommentCol = 0 Then
        CommentCol = LastJourneyCol + 1
        wsJourney.Cells(1, CommentCol).Value = "Comment"
        LastJourneyCol = CommentCol
    End If

    JourneyData = wsJourney.Range(wsJourney.Cells(1, 1), _
                                  wsJourney.Cells(LastJourneyRow, LastJourneyCol)).Value

    For r = 2 To LastJourneyRow
        JourneyData(r, ReasonCol) = ""
        JourneyData(r, CommentCol) = ""
    Next r

    Set DictRows = CreateObject("Scripting.Dictionary")

    For r = 2 To LastJourneyRow
        AddJourneyRowByDate DictRows, JourneyData(r, MCStartCol), r
        AddJourneyRowByDate DictRows, JourneyData(r, SFStartCol), r
    Next r

    For r = 1 To LastReportRow
        If IsDate(wsReport.Cells(r, REPORT_DATE_COL).Value) Then
            ReportDataDate = CDate(wsReport.Cells(r, REPORT_DATE_COL).Value)
        Else
            ReportNDG = SafeCellText(wsReport.Cells(r, REPORT_NDG_COL))
            If ReportNDG = "" Or Not IsNumeric(ReportNDG) Then GoTo NextReportRow
            If CLng(Val(ReportNDG)) <> CLng(Val(TargetNDG)) Then GoTo NextReportRow

            ReportStartDate = wsReport.Cells(r, REPORT_START_DATE_COL).Value
            If Not IsDate(ReportStartDate) Then GoTo NextReportRow

            ReportReason = SafeCellText(wsReport.Cells(r, REPORT_REASON_COL))
            ReportComment = SafeCellText(wsReport.Cells(r, REPORT_COMMENT_COL))
            If ReportReason = "" And ReportComment = "" Then GoTo NextReportRow

            Key = Format$(CDate(ReportStartDate), "yyyymmdd")

            If DictRows.Exists(Key) Then
                Set RowsForCase = DictRows(Key)

                For Each j In RowsForCase
                    SnapshotDate = JourneyData(j, DateCol)

                    If IsDate(SnapshotDate) Then
                        If CDate(SnapshotDate) < ReportDataDate Then
                            If ReportReason <> "" And Trim$(CStr(JourneyData(j, ReasonCol))) = "" Then
                                JourneyData(j, ReasonCol) = ReportReason
                            End If

                            If ReportComment <> "" And Trim$(CStr(JourneyData(j, CommentCol))) = "" Then
                                JourneyData(j, CommentCol) = ReportComment
                            End If
                        End If
                    End If
                Next j
            End If
        End If

NextReportRow:
    Next r

    wsJourney.Range(wsJourney.Cells(1, 1), _
                    wsJourney.Cells(LastJourneyRow, LastJourneyCol)).Value = JourneyData

CleanUp:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    On Error Resume Next
    If ReportWorkbookOpenedByCode Then
        wbReport.Close SaveChanges:=False
    End If
    On Error GoTo 0

    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription
    End If
End Sub

Private Sub AddJourneyRowByDate(ByVal RowsByDate As Object, ByVal Value As Variant, _
                                ByVal RowNumber As Long)
    Dim Key As String

    If Not IsDate(Value) Then Exit Sub

    Key = Format$(CDate(Value), "yyyymmdd")
    If Not RowsByDate.Exists(Key) Then Set RowsByDate(Key) = New Collection
    RowsByDate(Key).Add RowNumber
End Sub

Private Sub ClearEndedLoanCaseDetails(ByVal wsJourney As Worksheet)
    Dim ColumnsToClear As Variant
    Dim ColumnNumber As Variant
    Dim DateCol As Long
    Dim EventCol As Long
    Dim LastRow As Long
    Dim r As Long

    DateCol = FindColumnByHeader(wsJourney, "Snapshot Date")
    EventCol = FindColumnByHeader(wsJourney, "Event")
    If DateCol = 0 Or EventCol = 0 Then Exit Sub

    ColumnsToClear = Array( _
        FindColumnByHeader(wsJourney, "Reason MC/SF"), _
        FindColumnByHeader(wsJourney, "Comment"), _
        FindColumnByHeader(wsJourney, HEADER_MC_SINCE), _
        FindColumnByHeader(wsJourney, HEADER_SF_SINCE), _
        FindColumnByHeader(wsJourney, HEADER_MC_COMM), _
        FindColumnByHeader(wsJourney, HEADER_SF_COMM), _
        FindColumnByHeader(wsJourney, HEADER_LIQUIDITY_WARN), _
        FindColumnByHeader(wsJourney, HEADER_MC_SELL), _
        FindColumnByHeader(wsJourney, HEADER_SF_SELL))

    LastRow = wsJourney.Cells(wsJourney.Rows.Count, DateCol).End(xlUp).Row

    For r = 2 To LastRow
        If IsLoanEndedEvent(SafeCellText(wsJourney.Cells(r, EventCol))) Then
            For Each ColumnNumber In ColumnsToClear
                If ColumnNumber > 0 Then wsJourney.Cells(r, CLng(ColumnNumber)).ClearContents
            Next ColumnNumber
        End If
    Next r
End Sub

Public Function IsLoanEndedEvent(ByVal EventText As String) As Boolean
    IsLoanEndedEvent = InStr(1, EventText, EVENT_LOAN_ENDED, vbTextCompare) > 0
End Function

Public Function IsLoanRestartedEvent(ByVal EventText As String) As Boolean
    IsLoanRestartedEvent = InStr(1, EventText, EVENT_LOAN_RESTARTED, vbTextCompare) > 0
End Function

Private Sub RenameJourneyColumns()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)

    RenameHeader ws, HEADER_APPROVED, "Approved"
    RenameHeader ws, HEADER_DRAWN, "Drawn"
    RenameHeader ws, HEADER_MTM, "MTM"
    RenameHeader ws, HEADER_HCV, "HCV"
    RenameHeader ws, HEADER_MC, "MC"
    RenameHeader ws, HEADER_SF, "SF"
    RenameHeader ws, HEADER_MC_SINCE, "MC Since"
    RenameHeader ws, HEADER_SF_SINCE, "SF Since"
    RenameHeader ws, HEADER_MC_COMM, "MC Comm Date"
    RenameHeader ws, HEADER_SF_COMM, "SF Comm Date"
    RenameHeader ws, HEADER_LIQUIDITY_WARN, "Liquidity Warn Date"
    RenameHeader ws, HEADER_MC_SELL, "MC Sell Date"
    RenameHeader ws, HEADER_SF_SELL, "SF Sell Date"
    RenameHeader ws, "RM Code", "RM"
End Sub




