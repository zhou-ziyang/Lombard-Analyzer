Attribute VB_Name = "JourneyDashboardTable"
Option Explicit

' RELEASE: DASHBOARD_PCA_ALIGNED_HEADERS_F4_FREEZE_20260813_V6
' Customer Overview now follows the Position Change Analysis summary layout.
' The information-card and history-table headers share row 3; the sheet freezes
' at F4 without depending on the object that Excel currently has selected.

Private Const JOURNEY_SHEET_NAME As String = "NDG Journey"
Private Const DASHBOARD_SHEET_NAME As String = "NDG Dashboard"
Private Const LEGACY_TIME_SERIES_SHEET As String = "NDG Journey TS"
Private Const LEGACY_VISUAL_SHEET As String = "Journey Dashboard"

Private Const OVERVIEW_TOP_ROW As Long = 3
Private Const OVERVIEW_LEFT_COL As Long = 1
Private Const EVENT_TOP_ROW As Long = 3
Private Const EVENT_LEFT_COL As Long = 6
Private Const EVENT_COLUMN_COUNT As Long = 10

Private Const FIRST_DATA_ROW As Long = 2
Private Const VISUAL_HELPER_FIRST_COL As Long = 60
Private Const USE_DARK_THEME As Boolean = True

Private Const LABEL_MIN_WIDTH As Double = 16
Private Const LABEL_HEIGHT As Double = 11
Private Const LABEL_CHARACTER_WIDTH As Double = 8
Private Const LABEL_HORIZONTAL_MARGIN As Double = 0
Private Const COLLATERAL_DATE_SCALE As Double = 100
Private Const COLLATERAL_DATE_ORIGIN As Double = 1000
Private Const RISK_BAND_SHAPE_OFFSET As Double = 0
Private Const RISK_BAND_SHAPE_HEIGHT As Double = 7
Private Const RISK_BAND_LABEL_GAP As Double = 9
Private Const CHART_BOTTOM_MARGIN As Double = 5

Private Type JourneyColumns
    NDGCol As Long
    DateCol As Long
    EventCol As Long
    ReasonCol As Long
    ApprovedCol As Long
    DrawnCol As Long
    MTMCol As Long
    HCVCol As Long
    LTVCol As Long
    MCCol As Long
    SFCol As Long
    MCClearedCol As Long
    DeltaApprovedCol As Long
    DeltaDrawnCol As Long
    DeltaHCVCol As Long
    DeltaLTVCol As Long
End Type

Private ColorChartBackground As Long
Private ColorPlotBackground As Long
Private ColorChartText As Long
Private ColorBorder As Long
Private ColorGridline As Long
Private ColorMTM As Long
Private ColorHCV As Long
Private ColorApproved As Long
Private ColorApprovedArea As Long
Private ColorDrawn As Long
Private ColorLTV As Long
Private ColorLTVArea As Long
Private ColorHaircut As Long
Private ColorDanger As Long
Private ColorHCVOverlap As Long
Private ColorMTMOverlap As Long
Private ColorMCFill As Long
Private ColorMCLine As Long
Private ColorSFFill As Long
Private ColorInactiveFill As Long
Private ColorInactiveLine As Long
Private ColorThreshold As Long
Private ColorLabelBackground As Long
Private ColorLabelText As Long
Private ColorTechnicalLabelText As Long
Private ColorTechnicalLabelBorder As Long

Private TransparencyGridline As Double
Private TransparencyApprovedArea As Double
Private TransparencyLTVArea As Double
Private TransparencyHCVBand As Double
Private TransparencyMTMBand As Double
Private TransparencyHCVOverlap As Double
Private TransparencyMTMOverlap As Double
Private TransparencyRiskBand As Double
Private TransparencyRiskBandAccent As Double
Private TransparencyInactiveBand As Double
Private TransparencyInactiveLine As Double
Private TransparencyLabel As Double

Public Sub BuildJourneyDashboardTables(ByVal ActivateDashboard As Boolean)
    Dim PreviousScreenUpdating As Boolean
    Dim PreviousCalculation As XlCalculation
    Dim PreviousEnableEvents As Boolean
    Dim wsJourney As Worksheet
    Dim wsDashboard As Worksheet
    Dim DateCol As Long
    Dim NDGCol As Long
    Dim LastRow As Long
    Dim LatestRow As Long
    Dim HistoryLastRow As Long
    Dim StatusText As String
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

    Set wsJourney = ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME)
    DateCol = RequiredColumn( _
        wsJourney, "Snapshot Date", "BuildJourneyDashboardTables")
    LastRow = wsJourney.Cells(wsJourney.Rows.Count, DateCol).End(xlUp).Row

    If LastRow < 2 Then
        Err.Raise vbObjectError + 1500, "BuildJourneyDashboardTables", _
                  "No Journey data was found."
    End If

    LatestRow = FindLatestJourneyRow(wsJourney, LastRow, DateCol)
    NDGCol = RequiredColumn(wsJourney, "NDG", "BuildJourneyDashboardTables")
    StatusText = GetLoanStatusAtRow(wsJourney, LatestRow)

    Set wsDashboard = CreateOrReplaceSheet(DASHBOARD_SHEET_NAME)
    DeleteLegacyVisualisationSheets

    BuildCustomerOverviewTable wsJourney, wsDashboard, LatestRow
    HistoryLastRow = BuildHistoricalEventTable( _
        wsJourney, wsDashboard, LastRow)

    With wsDashboard.Range( _
            wsDashboard.Cells(1, OVERVIEW_LEFT_COL), _
            wsDashboard.Cells(1, EVENT_LEFT_COL + EVENT_COLUMN_COUNT - 1))
        .Merge
        .Value = "Lombard Loan Client Dashboard"
    End With
    ApplyUnifiedReportTitle wsDashboard.Range( _
        wsDashboard.Cells(1, OVERVIEW_LEFT_COL), _
        wsDashboard.Cells(1, EVENT_LEFT_COL + EVENT_COLUMN_COUNT - 1))

    With wsDashboard.Range( _
            wsDashboard.Cells(2, OVERVIEW_LEFT_COL), _
            wsDashboard.Cells(2, EVENT_LEFT_COL + EVENT_COLUMN_COUNT - 1))
        .Merge
        .Value = "NDG " & SafeCellText(wsJourney.Cells(LatestRow, NDGCol)) & _
                 " | Latest snapshot " & _
                 Format$(CDate(wsJourney.Cells(LatestRow, DateCol).Value), _
                         "dd/mm/yyyy") & _
                 " | " & StatusText
    End With
    ApplyUnifiedReportSubtitle wsDashboard.Range( _
        wsDashboard.Cells(2, OVERVIEW_LEFT_COL), _
        wsDashboard.Cells(2, EVENT_LEFT_COL + EVENT_COLUMN_COUNT - 1))

    AddJourneyVisualisation wsDashboard, HistoryLastRow + 2

    wsDashboard.Cells.Font.name = "Arial"
    wsDashboard.Columns(EVENT_LEFT_COL - 1).ColumnWidth = 3
    ApplyDashboardWindowLayout wsDashboard, ActivateDashboard

CleanUp:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = PreviousScreenUpdating
    Application.Calculation = PreviousCalculation
    Application.EnableEvents = PreviousEnableEvents
    On Error GoTo 0

    If ErrorNumber <> 0 Then _
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Sub

Private Sub ApplyDashboardWindowLayout(ByVal wsDashboard As Worksheet, _
                                       ByVal ActivateDashboard As Boolean)
    Dim PreviousSheet As Object
    Dim DashboardWindow As Window

    On Error Resume Next
    Set PreviousSheet = ActiveSheet
    On Error GoTo 0

    ' AddJourneyVisualisation can leave a chart object active.  Application.Goto
    ' explicitly returns Excel to a worksheet cell before the window is split.
    ThisWorkbook.Activate
    wsDashboard.Activate
    Application.GoTo _
        Reference:=wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL), _
        Scroll:=False

    Set DashboardWindow = ActiveWindow
    If DashboardWindow Is Nothing Then
        Err.Raise vbObjectError + 1501, "ApplyDashboardWindowLayout", _
                  "No workbook window is available for the Dashboard."
    End If

    With DashboardWindow
        .View = xlNormalView
        .DisplayGridlines = False
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0

        ' Both the information card and Historical Events use row 3 for
        ' their headers.  These values freeze A:E and rows 1:3 (F4).
        .SplitColumn = EVENT_LEFT_COL - 1
        .SplitRow = EVENT_TOP_ROW
        .FreezePanes = True
    End With

    wsDashboard.Range("A1").Select

    If Not ActivateDashboard Then
        If Not PreviousSheet Is Nothing Then PreviousSheet.Activate
    End If
End Sub

Private Sub BuildCustomerOverviewTable(ByVal wsJourney As Worksheet, _
                                       ByVal wsDashboard As Worksheet, _
                                       ByVal LatestRow As Long)
    Dim NDGCol As Long
    Dim MarketCol As Long
    Dim LegalEntityCol As Long
    Dim RMCol As Long
    Dim DateCol As Long
    Dim ApprovedCol As Long
    Dim DrawnCol As Long
    Dim MTMCol As Long
    Dim HCVCol As Long
    Dim MCCol As Long
    Dim SFCol As Long
    Dim ApprovedValue As Double
    Dim DrawnValue As Double
    Dim MTMValue As Double
    Dim HCVValue As Double
    Dim MCValue As Double
    Dim SFValue As Double
    Dim Utilisation As Double
    Dim StatusText As String
    Dim TitleRange As Range
    Dim TargetRange As Range
    Dim LabelRange As Range

    NDGCol = RequiredColumn(wsJourney, "NDG", "BuildCustomerOverviewTable")
    MarketCol = RequiredColumn(wsJourney, "Market", "BuildCustomerOverviewTable")
    LegalEntityCol = RequiredColumn( _
        wsJourney, "Legal Entity", "BuildCustomerOverviewTable")
    RMCol = RequiredColumn(wsJourney, "RM", "BuildCustomerOverviewTable")
    DateCol = RequiredColumn( _
        wsJourney, "Snapshot Date", "BuildCustomerOverviewTable")
    ApprovedCol = RequiredColumn( _
        wsJourney, "Approved", "BuildCustomerOverviewTable")
    DrawnCol = RequiredColumn(wsJourney, "Drawn", "BuildCustomerOverviewTable")
    MTMCol = RequiredColumn(wsJourney, "MTM", "BuildCustomerOverviewTable")
    HCVCol = RequiredColumn(wsJourney, "HCV", "BuildCustomerOverviewTable")
    MCCol = RequiredColumn(wsJourney, "MC", "BuildCustomerOverviewTable")
    SFCol = RequiredColumn(wsJourney, "SF", "BuildCustomerOverviewTable")

    ApprovedValue = ParseCsvDouble(wsJourney.Cells(LatestRow, ApprovedCol).Value2)
    DrawnValue = ParseCsvDouble(wsJourney.Cells(LatestRow, DrawnCol).Value2)
    MTMValue = ParseCsvDouble(wsJourney.Cells(LatestRow, MTMCol).Value2)
    HCVValue = ParseCsvDouble(wsJourney.Cells(LatestRow, HCVCol).Value2)
    MCValue = ParseCsvDouble(wsJourney.Cells(LatestRow, MCCol).Value2)
    SFValue = ParseCsvDouble(wsJourney.Cells(LatestRow, SFCol).Value2)
    If ApprovedValue <> 0 Then Utilisation = DrawnValue / ApprovedValue
    StatusText = GetLoanStatusAtRow(wsJourney, LatestRow)

    Set TitleRange = wsDashboard.Range( _
        wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL), _
        wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 3))
    TitleRange.UnMerge
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL).Value = _
        "Customer Overview"
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 1).Value = _
        "Value"
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 2).Value = _
        "Account Metric"
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 3).Value = _
        "Value"
    ApplyUnifiedSectionTitle TitleRange
    TitleRange.RowHeight = 18
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 1) _
        .HorizontalAlignment = xlCenter
    wsDashboard.Cells(OVERVIEW_TOP_ROW, OVERVIEW_LEFT_COL + 3) _
        .HorizontalAlignment = xlCenter

    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL) _
        .Resize(6).Value = Application.Transpose( _
            Array("NDG", "Market", "Legal Entity", "RM", _
                  "Latest Snapshot", "Current Status"))
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 1) _
        .Resize(6).Value = Application.Transpose( _
            Array( _
                wsJourney.Cells(LatestRow, NDGCol).Value, _
                wsJourney.Cells(LatestRow, MarketCol).Value, _
                wsJourney.Cells(LatestRow, LegalEntityCol).Value, _
                wsJourney.Cells(LatestRow, RMCol).Value, _
                wsJourney.Cells(LatestRow, DateCol).Value, _
                StatusText))
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 2) _
        .Resize(7).Value = Application.Transpose( _
            Array("Approved Limit", "Drawn", "Utilisation", _
                  "MTM Collateral", "HCV", "Margin Call", "Shortfall"))
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 3) _
        .Resize(7).Value = Application.Transpose( _
            Array(ApprovedValue, DrawnValue, Utilisation, MTMValue, _
                  HCVValue, MCValue, SFValue))

    Set TargetRange = wsDashboard.Range( _
        wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL), _
        wsDashboard.Cells(OVERVIEW_TOP_ROW + 7, OVERVIEW_LEFT_COL + 3))
    With TargetRange
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
        .Font.name = "Arial"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(38, 38, 38)
        .VerticalAlignment = xlCenter
        .Rows.RowHeight = 18
    End With

    With TargetRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
    With TargetRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
    With TargetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
    With TargetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
    With TargetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With

    Set LabelRange = Union( _
        wsDashboard.Range( _
            wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL), _
            wsDashboard.Cells(OVERVIEW_TOP_ROW + 7, OVERVIEW_LEFT_COL)), _
        wsDashboard.Range( _
            wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 2), _
            wsDashboard.Cells(OVERVIEW_TOP_ROW + 7, OVERVIEW_LEFT_COL + 2)))
    With LabelRange
        .Font.Bold = False
        .Font.Color = RGB(38, 38, 38)
        .HorizontalAlignment = xlLeft
    End With

    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 1) _
        .Resize(7).HorizontalAlignment = xlLeft
    With wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 3).Resize(7)
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With

    wsDashboard.Cells(OVERVIEW_TOP_ROW + 1, OVERVIEW_LEFT_COL + 3) _
        .Resize(2).NumberFormat = "#,##0;[Red]-#,##0;-"
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 3, OVERVIEW_LEFT_COL + 3) _
        .NumberFormat = "0.0%"
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 4, OVERVIEW_LEFT_COL + 3) _
        .Resize(4).NumberFormat = "#,##0;[Red]-#,##0;-"
    wsDashboard.Cells(OVERVIEW_TOP_ROW + 5, OVERVIEW_LEFT_COL + 1) _
        .NumberFormat = "dd/mm/yyyy"

    FormatStatusCell _
        wsDashboard.Cells(OVERVIEW_TOP_ROW + 6, OVERVIEW_LEFT_COL + 1), _
        StatusText

    wsDashboard.Columns(OVERVIEW_LEFT_COL).ColumnWidth = 17
    wsDashboard.Columns(OVERVIEW_LEFT_COL + 1).ColumnWidth = 17
    wsDashboard.Columns(OVERVIEW_LEFT_COL + 2).ColumnWidth = 18
    wsDashboard.Columns(OVERVIEW_LEFT_COL + 3).ColumnWidth = 18
End Sub

Private Function BuildHistoricalEventTable(ByVal wsJourney As Worksheet, _
                                           ByVal wsDashboard As Worksheet, _
                                           ByVal LastRow As Long) As Long
    Dim Headers As Variant
    Dim SourceHeaders As Variant
    Dim SourceCols() As Long
    Dim EventRows As Collection
    Dim OutputData() As Variant
    Dim TargetRange As Range
    Dim HistoryTable As ListObject
    Dim EventText As String
    Dim ApprovedDelta As Double
    Dim DrawnDelta As Double
    Dim DeltaApprovedCol As Long
    Dim DeltaDrawnCol As Long
    Dim SourceRow As Long
    Dim OutputRow As Long
    Dim LastOutputRow As Long
    Dim i As Long
    Dim c As Long

    Headers = Array( _
        "Position Analysis", "Date", "Event", "Approved", "Drawn", _
        "HCV", "MC", "SF", "Reason", "Comment")
    SourceHeaders = Array( _
        "Snapshot Date", "Event", "Approved", "Drawn", "HCV", _
        "MC", "SF", "Reason MC/SF", "Comment")
    ReDim SourceCols(LBound(SourceHeaders) To UBound(SourceHeaders))

    For i = LBound(SourceHeaders) To UBound(SourceHeaders)
        SourceCols(i) = RequiredColumn( _
            wsJourney, CStr(SourceHeaders(i)), "BuildHistoricalEventTable")
    Next i

    DeltaApprovedCol = RequiredColumn( _
        wsJourney, "Delta Approved", "BuildHistoricalEventTable")
    DeltaDrawnCol = RequiredColumn( _
        wsJourney, "Delta Drawn", "BuildHistoricalEventTable")

    Set EventRows = New Collection
    For SourceRow = 2 To LastRow
        EventText = SafeCellText(wsJourney.Cells(SourceRow, SourceCols(1)))
        If SourceRow = 2 Or EventText <> "" Or SourceRow = LastRow Then
            EventRows.Add SourceRow
        End If
    Next SourceRow

    ReDim OutputData(1 To EventRows.Count, 1 To EVENT_COLUMN_COUNT)
    For i = 1 To EventRows.Count
        SourceRow = CLng(EventRows(i))
        OutputData(i, 2) = wsJourney.Cells(SourceRow, SourceCols(0)).Value
        OutputData(i, 3) = SafeCellText( _
            wsJourney.Cells(SourceRow, SourceCols(1)))

        For c = 4 To 8
            OutputData(i, c) = ParseCsvDouble( _
                wsJourney.Cells(SourceRow, SourceCols(c - 2)).Value2)
        Next c

        OutputData(i, 9) = wsJourney.Cells(SourceRow, SourceCols(7)).Value
        OutputData(i, 10) = wsJourney.Cells(SourceRow, SourceCols(8)).Value
    Next i

    For i = LBound(Headers) To UBound(Headers)
        wsDashboard.Cells(EVENT_TOP_ROW, EVENT_LEFT_COL + i).Value = Headers(i)
    Next i

    wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL) _
        .Resize(EventRows.Count, EVENT_COLUMN_COUNT).Value = OutputData
    LastOutputRow = EVENT_TOP_ROW + EventRows.Count
    Set TargetRange = wsDashboard.Range( _
        wsDashboard.Cells(EVENT_TOP_ROW, EVENT_LEFT_COL), _
        wsDashboard.Cells( _
            LastOutputRow, EVENT_LEFT_COL + EVENT_COLUMN_COUNT - 1))
    Set HistoryTable = wsDashboard.ListObjects.Add( _
        SourceType:=xlSrcRange, Source:=TargetRange, _
        XlListObjectHasHeaders:=xlYes)
    HistoryTable.name = "tblDashboardHistory"
    HistoryTable.TableStyle = "TableStyleMedium2"

    For i = 1 To EventRows.Count
        SourceRow = CLng(EventRows(i))
        OutputRow = EVENT_TOP_ROW + i
        EventText = CStr(OutputData(i, 3))

        If IsLoanEndedEvent(EventText) Then
            With wsDashboard.Cells(OutputRow, EVENT_LEFT_COL + 3).Resize(1, 5)
                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(242, 242, 242)
                .Font.Color = RGB(89, 89, 89)
                .Font.Bold = False
            End With
            ApplyUnifiedSignalCell _
                wsDashboard.Cells(OutputRow, EVENT_LEFT_COL + 2), _
                RGB(217, 217, 217), RGB(89, 89, 89)
        Else
            ApprovedDelta = ParseCsvDouble( _
                wsJourney.Cells(SourceRow, DeltaApprovedCol).Value2)
            DrawnDelta = ParseCsvDouble( _
                wsJourney.Cells(SourceRow, DeltaDrawnCol).Value2)
            ApplyUnifiedDeltaHighlight _
                wsDashboard.Cells(OutputRow, EVENT_LEFT_COL + 3), ApprovedDelta
            ApplyUnifiedDeltaHighlight _
                wsDashboard.Cells(OutputRow, EVENT_LEFT_COL + 4), DrawnDelta

            If IsLoanRestartedEvent(EventText) Then
                ApplyUnifiedSignalCell _
                    wsDashboard.Cells(OutputRow, EVENT_LEFT_COL + 2), _
                    RGB(221, 235, 247), RGB(31, 78, 121)
            End If
        End If

        If SourceRow >= 3 Then
            AddPositionAnalysisButton _
                wsDashboard.Cells(OutputRow, EVENT_LEFT_COL), _
                wsJourney, SourceRow
        End If
    Next i

    With TargetRange.Rows(1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL + 1) _
            .Resize(EventRows.Count)
        .NumberFormat = "dd/mm/yyyy"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL + 3) _
            .Resize(EventRows.Count, 5)
        .NumberFormat = "#,##0;[Red]-#,##0;-"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    With wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL + 2) _
            .Resize(EventRows.Count)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With wsDashboard.Cells(EVENT_TOP_ROW + 1, EVENT_LEFT_COL + 8) _
            .Resize(EventRows.Count, 2)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    wsDashboard.Columns(EVENT_LEFT_COL).ColumnWidth = 12
    wsDashboard.Columns(EVENT_LEFT_COL + 1).ColumnWidth = 12
    wsDashboard.Columns(EVENT_LEFT_COL + 2).ColumnWidth = 28
    wsDashboard.Columns(EVENT_LEFT_COL + 3).ColumnWidth = 16
    wsDashboard.Columns(EVENT_LEFT_COL + 4).ColumnWidth = 16
    wsDashboard.Columns(EVENT_LEFT_COL + 5).ColumnWidth = 16
    wsDashboard.Columns(EVENT_LEFT_COL + 6).ColumnWidth = 14
    wsDashboard.Columns(EVENT_LEFT_COL + 7).ColumnWidth = 14
    wsDashboard.Columns(EVENT_LEFT_COL + 8).ColumnWidth = 22
    wsDashboard.Columns(EVENT_LEFT_COL + 9).ColumnWidth = 36
    TargetRange.Rows.AutoFit
    ApplyHistoricalEventsLeftDivider wsDashboard, LastOutputRow
    BuildHistoricalEventTable = LastOutputRow
End Function

Private Sub ApplyHistoricalEventsLeftDivider(ByVal wsDashboard As Worksheet, _
                                             ByVal LastOutputRow As Long)
    Dim DividerRange As Range

    Set DividerRange = wsDashboard.Range( _
        wsDashboard.Cells(EVENT_TOP_ROW, EVENT_LEFT_COL), _
        wsDashboard.Cells(LastOutputRow, EVENT_LEFT_COL))

    With DividerRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(68, 114, 196)
        .Weight = xlThin
    End With
End Sub

Private Sub DeleteLegacyVisualisationSheets()
    Dim PreviousDisplayAlerts As Boolean
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    PreviousDisplayAlerts = Application.DisplayAlerts
    On Error GoTo CleanUp
    Application.DisplayAlerts = False

    If SheetExists(LEGACY_TIME_SERIES_SHEET) Then _
        ThisWorkbook.Worksheets(LEGACY_TIME_SERIES_SHEET).Delete
    If SheetExists(LEGACY_VISUAL_SHEET) Then _
        ThisWorkbook.Worksheets(LEGACY_VISUAL_SHEET).Delete

CleanUp:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = PreviousDisplayAlerts
    On Error GoTo 0

    If ErrorNumber <> 0 Then _
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Sub

Private Function RequiredColumn(ByVal ws As Worksheet, _
                                ByVal HeaderName As String, _
                                ByVal ProcedureName As String) As Long
    RequiredColumn = FindColumnByHeader(ws, HeaderName)
    If RequiredColumn = 0 Then
        Err.Raise vbObjectError + 1501, ProcedureName, _
                  "Column '" & HeaderName & "' could not be found."
    End If
End Function

Private Function FindLatestJourneyRow(ByVal ws As Worksheet, _
                                      ByVal LastRow As Long, _
                                      ByVal DateCol As Long) As Long
    Dim LatestDate As Double
    Dim CurrentDate As Double
    Dim r As Long

    FindLatestJourneyRow = 2
    For r = 2 To LastRow
        If IsDate(ws.Cells(r, DateCol).Value) Then
            CurrentDate = CDbl(CDate(ws.Cells(r, DateCol).Value))
            If CurrentDate >= LatestDate Then
                LatestDate = CurrentDate
                FindLatestJourneyRow = r
            End If
        End If
    Next r
End Function

Private Sub FormatStatusCell(ByVal TargetCell As Range, ByVal StatusText As String)
    ApplyUnifiedStatusHighlight TargetCell, StatusText
End Sub

Private Function GetLoanStatusAtRow(ByVal wsJourney As Worksheet, _
                                    ByVal JourneyRow As Long) As String
    Dim EventCol As Long
    Dim MCCol As Long
    Dim SFCol As Long
    Dim EventText As String
    Dim MCValue As Double
    Dim SFValue As Double

    If JourneyRow < 2 Then
        GetLoanStatusAtRow = "No Data"
        Exit Function
    End If

    EventCol = RequiredColumn(wsJourney, "Event", "GetLoanStatusAtRow")
    MCCol = FindColumnByHeader(wsJourney, "MC")
    If MCCol = 0 Then
        MCCol = FindColumnByHeader( _
            wsJourney, "Margin Call (HCV_t <= Max Approved Loan)")
    End If

    SFCol = FindColumnByHeader(wsJourney, "SF")
    If SFCol = 0 Then
        SFCol = FindColumnByHeader( _
            wsJourney, "Shortfall (HTM_t <= Max Approved Loan)")
    End If

    If MCCol = 0 Or SFCol = 0 Then
        Err.Raise vbObjectError + 1800, "GetLoanStatusAtRow", _
                  "MC or SF column could not be found."
    End If

    EventText = SafeCellText(wsJourney.Cells(JourneyRow, EventCol))
    MCValue = ParseCsvDouble(wsJourney.Cells(JourneyRow, MCCol).Value2)
    SFValue = ParseCsvDouble(wsJourney.Cells(JourneyRow, SFCol).Value2)

    If IsLoanEndedEvent(EventText) Then
        GetLoanStatusAtRow = "Ended"
    ElseIf SFValue > 0 Then
        GetLoanStatusAtRow = "Shortfall"
    ElseIf MCValue > 0 Then
        GetLoanStatusAtRow = "Margin Call"
    Else
        GetLoanStatusAtRow = "Normal"
    End If
End Function

Private Sub AddJourneyVisualisation(ByVal wsDash As Worksheet, _
                                    ByVal ChartTopRow As Long)
    Dim wsData As Worksheet
    Dim Columns As JourneyColumns
    Dim LastRow As Long
    Dim ChartLastRow As Long
    Dim HelperLastColumn As Long
    Dim ChartLeft As Double
    Dim ChartTop As Double
    Dim ChartWidth As Double
    Dim MainChart As ChartObject
    Dim LTVChart As ChartObject

    InitialiseDashboardTheme
    Set wsData = wsDash
    BuildJourneyTimeSeries _
        ThisWorkbook.Worksheets(JOURNEY_SHEET_NAME), wsData, _
        Columns, LastRow, HelperLastColumn

    If LastRow < FIRST_DATA_ROW Then _
        Err.Raise vbObjectError + 1000, "AddJourneyVisualisation", _
                  JOURNEY_SHEET_NAME & " contains no data."

    ChartLastRow = LastRow
    If IsLoanEndedEvent(SafeCellText(wsData.Cells(LastRow, Columns.EventCol))) Then _
        ChartLastRow = LastRow - 1

    If ChartLastRow < FIRST_DATA_ROW Then _
        Err.Raise vbObjectError + 1001, "AddJourneyVisualisation", _
                  "No active Journey row is available for plotting."

    ChartLeft = wsDash.Cells(ChartTopRow, EVENT_LEFT_COL).Left
    ChartTop = wsDash.Cells(ChartTopRow, EVENT_LEFT_COL).Top
    ChartWidth = _
        wsDash.Cells(ChartTopRow, EVENT_LEFT_COL + EVENT_COLUMN_COUNT).Left - _
        ChartLeft

    Set MainChart = wsDash.ChartObjects.Add( _
        ChartLeft, ChartTop, ChartWidth, 320)
    Set LTVChart = wsDash.ChartObjects.Add( _
        ChartLeft, ChartTop + 315, ChartWidth, 320)

    BuildMainChart MainChart, wsData, Columns, ChartLastRow
    RebuildMainChartLegend MainChart
    PositionTitleAndLegend MainChart

    BuildLTVChart LTVChart, wsData, Columns, ChartLastRow
    ReserveRiskBandShapeSpace MainChart
    AlignChartPlotAreas MainChart, LTVChart

    AddRiskBandAccentShapes MainChart, wsData, Columns, ChartLastRow, _
                            AbovePlotArea:=False
    AddRiskBandAccentShapes LTVChart, wsData, Columns, ChartLastRow, _
                            AbovePlotArea:=True
    AddMCDurationLabels MainChart, wsData, Columns, ChartLastRow
    wsData.Range( _
        wsData.Cells(1, VISUAL_HELPER_FIRST_COL), _
        wsData.Cells(1, HelperLastColumn)).EntireColumn.Hidden = True
End Sub

Private Sub BuildJourneyTimeSeries(ByVal wsSource As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByRef Columns As JourneyColumns, _
                                   ByRef LastTargetRow As Long, _
                                   ByRef HelperLastColumn As Long)
    Dim SourceColumns As JourneyColumns
    Dim RowsByDate As Object
    Dim LastRow As Long
    Dim LastCol As Long
    Dim ColumnOffset As Long
    Dim SourceRow As Long
    Dim OutputRow As Long
    Dim CurrentDate As Date
    Dim EndDate As Date
    Dim DateKey As String
    Dim EventText As String
    Dim LoanIsActive As Boolean
    Dim HasSnapshot As Boolean
    Dim r As Long

    LastRow = GetLastRow(wsSource, "B")
    If LastRow < FIRST_DATA_ROW Then _
        Err.Raise vbObjectError + 1100, "BuildJourneyTimeSeries", _
                  JOURNEY_SHEET_NAME & " contains no data."

    LoadSourceColumns wsSource, SourceColumns
    LastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    ColumnOffset = VISUAL_HELPER_FIRST_COL - 1
    HelperLastColumn = ColumnOffset + LastCol
    ShiftJourneyColumns SourceColumns, Columns, ColumnOffset

    wsTarget.Range( _
        wsTarget.Cells(1, VISUAL_HELPER_FIRST_COL), _
        wsTarget.Cells(1, HelperLastColumn)).Value = _
        wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, LastCol)).Value

    Set RowsByDate = CreateObject("Scripting.Dictionary")
    For r = FIRST_DATA_ROW To LastRow
        If IsDate(wsSource.Cells(r, SourceColumns.DateCol).Value) Then
            RowsByDate( _
                DateKeyFor(wsSource.Cells(r, SourceColumns.DateCol).Value)) = r
        End If
    Next r

    CurrentDate = CDate( _
        wsSource.Cells(FIRST_DATA_ROW, SourceColumns.DateCol).Value)
    EndDate = CDate(wsSource.Cells(LastRow, SourceColumns.DateCol).Value)
    OutputRow = FIRST_DATA_ROW
    LoanIsActive = True

    Do While CurrentDate <= EndDate
        DateKey = DateKeyFor(CurrentDate)
        HasSnapshot = RowsByDate.Exists(DateKey)

        If HasSnapshot Then
            SourceRow = CLng(RowsByDate(DateKey))
            CopyRowValues _
                wsSource, SourceRow, 1, _
                wsTarget, OutputRow, VISUAL_HELPER_FIRST_COL, LastCol

            EventText = SafeCellText( _
                wsSource.Cells(SourceRow, SourceColumns.EventCol))
            If IsLoanEndedEvent(EventText) Then
                LoanIsActive = False
            ElseIf IsLoanRestartedEvent(EventText) Then
                LoanIsActive = True
            End If
        Else
            CopyRowValues _
                wsTarget, OutputRow - 1, VISUAL_HELPER_FIRST_COL, _
                wsTarget, OutputRow, VISUAL_HELPER_FIRST_COL, LastCol
            wsTarget.Cells(OutputRow, Columns.DateCol).Value = CurrentDate
            wsTarget.Cells(OutputRow, Columns.EventCol).ClearContents

            If Columns.MCClearedCol > 0 Then _
                wsTarget.Cells(OutputRow, Columns.MCClearedCol).Value = 0
        End If

        If Not LoanIsActive Then
            ClearColumns wsTarget, OutputRow, Array( _
                Columns.ApprovedCol, _
                Columns.DrawnCol, _
                Columns.MTMCol, _
                Columns.HCVCol, _
                Columns.LTVCol)
        End If

        If Not HasSnapshot Or Not LoanIsActive Then
            ClearColumns wsTarget, OutputRow, Array( _
                Columns.DeltaApprovedCol, _
                Columns.DeltaDrawnCol, _
                Columns.DeltaHCVCol, _
                Columns.DeltaLTVCol)
        End If

        OutputRow = OutputRow + 1
        CurrentDate = CurrentDate + 1
    Loop

    wsTarget.Columns(Columns.DateCol).NumberFormat = "dd/mm/yyyy"
    LastTargetRow = OutputRow - 1
End Sub

Private Sub BuildMainChart( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long)

    Dim ApprovedSeries As Series
    Dim RiskBandHeight As Double

    With ch.Chart
        .ChartType = xlLine
        .PlotVisibleOnly = False
        .DisplayBlanksAs = xlNotPlotted
        .HasTitle = True
        .ChartTitle.Text = _
            " Lombard Loan - NDG " & _
            SafeCellText(wsData.Cells(FIRST_DATA_ROW, Columns.NDGCol))
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
    End With

    AddLineSeries ch, wsData, Columns.DateCol, Columns.MTMCol, LastRow, _
                  "Collateral MV", ColorMTM
    AddLineSeries ch, wsData, Columns.DateCol, Columns.HCVCol, LastRow, _
                  "HCV", ColorHCV
'    AddLineSeries ch, wsData, Columns.DateCol, Columns.DrawnCol, LastRow, _
'                  "Drawn", ColorDrawn
    Set ApprovedSeries = AddLineSeries( _
        ch, wsData, Columns.DateCol, Columns.ApprovedCol, LastRow, _
        "Approved / LTV", ColorLTV, msoLineDash)
    AddAreaSeries ch, wsData, Columns.DateCol, Columns.ApprovedCol, _
                  LastRow, "Approved Area", _
                  ColorApprovedArea, TransparencyApprovedArea

    FormatAmountAxis ch, wsData, Columns, LastRow
    FormatDateAxis ch, wsData, Columns.DateCol, LastRow, False
    FormatJourneyChart ch

    AddCollateralBandSeries ch, wsData, Columns, LastRow
    NormaliseCollateralBandGroup ch
    RiskBandHeight = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
    AddRiskBandSeries ch, wsData, Columns, LastRow, RiskBandHeight
    MatchSecondaryAxisToPrimary ch

    ' Excel can regroup combo-chart series when area bands are added.
    FormatLineSeries ApprovedSeries, ColorLTV, msoLineDash
End Sub

Private Sub FormatJourneyChart(ByVal ch As ChartObject)
    FormatChartFrame ch
    FormatPrimaryGridlines ch
End Sub

Private Sub BuildLTVChart( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long)

    Dim LTVSeries As Series
    Dim HCVRatioSeries As Series
    Dim ThresholdSeries As Series
    Dim RiskBandHeight As Double

    With ch.Chart
        .ChartType = xlLine
        .PlotVisibleOnly = False
        .DisplayBlanksAs = xlNotPlotted
        .HasTitle = False
        .HasLegend = False
    End With

    Set LTVSeries = AddLineSeries( _
        ch, wsData, Columns.DateCol, Columns.LTVCol, LastRow, _
        "LTV", ColorLTV)
    AddAreaSeries ch, wsData, Columns.DateCol, Columns.LTVCol, _
                  LastRow, "LTV Area", ColorLTVArea, TransparencyLTVArea
    Set HCVRatioSeries = AddRatioLineSeries( _
        ch, wsData, Columns.DateCol, Columns.HCVCol, _
        Columns.MTMCol, LastRow, "HCV / MTM", ColorHCV)

    FormatDateAxis ch, wsData, Columns.DateCol, LastRow, True
    FormatLTVPrimaryAxis ch
    RiskBandHeight = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
    AddRiskBandSeries ch, wsData, Columns, LastRow, RiskBandHeight

    Set ThresholdSeries = ch.Chart.SeriesCollection.NewSeries
    With ThresholdSeries
        .name = "100% Threshold"
        .XValues = DataRange(wsData, Columns.DateCol, LastRow)
        .Values = ConstantValues(LastRow - FIRST_DATA_ROW + 1, 1)
    End With
    FormatLineSeries ThresholdSeries, ColorMTM

    ' Excel can regroup combo-chart series when area bands are added.
    FormatLineSeries LTVSeries, ColorLTV, msoLineDash
    FormatLineSeries HCVRatioSeries, ColorHCV
    With ch.Chart.Axes(xlValue, xlPrimary)
        .MaximumScaleIsAuto = False
        .MaximumScale = RiskBandHeight
    End With
    FormatChartFrame ch
End Sub

Private Sub AddAreaSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal ValueCol As Long, _
    ByVal LastRow As Long, _
    ByVal SeriesName As String, _
    ByVal FillColour As Long, _
    ByVal FillTransparency As Double)

    Dim AreaSeries As Series

    Set AreaSeries = ch.Chart.SeriesCollection.NewSeries
    With AreaSeries
        .name = SeriesName
        .XValues = DataRange(wsData, DateCol, LastRow)
        .Values = DataRange(wsData, ValueCol, LastRow)
        .ChartType = xlArea
        .AxisGroup = xlPrimary
        .Format.Line.Visible = msoFalse

        With .Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = FillColour
            .Transparency = FillTransparency
        End With
    End With
End Sub

Private Function AddLineSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal ValueCol As Long, _
    ByVal LastRow As Long, _
    ByVal SeriesName As String, _
    ByVal LineColour As Long, _
    Optional ByVal DashStyle As MsoLineDashStyle = msoLineSolid) As Series

    Set AddLineSeries = ch.Chart.SeriesCollection.NewSeries
    With AddLineSeries
        .name = SeriesName
        .XValues = DataRange(wsData, DateCol, LastRow)
        .Values = DataRange(wsData, ValueCol, LastRow)
    End With
    FormatLineSeries AddLineSeries, LineColour, DashStyle
End Function

Private Function AddRatioLineSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal NumeratorCol As Long, _
    ByVal DenominatorCol As Long, _
    ByVal LastRow As Long, _
    ByVal SeriesName As String, _
    ByVal LineColour As Long) As Series

    Dim RatioValues() As Variant
    Dim PointIndex As Long
    Dim SourceRow As Long
    Dim NumeratorValue As Variant
    Dim DenominatorValue As Variant

    ReDim RatioValues(1 To LastRow - FIRST_DATA_ROW + 1)

    For SourceRow = FIRST_DATA_ROW To LastRow
        PointIndex = SourceRow - FIRST_DATA_ROW + 1
        NumeratorValue = wsData.Cells(SourceRow, NumeratorCol).Value2
        DenominatorValue = wsData.Cells(SourceRow, DenominatorCol).Value2
        RatioValues(PointIndex) = CVErr(xlErrNA)

        If Not IsError(NumeratorValue) And _
           Not IsError(DenominatorValue) Then
            If IsNumeric(NumeratorValue) And _
               IsNumeric(DenominatorValue) Then
                If CDbl(DenominatorValue) <> 0 Then
                    RatioValues(PointIndex) = _
                        CDbl(NumeratorValue) / CDbl(DenominatorValue)
                End If
            End If
        End If
    Next SourceRow

    Set AddRatioLineSeries = ch.Chart.SeriesCollection.NewSeries
    With AddRatioLineSeries
        .name = SeriesName
        .XValues = DataRange(wsData, DateCol, LastRow)
        .Values = RatioValues
    End With
    FormatLineSeries AddRatioLineSeries, LineColour
End Function

Private Sub FormatLineSeries( _
    ByVal DataSeries As Series, _
    ByVal LineColour As Long, _
    Optional ByVal DashStyle As MsoLineDashStyle = msoLineSolid, _
    Optional ByVal LineWeight As Double = 1.25)

    With DataSeries
        .ChartType = xlLine
        .AxisGroup = xlPrimary
        .MarkerStyle = xlMarkerStyleNone
        With .Format.Line
            .Visible = msoTrue
            .Weight = LineWeight
            .ForeColor.RGB = LineColour
            .DashStyle = DashStyle
        End With
    End With
End Sub

Private Sub FormatAmountAxis( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long)

    Const APPROVED_MIN_POSITION As Double = 0.3

    Dim DataMinimum As Double
    Dim ApprovedMinimum As Double
    Dim AxisMinimum As Double
    Dim AxisMaximum As Double
    Dim TargetMinimum As Double

    DataMinimum = WorksheetFunction.Min( _
        DataRange(wsData, Columns.MTMCol, LastRow), _
        DataRange(wsData, Columns.HCVCol, LastRow), _
        DataRange(wsData, Columns.ApprovedCol, LastRow), _
        DataRange(wsData, Columns.DrawnCol, LastRow))

    ApprovedMinimum = WorksheetFunction.Min( _
        DataRange(wsData, Columns.ApprovedCol, LastRow))

    If DataMinimum >= 0 Then
        AxisMinimum = DataMinimum * 0.95
    Else
        AxisMinimum = DataMinimum * 1.05
    End If

    With ch.Chart.Axes(xlValue, xlPrimary)
        .MinimumScaleIsAuto = False
        .MinimumScale = AxisMinimum
        .MaximumScaleIsAuto = True
        .TickLabels.NumberFormat = "#,##0,"
        .HasTitle = True
        .AxisTitle.Text = "Amount (thousand Euros)"
    End With

    ch.Activate
    DoEvents

    With ch.Chart.Axes(xlValue, xlPrimary)
        AxisMaximum = .MaximumScale
        .MaximumScaleIsAuto = False
        .MaximumScale = AxisMaximum
    End With

    If AxisMaximum > ApprovedMinimum Then
        TargetMinimum = _
            (ApprovedMinimum - APPROVED_MIN_POSITION * AxisMaximum) / _
            (1 - APPROVED_MIN_POSITION)

        ' Non-negative data should not force the axis below zero.
        If DataMinimum >= 0 Then _
            TargetMinimum = Application.Max(0, TargetMinimum)

        If TargetMinimum < AxisMinimum Then
            ch.Chart.Axes(xlValue, xlPrimary).MinimumScale = TargetMinimum
        End If
    End If
End Sub

Private Sub FormatLTVPrimaryAxis(ByVal ch As ChartObject)
    Const MINIMUM_MAXIMUM As Double = 1
    Const MAXIMUM_MAXIMUM As Double = 2
    Dim AutoMaximum As Double

    With ch.Chart.Axes(xlValue, xlPrimary)
        .MinimumScaleIsAuto = False
        .MinimumScale = 0
        .MaximumScaleIsAuto = True
        .MajorUnitIsAuto = True
        .TickLabels.NumberFormat = "0%"
    End With

    ch.Activate
    DoEvents
    AutoMaximum = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale

    With ch.Chart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "LTV (% of Collateral MV)"
        .MaximumScaleIsAuto = False
        .MaximumScale = Application.Max( _
            MINIMUM_MAXIMUM, Application.Min(MAXIMUM_MAXIMUM, AutoMaximum))
        On Error Resume Next
        .MajorGridlines.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With
End Sub

Private Sub FormatDateAxis( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal LastRow As Long, _
    ByVal ShowLabels As Boolean)

    With ch.Chart.Axes(xlCategory, xlPrimary)
        .CategoryType = xlTimeScale
        .AxisBetweenCategories = True
        .BaseUnit = xlDays
        .MinimumScaleIsAuto = False
        .MinimumScale = CDbl(CDate(wsData.Cells(FIRST_DATA_ROW, DateCol).Value))
        .MaximumScaleIsAuto = False
        .MaximumScale = CDbl(CDate(wsData.Cells(LastRow, DateCol).Value))
        .TickLabels.NumberFormat = "dd-mmm-yyyy"
        If ShowLabels Then
            .TickLabelPosition = xlTickLabelPositionNextToAxis
        Else
            .TickLabelPosition = xlTickLabelPositionNone
        End If
    End With
End Sub

Private Sub FormatChartFrame(ByVal ch As ChartObject)
    With ch.Chart.ChartArea.Format
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = ColorChartBackground
        .Line.Visible = msoFalse
'        .Line.ForeColor.RGB = ColorBorder
    End With

    With ch.Chart.PlotArea.Format
        .Fill.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = ColorPlotBackground
        .Line.Visible = msoFalse
'        .Line.ForeColor.RGB = ColorBorder
    End With

    FormatChartText ch
End Sub

Private Sub FormatChartText(ByVal ch As ChartObject)
    On Error Resume Next

    With ch.Chart
        If .HasTitle Then .ChartTitle.Font.Color = ColorChartText
        If .HasLegend Then .Legend.Font.Color = ColorChartText

        With .Axes(xlValue, xlPrimary)
            .TickLabels.Font.Color = ColorChartText
            .Format.Line.ForeColor.RGB = ColorBorder
            If .HasTitle Then .AxisTitle.Font.Color = ColorChartText
        End With

        With .Axes(xlCategory, xlPrimary)
            .TickLabels.Font.Color = ColorChartText
            .Format.Line.ForeColor.RGB = ColorBorder
            If .HasTitle Then .AxisTitle.Font.Color = ColorChartText
        End With
    End With

    On Error GoTo 0
End Sub

Private Sub FormatPrimaryGridlines(ByVal ch As ChartObject)
    On Error Resume Next
    With ch.Chart.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = ColorGridline
        .Weight = 0.4
        .Transparency = TransparencyGridline
        .DashStyle = msoLineDash
    End With
    On Error GoTo 0
End Sub

Private Sub AddCollateralBandSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long)

    Dim BaseValues() As Variant
    Dim HCVOverlapValues() As Variant
    Dim HaircutRestValues() As Variant
    Dim MTMOverlapValues() As Variant
    Dim DangerRestValues() As Variant
    Dim Dates() As Variant
    Dim AxisMaximum As Double
    Dim OriginalPointCount As Long
    Dim MaximumPointCount As Long
    Dim OutputPointIndex As Long
    Dim SourceRow As Long
    Dim DateValue As Double
    Dim ApprovedValue As Double
    Dim MTMValue As Double
    Dim HCVValue As Double
    Dim NextDateValue As Double
    Dim NextApprovedValue As Double
    Dim NextMTMValue As Double
    Dim NextHCVValue As Double
    Dim HasApproved As Boolean
    Dim HasCollateral As Boolean
    Dim CrossingFractions(1 To 2) As Double
    Dim CrossingCount As Long
    Dim CrossingIndex As Long
    Dim CrossingValue As Double

    ch.Activate
    DoEvents
    AxisMaximum = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
    
    OriginalPointCount = LastRow - FIRST_DATA_ROW + 1
    If OriginalPointCount < 1 Then Exit Sub
    MaximumPointCount = OriginalPointCount * 3

    ReDim Dates(1 To MaximumPointCount)
    ReDim BaseValues(1 To MaximumPointCount)
    ReDim HCVOverlapValues(1 To MaximumPointCount)
    ReDim HaircutRestValues(1 To MaximumPointCount)
    ReDim MTMOverlapValues(1 To MaximumPointCount)
    ReDim DangerRestValues(1 To MaximumPointCount)

    For SourceRow = FIRST_DATA_ROW To LastRow
        DateValue = CDbl(CDate( _
            wsData.Cells(SourceRow, Columns.DateCol).Value))
        HasApproved = IsNumeric( _
            wsData.Cells(SourceRow, Columns.ApprovedCol).Value2)
        HasCollateral = _
            IsNumeric(wsData.Cells(SourceRow, Columns.MTMCol).Value2) _
            And IsNumeric(wsData.Cells(SourceRow, Columns.HCVCol).Value2)

        ApprovedValue = 0
        MTMValue = 0
        HCVValue = 0
        If HasApproved Then _
            ApprovedValue = CDbl( _
                wsData.Cells(SourceRow, Columns.ApprovedCol).Value2)
        If HasCollateral Then
            MTMValue = CDbl(wsData.Cells(SourceRow, Columns.MTMCol).Value2)
            HCVValue = CDbl(wsData.Cells(SourceRow, Columns.HCVCol).Value2)
        End If

        AppendCollateralBandPoint Dates, BaseValues, HCVOverlapValues, _
            HaircutRestValues, MTMOverlapValues, DangerRestValues, _
            OutputPointIndex, DateValue, HasApproved, ApprovedValue, _
            HasCollateral, MTMValue, HCVValue, AxisMaximum

        If SourceRow < LastRow And HasApproved And HasCollateral _
           And IsDate(wsData.Cells(SourceRow + 1, Columns.DateCol).Value) _
           And IsNumeric( _
               wsData.Cells(SourceRow + 1, Columns.ApprovedCol).Value2) _
           And IsNumeric( _
               wsData.Cells(SourceRow + 1, Columns.MTMCol).Value2) _
           And IsNumeric( _
               wsData.Cells(SourceRow + 1, Columns.HCVCol).Value2) Then

            NextDateValue = CDbl(CDate( _
                wsData.Cells(SourceRow + 1, Columns.DateCol).Value))
            NextApprovedValue = CDbl( _
                wsData.Cells(SourceRow + 1, Columns.ApprovedCol).Value2)
            NextMTMValue = CDbl( _
                wsData.Cells(SourceRow + 1, Columns.MTMCol).Value2)
            NextHCVValue = CDbl( _
                wsData.Cells(SourceRow + 1, Columns.HCVCol).Value2)

            CrossingCount = 0
            CrossingValue = LinearCrossingFraction( _
                ApprovedValue - HCVValue, _
                NextApprovedValue - NextHCVValue)
            If CrossingValue > 0 And CrossingValue < 1 Then
                CrossingCount = 1
                CrossingFractions(1) = CrossingValue
            End If

            CrossingValue = LinearCrossingFraction( _
                ApprovedValue - MTMValue, _
                NextApprovedValue - NextMTMValue)
            If CrossingValue > 0 And CrossingValue < 1 Then
                If CrossingCount = 0 Then
                    CrossingCount = 1
                    CrossingFractions(1) = CrossingValue
                ElseIf Abs(CrossingValue - CrossingFractions(1)) > 0.0000001 Then
                    CrossingCount = 2
                    CrossingFractions(2) = CrossingValue
                    If CrossingFractions(2) < CrossingFractions(1) Then
                        CrossingValue = CrossingFractions(1)
                        CrossingFractions(1) = CrossingFractions(2)
                        CrossingFractions(2) = CrossingValue
                    End If
                End If
            End If

            For CrossingIndex = 1 To CrossingCount
                CrossingValue = CrossingFractions(CrossingIndex)
                AppendCollateralBandPoint Dates, BaseValues, _
                    HCVOverlapValues, HaircutRestValues, _
                    MTMOverlapValues, DangerRestValues, OutputPointIndex, _
                    DateValue + CrossingValue * (NextDateValue - DateValue), _
                    True, ApprovedValue + CrossingValue * _
                        (NextApprovedValue - ApprovedValue), _
                    True, MTMValue + CrossingValue * _
                        (NextMTMValue - MTMValue), _
                    HCVValue + CrossingValue * _
                        (NextHCVValue - HCVValue), AxisMaximum
            Next CrossingIndex
        End If
    Next SourceRow

    ReDim Preserve Dates(1 To OutputPointIndex)
    ReDim Preserve BaseValues(1 To OutputPointIndex)
    ReDim Preserve HCVOverlapValues(1 To OutputPointIndex)
    ReDim Preserve HaircutRestValues(1 To OutputPointIndex)
    ReDim Preserve MTMOverlapValues(1 To OutputPointIndex)
    ReDim Preserve DangerRestValues(1 To OutputPointIndex)

    ScaleCollateralDates Dates, _
        ch.Chart.Axes(xlCategory, xlPrimary).MinimumScale

    AddStackedAreaSeries ch, Dates, _
                         BaseValues, "Collateral Base", _
                         ColorPlotBackground, 1, False
    AddStackedAreaSeries ch, Dates, _
                         HCVOverlapValues, "Approved in HCV Band", _
                         ColorHCVOverlap, TransparencyHCVOverlap, True
    AddStackedAreaSeries ch, Dates, _
                         HaircutRestValues, "Haircut", ColorHaircut, _
                         TransparencyHCVBand, True
    AddStackedAreaSeries ch, Dates, _
                         MTMOverlapValues, "Approved Above MTM", _
                         ColorMTMOverlap, TransparencyMTMOverlap, True
    AddStackedAreaSeries ch, Dates, _
                         DangerRestValues, "Above MTM Risk", ColorDanger, _
                         TransparencyMTMBand, True

    With ch.Chart.Axes(xlValue, xlPrimary)
        .MaximumScaleIsAuto = False
        .MaximumScale = AxisMaximum
    End With
End Sub

Private Sub ScaleCollateralDates( _
    ByRef Dates() As Variant, _
    ByVal MinimumDate As Double)

    Dim PointIndex As Long

    For PointIndex = LBound(Dates) To UBound(Dates)
        Dates(PointIndex) = COLLATERAL_DATE_ORIGIN + _
            WorksheetFunction.Round( _
                (CDbl(Dates(PointIndex)) - MinimumDate) * _
                COLLATERAL_DATE_SCALE, 0)
    Next PointIndex
End Sub

Private Sub AppendCollateralBandPoint( _
    ByRef Dates() As Variant, _
    ByRef BaseValues() As Variant, _
    ByRef HCVOverlapValues() As Variant, _
    ByRef HaircutRestValues() As Variant, _
    ByRef MTMOverlapValues() As Variant, _
    ByRef DangerRestValues() As Variant, _
    ByRef PointIndex As Long, _
    ByVal DateValue As Double, _
    ByVal HasApproved As Boolean, _
    ByVal ApprovedValue As Double, _
    ByVal HasCollateral As Boolean, _
    ByVal MTMValue As Double, _
    ByVal HCVValue As Double, _
    ByVal AxisMaximum As Double)

    Dim HaircutValue As Double
    Dim DangerValue As Double

    PointIndex = PointIndex + 1
    Dates(PointIndex) = DateValue

    If Not HasCollateral Then
        BaseValues(PointIndex) = CVErr(xlErrNA)
        HCVOverlapValues(PointIndex) = CVErr(xlErrNA)
        HaircutRestValues(PointIndex) = CVErr(xlErrNA)
        MTMOverlapValues(PointIndex) = CVErr(xlErrNA)
        DangerRestValues(PointIndex) = CVErr(xlErrNA)
        Exit Sub
    End If

    HaircutValue = Abs(MTMValue - HCVValue)
    DangerValue = Application.Max(0, AxisMaximum - MTMValue)
    BaseValues(PointIndex) = HCVValue

    If HasApproved Then
        HCVOverlapValues(PointIndex) = Application.Max(0, _
            Application.Min(HaircutValue, _
                Application.Min(ApprovedValue, MTMValue) - HCVValue))
        MTMOverlapValues(PointIndex) = Application.Max(0, _
            Application.Min(DangerValue, _
                Application.Min(ApprovedValue, AxisMaximum) - MTMValue))
    Else
        HCVOverlapValues(PointIndex) = 0
        MTMOverlapValues(PointIndex) = 0
    End If

    HaircutRestValues(PointIndex) = _
        HaircutValue - CDbl(HCVOverlapValues(PointIndex))
    DangerRestValues(PointIndex) = _
        DangerValue - CDbl(MTMOverlapValues(PointIndex))
End Sub

Private Function LinearCrossingFraction( _
    ByVal StartDifference As Double, _
    ByVal EndDifference As Double) As Double

    LinearCrossingFraction = -1
    If (StartDifference < 0 And EndDifference > 0) _
       Or (StartDifference > 0 And EndDifference < 0) Then _
        LinearCrossingFraction = _
            -StartDifference / (EndDifference - StartDifference)
End Function

Private Sub AddStackedAreaSeries( _
    ByVal ch As ChartObject, _
    ByVal XValues As Variant, _
    ByVal YValues As Variant, _
    ByVal SeriesName As String, _
    ByVal FillColour As Long, _
    ByVal FillTransparency As Double, _
    ByVal ShowFill As Boolean)

    Dim DataSeries As Series

    Set DataSeries = ch.Chart.SeriesCollection.NewSeries
    With DataSeries
        .name = SeriesName
        .XValues = XValues
        .Values = YValues
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .Format.Line.Visible = msoFalse
        With .Format.Fill
            .Visible = IIf(ShowFill, msoTrue, msoFalse)
            If ShowFill Then
                .Solid
                .ForeColor.RGB = FillColour
                .Transparency = FillTransparency
            End If
        End With
    End With
End Sub

Private Sub NormaliseCollateralBandGroup(ByVal ch As ChartObject)
    With ch.Chart.SeriesCollection("Collateral Base")
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .PlotOrder = 1
    End With
    With ch.Chart.SeriesCollection("Approved in HCV Band")
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .PlotOrder = 2
    End With
    With ch.Chart.SeriesCollection("Haircut")
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .PlotOrder = 3
    End With
    With ch.Chart.SeriesCollection("Approved Above MTM")
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .PlotOrder = 4
    End With
    With ch.Chart.SeriesCollection("Above MTM Risk")
        .ChartType = xlAreaStacked
        .AxisGroup = xlSecondary
        .PlotOrder = 5
    End With
    ch.Activate
    DoEvents
End Sub

Private Sub MatchSecondaryAxisToPrimary(ByVal ch As ChartObject)
    Dim MinimumValue As Double
    Dim MaximumValue As Double
    Dim MinimumDate As Double
    Dim MaximumDate As Double
    Dim AxisBetweenCategories As Boolean
    Dim DatePadding As Double

    With ch.Chart.Axes(xlValue, xlPrimary)
        MinimumValue = .MinimumScale
        MaximumValue = .MaximumScale
    End With

    With ch.Chart.Axes(xlCategory, xlPrimary)
        MinimumDate = .MinimumScale
        MaximumDate = .MaximumScale
        AxisBetweenCategories = .AxisBetweenCategories
    End With

    If AxisBetweenCategories Then DatePadding = 0.5

    ch.Chart.HasAxis(xlValue, xlSecondary) = True
    With ch.Chart.Axes(xlValue, xlSecondary)
        .MinimumScaleIsAuto = False
        .MinimumScale = MinimumValue
        .MaximumScaleIsAuto = False
        .MaximumScale = MaximumValue
        .HasTitle = False
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = msoFalse
        On Error Resume Next
        .MajorGridlines.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With

    ch.Chart.HasAxis(xlCategory, xlSecondary) = True
    With ch.Chart.Axes(xlCategory, xlSecondary)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .AxisBetweenCategories = False
        .MinimumScaleIsAuto = False
        .MinimumScale = COLLATERAL_DATE_ORIGIN - _
                        DatePadding * COLLATERAL_DATE_SCALE
        .MaximumScaleIsAuto = False
        .MaximumScale = COLLATERAL_DATE_ORIGIN + _
                        (MaximumDate - MinimumDate + DatePadding) * _
                        COLLATERAL_DATE_SCALE
        .HasTitle = False
        .TickLabelPosition = xlTickLabelPositionNone
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .Format.Line.Visible = msoFalse
        On Error Resume Next
        .MajorGridlines.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With
End Sub

Private Sub AddRiskBandSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long, _
    ByVal BandHeight As Double)

    AddInactiveLoanBandSeries ch, wsData, Columns, LastRow, BandHeight
    AddStatusBandSeries ch, wsData, Columns, LastRow, _
                        Columns.MCCol, Columns.MCClearedCol, "MC", _
                        ColorMCFill, BandHeight
    AddStatusBandSeries ch, wsData, Columns, LastRow, _
                        Columns.SFCol, 0, "SF", _
                        ColorSFFill, BandHeight
End Sub

Private Sub AddRiskBandAccentShapes( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long, _
    ByVal AbovePlotArea As Boolean)

    Dim Episodes As Collection

    ch.Activate
    DoEvents

    Set Episodes = GetInactiveLoanEpisodes(wsData, Columns, LastRow)
    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
                           ColorInactiveLine, AbovePlotArea

    Set Episodes = GetStatusEpisodes( _
        wsData, Columns.MCCol, Columns.MCClearedCol, LastRow)
    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
                           ColorMCFill, AbovePlotArea

    Set Episodes = GetStatusEpisodes( _
        wsData, Columns.SFCol, 0, LastRow)
    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
                           ColorSFFill, AbovePlotArea
End Sub

Private Sub AddEpisodeAccentShapes( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal Episodes As Collection, _
    ByVal FillColour As Long, _
    ByVal AbovePlotArea As Boolean)

    Dim Episode As Variant

    For Each Episode In Episodes
        AddRiskBandAccentShape ch, wsData, DateCol, _
            CLng(Episode(0)), CLng(Episode(1)), FillColour, _
            AbovePlotArea
    Next Episode
End Sub

Private Sub AddRiskBandAccentShape( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    ByVal FillColour As Long, _
    ByVal AbovePlotArea As Boolean)

    Dim AxisMinimum As Double
    Dim AxisMaximum As Double
    Dim AxisPadding As Double
    Dim VisibleMinimum As Double
    Dim VisibleMaximum As Double
    Dim StartDate As Double
    Dim EndDate As Double
    Dim ShapeLeft As Double
    Dim ShapeRight As Double
    Dim ShapeTop As Double
    Dim ShapeHeight As Double
    Dim shp As Shape

    With ch.Chart.Axes(xlCategory, xlPrimary)
        AxisMinimum = .MinimumScale
        AxisMaximum = .MaximumScale
        If .AxisBetweenCategories Then AxisPadding = 0.5
    End With

    VisibleMinimum = AxisMinimum - AxisPadding
    VisibleMaximum = AxisMaximum + AxisPadding
    If VisibleMaximum <= VisibleMinimum Then Exit Sub

    StartDate = CDbl(CDate(wsData.Cells(StartRow, DateCol).Value))
    EndDate = CDbl(CDate(wsData.Cells(EndRow, DateCol).Value))

    With ch.Chart.PlotArea
        ShapeLeft = .InsideLeft + _
            (StartDate - VisibleMinimum) / _
            (VisibleMaximum - VisibleMinimum) * .InsideWidth
        ShapeRight = .InsideLeft + _
            (EndDate - VisibleMinimum) / _
            (VisibleMaximum - VisibleMinimum) * .InsideWidth
        If AbovePlotArea Then
            ShapeTop = .InsideTop - RISK_BAND_SHAPE_OFFSET - _
                       RISK_BAND_SHAPE_HEIGHT
        Else
            ShapeTop = .InsideTop + .InsideHeight + _
                       RISK_BAND_SHAPE_OFFSET
        End If
    End With

    ShapeLeft = Application.Max( _
        ch.Chart.PlotArea.InsideLeft, ShapeLeft)
    ShapeRight = Application.Min( _
        ch.Chart.PlotArea.InsideLeft + ch.Chart.PlotArea.InsideWidth, _
        ShapeRight)
    If ShapeRight <= ShapeLeft Then ShapeRight = ShapeLeft + 1

    ShapeHeight = RISK_BAND_SHAPE_HEIGHT
    If ShapeTop < 1 Then _
        Err.Raise vbObjectError + 1300, "AddRiskBandAccentShape", _
                  "No chart space is available above the plot area."
    If ShapeTop + ShapeHeight > ch.Chart.ChartArea.Height - 1 Then _
        Err.Raise vbObjectError + 1300, "AddRiskBandAccentShape", _
                  "No chart space is available below the category axis."

    Set shp = ch.Chart.Shapes.AddShape( _
        msoShapeRectangle, ShapeLeft, ShapeTop, _
        ShapeRight - ShapeLeft, ShapeHeight)

    With shp
        .name = "RiskBand_Accent_" & _
                Format$(ch.Chart.Shapes.Count, "000")
        With .Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = FillColour
            .Transparency = TransparencyRiskBandAccent
        End With
        .Line.Visible = msoFalse
        .ZOrder msoBringToFront
    End With
End Sub

Private Sub AddStatusBandSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long, _
    ByVal ActiveCol As Long, _
    ByVal CloseCol As Long, _
    ByVal LabelPrefix As String, _
    ByVal FillColour As Long, _
    ByVal BandHeight As Double)

    Dim Episodes As Collection
    Dim Episode As Variant

    Set Episodes = GetStatusEpisodes(wsData, ActiveCol, CloseCol, LastRow)

    For Each Episode In Episodes
        AddSingleBandEpisodeSeries ch, wsData, Columns.DateCol, LastRow, _
            CLng(Episode(0)), CLng(Episode(1)), BandHeight, LabelPrefix, _
            FillColour, FillColour, TransparencyRiskBand, 1
    Next Episode
End Sub

Private Function GetStatusEpisodes( _
    ByVal wsData As Worksheet, _
    ByVal ActiveCol As Long, _
    ByVal CloseCol As Long, _
    ByVal LastRow As Long) As Collection

    Dim Result As New Collection
    Dim StartRow As Long
    Dim InBand As Boolean
    Dim r As Long

    For r = FIRST_DATA_ROW To LastRow
        If Not InBand Then
            If WorksheetDouble(wsData.Cells(r, ActiveCol).Value2) > 0 Then
                StartRow = r
                InBand = True
            End If
        End If

        If InBand And StatusEpisodeClosed(wsData, r, ActiveCol, CloseCol) Then
            Result.Add Array(StartRow, r, True)
            InBand = False
        End If
    Next r

    If InBand Then Result.Add Array(StartRow, LastRow, False)
    Set GetStatusEpisodes = Result
End Function

Private Function StatusEpisodeClosed( _
    ByVal wsData As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal ActiveCol As Long, _
    ByVal CloseCol As Long) As Boolean

    If CloseCol > 0 Then
        StatusEpisodeClosed = _
            (WorksheetDouble(wsData.Cells(RowNumber, CloseCol).Value2) = 1)
    Else
        StatusEpisodeClosed = _
            (WorksheetDouble(wsData.Cells(RowNumber, ActiveCol).Value2) = 0)
    End If
End Function

Private Function GetInactiveLoanEpisodes( _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long) As Collection

    Dim Result As New Collection
    Dim StartRow As Long
    Dim InInactivePeriod As Boolean
    Dim EventText As String
    Dim r As Long

    For r = FIRST_DATA_ROW To LastRow
        EventText = SafeCellText(wsData.Cells(r, Columns.EventCol))

        If Not InInactivePeriod Then
            If IsLoanEndedEvent(EventText) Then
                StartRow = r
                InInactivePeriod = True
            End If
        ElseIf IsLoanRestartedEvent(EventText) Then
            Result.Add Array(StartRow, r)
            InInactivePeriod = False
        End If
    Next r

    Set GetInactiveLoanEpisodes = Result
End Function

Private Sub AddInactiveLoanBandSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long, _
    ByVal BandHeight As Double)

    Dim Episodes As Collection
    Dim Episode As Variant

    Set Episodes = GetInactiveLoanEpisodes(wsData, Columns, LastRow)

    For Each Episode In Episodes
        AddSingleBandEpisodeSeries ch, wsData, Columns.DateCol, LastRow, _
            CLng(Episode(0)), CLng(Episode(1)), BandHeight, _
            "Inactive Episode", ColorInactiveFill, ColorInactiveLine, _
            TransparencyInactiveBand, TransparencyInactiveLine
    Next Episode
End Sub

Private Sub AddSingleBandEpisodeSeries( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByVal DateCol As Long, _
    ByVal LastRow As Long, _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    ByVal BandHeight As Double, _
    ByVal SeriesName As String, _
    ByVal FillColour As Long, _
    ByVal LineColour As Long, _
    ByVal FillTransparency As Double, _
    ByVal LineTransparency As Double)

    Dim BandValues() As Variant
    Dim DataSeries As Series
    Dim PointCount As Long
    Dim PointIndex As Long
    Dim SourceRow As Long

    PointCount = LastRow - FIRST_DATA_ROW + 1
    ReDim BandValues(1 To PointCount)

    For SourceRow = FIRST_DATA_ROW To LastRow
        PointIndex = SourceRow - FIRST_DATA_ROW + 1
        If SourceRow >= StartRow And SourceRow <= EndRow Then
            BandValues(PointIndex) = BandHeight
        Else
            BandValues(PointIndex) = CVErr(xlErrNA)
        End If
    Next SourceRow

    Set DataSeries = ch.Chart.SeriesCollection.NewSeries
    With DataSeries
        .name = SeriesName
        .XValues = DataRange(wsData, DateCol, LastRow)
        .Values = BandValues
        .AxisGroup = xlPrimary
        .ChartType = xlArea
        .MarkerStyle = xlMarkerStyleNone

        With .Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = FillColour
            .Transparency = FillTransparency
        End With

        With .Format.Line
            If LineTransparency >= 1 Then
                .Visible = msoFalse
            Else
                .Visible = msoTrue
                .ForeColor.RGB = LineColour
                .Transparency = LineTransparency
                .Weight = 1
                .DashStyle = msoLineSolid
            End If
        End With
    End With
End Sub

Private Sub AddMCDurationLabels( _
    ByVal ch As ChartObject, _
    ByVal wsData As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal LastRow As Long)

    Dim Episodes As Collection
    Dim Episode As Variant
    Dim DurationDays As Long
    Dim LabelText As String
    Dim ReasonRow As Long
    Dim IsTechnical As Boolean

    Set Episodes = GetStatusEpisodes( _
        wsData, Columns.MCCol, Columns.MCClearedCol, LastRow)

    ch.Activate
    DoEvents

    For Each Episode In Episodes
        DurationDays = DateDiff( _
            "d", _
            CDate(wsData.Cells(CLng(Episode(0)), Columns.DateCol).Value), _
            CDate(wsData.Cells(CLng(Episode(1)), Columns.DateCol).Value))

        If CBool(Episode(2)) Then
            LabelText = DurationDays & "d"
        Else
            LabelText = (DurationDays + 1) & "d (open)"
        End If

        ReasonRow = CLng(Episode(1))
        If CBool(Episode(2)) Then ReasonRow = ReasonRow - 1
        If ReasonRow < CLng(Episode(0)) Then ReasonRow = CLng(Episode(0))
        IsTechnical = IsTechnicalReason( _
            SafeCellText(wsData.Cells(ReasonRow, Columns.ReasonCol)))

        AddMCTextLabelByRows ch, LastRow, CLng(Episode(0)), _
                             CLng(Episode(1)), LabelText, IsTechnical
    Next Episode
End Sub

Private Sub AddMCTextLabelByRows( _
    ByVal ch As ChartObject, _
    ByVal LastRow As Long, _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    ByVal LabelText As String, _
    ByVal IsTechnical As Boolean)

    Dim PointCount As Long
    Dim MidPoint As Double
    Dim RelativeX As Double
    Dim XPos As Double
    Dim YPos As Double
    Dim PlotLeft As Double
    Dim PlotRight As Double
    Dim LabelWidth As Double
    Dim LabelFillColour As Long
    Dim LabelFontColour As Long
    Dim LabelFillTransparency As Double
    Dim shp As Shape

    PointCount = LastRow - FIRST_DATA_ROW + 1
    If PointCount < 1 Then Exit Sub

    MidPoint = ((StartRow - 1) + (EndRow - 1)) / 2

    If ch.Chart.Axes(xlCategory, xlPrimary).AxisBetweenCategories Then
        RelativeX = (MidPoint - 0.5) / PointCount
    ElseIf PointCount > 1 Then
        RelativeX = (MidPoint - 1) / (PointCount - 1)
    Else
        RelativeX = 0.5
    End If
    RelativeX = Application.Max(0, Application.Min(1, RelativeX))

    With ch.Chart.PlotArea
        PlotLeft = .InsideLeft
        PlotRight = .InsideLeft + .InsideWidth
        XPos = .InsideLeft + RelativeX * .InsideWidth
        YPos = .InsideTop + .InsideHeight + _
               RISK_BAND_SHAPE_OFFSET + RISK_BAND_SHAPE_HEIGHT + _
               RISK_BAND_LABEL_GAP
    End With

    If YPos + LABEL_HEIGHT > _
       ch.Chart.ChartArea.Height - CHART_BOTTOM_MARGIN Then _
        Err.Raise vbObjectError + 1301, "AddMCTextLabelByRows", _
                  "No chart space is available below the risk band strip."

    LabelWidth = Application.Min( _
        PlotRight - PlotLeft, _
        Application.Max( _
            LABEL_MIN_WIDTH, _
            Len(LabelText) * LABEL_CHARACTER_WIDTH + _
            2 * LABEL_HORIZONTAL_MARGIN))

    XPos = Application.Max( _
        PlotLeft + LabelWidth / 2, _
        Application.Min(PlotRight - LabelWidth / 2, XPos))

    If IsTechnical Then
        LabelFontColour = ColorTechnicalLabelText
    Else
        LabelFillColour = ColorLabelBackground
        LabelFontColour = ColorLabelText
        LabelFillTransparency = TransparencyLabel
    End If

    Set shp = ch.Chart.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, XPos - LabelWidth / 2, YPos, _
        LabelWidth, LABEL_HEIGHT)

    With shp
        .name = "MC_Duration_Label_" & Format$(ch.Chart.Shapes.Count, "000")
        .Visible = msoTrue
        .TextFrame.Characters.Text = LabelText
        .TextFrame.MarginLeft = LABEL_HORIZONTAL_MARGIN
        .TextFrame.MarginRight = LABEL_HORIZONTAL_MARGIN
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        With .TextFrame.Characters.Font
            .Bold = True
            .Size = 9
            .Color = LabelFontColour
        End With
        With .Fill
            If IsTechnical Then
                .Visible = msoFalse
            Else
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = LabelFillColour
                .Transparency = LabelFillTransparency
            End If
        End With
        With .Line
            If IsTechnical Then
                .Visible = msoTrue
                .ForeColor.RGB = ColorTechnicalLabelBorder
                .Transparency = 0
                .Weight = 0.5
'                .DashStyle = msoLineRoundDot
            Else
                .Visible = msoFalse
            End If
        End With
        .ZOrder msoBringToFront
    End With
End Sub

Private Sub RebuildMainChartLegend(ByVal ch As ChartObject)
    With ch.Chart
        .HasLegend = False
        .HasLegend = True
    End With

    ch.Activate
    DoEvents

    RemoveHelperAreaLegendEntries ch
    RemoveDuplicateRiskBandLegendEntries ch
End Sub

Private Sub RemoveHelperAreaLegendEntries(ByVal ch As ChartObject)
    Dim i As Long

    For i = ch.Chart.Legend.LegendEntries.Count To 1 Step -1
        If IsHelperAreaLegendEntry( _
                ch.Chart.Legend.LegendEntries(i)) Then _
            ch.Chart.Legend.LegendEntries(i).Delete
    Next i
End Sub

Private Function IsHelperAreaLegendEntry(ByVal Entry As Object) As Boolean
    On Error GoTo ExitFunction

    With Entry.LegendKey.Format
        If .Fill.Visible <> msoTrue And .Line.Visible <> msoTrue Then
            IsHelperAreaLegendEntry = True
        ElseIf .Fill.Visible = msoTrue And .Line.Visible <> msoTrue Then
            Select Case .Fill.ForeColor.RGB
                Case ColorApprovedArea, ColorHaircut, ColorDanger, _
                     ColorHCVOverlap, ColorMTMOverlap
                    IsHelperAreaLegendEntry = True
            End Select
        End If
    End With

ExitFunction:
End Function

Private Sub RemoveDuplicateRiskBandLegendEntries(ByVal ch As ChartObject)
    Dim SeenKeys As Object
    Dim LegendKey As String
    Dim i As Long

    Set SeenKeys = CreateObject("Scripting.Dictionary")

    For i = ch.Chart.Legend.LegendEntries.Count To 1 Step -1
        LegendKey = RiskBandLegendKey(ch.Chart.Legend.LegendEntries(i))

        If Len(LegendKey) > 0 Then
            If SeenKeys.Exists(LegendKey) Then
                ch.Chart.Legend.LegendEntries(i).Delete
            Else
                SeenKeys.Add LegendKey, True
            End If
        End If
    Next i
End Sub

Private Function RiskBandLegendKey(ByVal Entry As Object) As String
    Dim FillColour As Long

    On Error GoTo ExitFunction
    With Entry.LegendKey.Format.Fill
        If .Visible <> msoTrue Then Exit Function
        FillColour = .ForeColor.RGB
    End With

    Select Case FillColour
        Case ColorMCFill, ColorSFFill, ColorInactiveFill
            RiskBandLegendKey = CStr(FillColour)
    End Select

ExitFunction:
End Function

Private Sub ReserveRiskBandShapeSpace(ByVal ch As ChartObject)
    Dim RequiredBottomSpace As Double
    Dim MaximumInsideHeight As Double
    Dim Pass As Long

    RequiredBottomSpace = RISK_BAND_SHAPE_OFFSET + _
                          RISK_BAND_SHAPE_HEIGHT + _
                          RISK_BAND_LABEL_GAP + LABEL_HEIGHT + _
                          CHART_BOTTOM_MARGIN

    For Pass = 1 To 2
        ch.Activate
        DoEvents

        With ch.Chart.PlotArea
            MaximumInsideHeight = ch.Chart.ChartArea.Height - _
                                  .InsideTop - RequiredBottomSpace
            If MaximumInsideHeight <= 0 Then _
                Err.Raise vbObjectError + 1299, _
                          "ReserveRiskBandShapeSpace", _
                          "The main chart is too short for the risk band strip."
            If .InsideHeight > MaximumInsideHeight Then _
                .InsideHeight = MaximumInsideHeight
        End With
    Next Pass
End Sub

Private Sub AlignChartPlotAreas( _
    ByVal UpperChart As ChartObject, _
    ByVal LowerChart As ChartObject)

    Dim TargetInsideLeft As Double
    Dim TargetInsideWidth As Double
    Dim Pass As Long

    UpperChart.Activate
    DoEvents
    LowerChart.Activate
    DoEvents

    TargetInsideLeft = UpperChart.Chart.PlotArea.InsideLeft
    TargetInsideWidth = UpperChart.Chart.PlotArea.InsideWidth

    For Pass = 1 To 2
        With LowerChart.Chart.PlotArea
            .Left = .Left + TargetInsideLeft - .InsideLeft
            .Width = .Width + TargetInsideWidth - .InsideWidth
        End With
        LowerChart.Activate
        DoEvents
    Next Pass
End Sub

Private Sub PositionTitleAndLegend(ByVal ch As ChartObject)
    Const OUTER_MARGIN As Double = 12
    Const VERTICAL_MARGIN As Double = 6

    With ch.Chart
        .HasTitle = True
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop

        ch.Activate
        DoEvents

        .ChartTitle.Left = OUTER_MARGIN
        .ChartTitle.Top = VERTICAL_MARGIN
        .Legend.Left = .ChartArea.Width - .Legend.Width - OUTER_MARGIN
        .Legend.Top = VERTICAL_MARGIN
    End With

    FormatChartText ch
End Sub

Private Sub LoadSourceColumns( _
    ByVal ws As Worksheet, _
    ByRef Columns As JourneyColumns)

    LoadCoreColumns ws, Columns, "BuildJourneyTimeSeries"
    Columns.LTVCol = RequiredColumn(ws, "LTV", "BuildJourneyTimeSeries")
    Columns.MCCol = RequiredColumn(ws, "MC", "BuildJourneyTimeSeries")
    Columns.SFCol = RequiredColumn(ws, "SF", "BuildJourneyTimeSeries")
    Columns.MCClearedCol = _
        RequiredColumn(ws, "MC Cleared", "BuildJourneyTimeSeries")
    Columns.ReasonCol = _
        RequiredColumn(ws, "Reason MC/SF", "BuildJourneyTimeSeries")
    Columns.DeltaApprovedCol = _
        RequiredColumn(ws, "Delta Approved", "BuildJourneyTimeSeries")
    Columns.DeltaDrawnCol = _
        RequiredColumn(ws, "Delta Drawn", "BuildJourneyTimeSeries")
    Columns.DeltaHCVCol = _
        RequiredColumn(ws, "Delta HCV", "BuildJourneyTimeSeries")
    Columns.DeltaLTVCol = _
        RequiredColumn(ws, "Delta LTV", "BuildJourneyTimeSeries")
End Sub

Private Sub LoadCoreColumns( _
    ByVal ws As Worksheet, _
    ByRef Columns As JourneyColumns, _
    ByVal CallerName As String)

    Columns.NDGCol = RequiredColumn(ws, "NDG", CallerName)
    Columns.DateCol = RequiredColumn(ws, "Snapshot Date", CallerName)
    Columns.EventCol = RequiredColumn(ws, "Event", CallerName)
    Columns.ApprovedCol = RequiredColumn(ws, "Approved", CallerName)
    Columns.DrawnCol = RequiredColumn(ws, "Drawn", CallerName)
    Columns.MTMCol = RequiredColumn(ws, "MTM", CallerName)
    Columns.HCVCol = RequiredColumn(ws, "HCV", CallerName)
End Sub

Private Sub ShiftJourneyColumns(ByRef SourceColumns As JourneyColumns, _
                                ByRef TargetColumns As JourneyColumns, _
                                ByVal ColumnOffset As Long)
    TargetColumns.NDGCol = SourceColumns.NDGCol + ColumnOffset
    TargetColumns.DateCol = SourceColumns.DateCol + ColumnOffset
    TargetColumns.EventCol = SourceColumns.EventCol + ColumnOffset
    TargetColumns.ReasonCol = ShiftedColumn( _
        SourceColumns.ReasonCol, ColumnOffset)
    TargetColumns.ApprovedCol = SourceColumns.ApprovedCol + ColumnOffset
    TargetColumns.DrawnCol = SourceColumns.DrawnCol + ColumnOffset
    TargetColumns.MTMCol = SourceColumns.MTMCol + ColumnOffset
    TargetColumns.HCVCol = SourceColumns.HCVCol + ColumnOffset
    TargetColumns.LTVCol = ShiftedColumn( _
        SourceColumns.LTVCol, ColumnOffset)
    TargetColumns.MCCol = ShiftedColumn( _
        SourceColumns.MCCol, ColumnOffset)
    TargetColumns.SFCol = ShiftedColumn( _
        SourceColumns.SFCol, ColumnOffset)
    TargetColumns.MCClearedCol = ShiftedColumn( _
        SourceColumns.MCClearedCol, ColumnOffset)
    TargetColumns.DeltaApprovedCol = ShiftedColumn( _
        SourceColumns.DeltaApprovedCol, ColumnOffset)
    TargetColumns.DeltaDrawnCol = ShiftedColumn( _
        SourceColumns.DeltaDrawnCol, ColumnOffset)
    TargetColumns.DeltaHCVCol = ShiftedColumn( _
        SourceColumns.DeltaHCVCol, ColumnOffset)
    TargetColumns.DeltaLTVCol = ShiftedColumn( _
        SourceColumns.DeltaLTVCol, ColumnOffset)
End Sub

Private Function ShiftedColumn(ByVal SourceColumn As Long, _
                               ByVal ColumnOffset As Long) As Long
    If SourceColumn > 0 Then _
        ShiftedColumn = SourceColumn + ColumnOffset
End Function

Private Function DataRange( _
    ByVal ws As Worksheet, _
    ByVal ColumnNumber As Long, _
    ByVal LastRow As Long) As Range

    Set DataRange = ws.Range( _
        ws.Cells(FIRST_DATA_ROW, ColumnNumber), ws.Cells(LastRow, ColumnNumber))
End Function

Private Function ConstantValues( _
    ByVal Count As Long, _
    ByVal Value As Double) As Variant

    Dim Result() As Double
    Dim i As Long

    ReDim Result(1 To Count)
    For i = 1 To Count
        Result(i) = Value
    Next i
    ConstantValues = Result
End Function

Private Function DateKeyFor(ByVal Value As Variant) As String
    DateKeyFor = Format$(CDate(Value), "yyyymmdd")
End Function

Private Sub CopyRowValues( _
    ByVal SourceSheet As Worksheet, _
    ByVal SourceRow As Long, _
    ByVal SourceFirstColumn As Long, _
    ByVal TargetSheet As Worksheet, _
    ByVal TargetRow As Long, _
    ByVal TargetFirstColumn As Long, _
    ByVal ColumnCount As Long)

    TargetSheet.Cells(TargetRow, TargetFirstColumn) _
        .Resize(1, ColumnCount).Value = _
        SourceSheet.Cells(SourceRow, SourceFirstColumn) _
        .Resize(1, ColumnCount).Value
End Sub

Private Sub ClearColumns( _
    ByVal ws As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal ColumnNumbers As Variant)

    Dim ColumnNumber As Variant

    For Each ColumnNumber In ColumnNumbers
        If CLng(ColumnNumber) > 0 Then _
            ws.Cells(RowNumber, CLng(ColumnNumber)).ClearContents
    Next ColumnNumber
End Sub

Private Sub InitialiseDashboardTheme()
    If USE_DARK_THEME Then
        InitialiseDarkDashboardTheme
    Else
        InitialiseLightDashboardTheme
    End If
End Sub


Private Sub InitialiseLightDashboardTheme()
    ColorChartBackground = RGB(255, 255, 255)
    ColorPlotBackground = RGB(255, 255, 255)
    ColorChartText = RGB(60, 60, 60)
    ColorBorder = RGB(220, 220, 220)
    ColorGridline = RGB(225, 225, 225)
    ColorMTM = RGB(231, 171, 120)
    ColorHCV = RGB(243, 214, 112)
    ColorApproved = RGB(45, 45, 45)
    ColorApprovedArea = RGB(66, 133, 244)
    ColorDrawn = RGB(164, 139, 193)
    ColorLTV = RGB(66, 133, 244)
    ColorLTVArea = RGB(66, 133, 244)
    ColorHaircut = RGB(245, 166, 35)
    ColorDanger = RGB(220, 65, 65)
    ColorHCVOverlap = RGB(245, 130, 32)
    ColorMTMOverlap = RGB(200, 0, 0)
    ColorMCFill = RGB(255, 165, 80)
    ColorMCLine = RGB(255, 100, 0)
    ColorSFFill = RGB(255, 150, 150)
    ColorInactiveFill = RGB(190, 190, 190)
    ColorInactiveLine = RGB(135, 135, 135)
    ColorThreshold = RGB(220, 0, 0)
    ColorLabelBackground = RGB(255, 255, 255)
    ColorLabelText = ColorMCLine
    ColorTechnicalLabelText = RGB(255, 255, 255)
    ColorTechnicalLabelBorder = RGB(255, 255, 255)

    TransparencyGridline = 0.15
    TransparencyApprovedArea = 0.82
    TransparencyLTVArea = 0.82
    TransparencyHCVBand = 0.8
    TransparencyMTMBand = 0.5
    TransparencyHCVOverlap = 0.35
    TransparencyMTMOverlap = 0.2
    TransparencyRiskBand = 0.85
    TransparencyRiskBandAccent = 0.08
    TransparencyInactiveBand = 0.72
    TransparencyInactiveLine = 0.35
    TransparencyLabel = 0.15
End Sub

Private Sub InitialiseDarkDashboardTheme()
    ColorChartBackground = RGB(30, 30, 30)
    ColorPlotBackground = RGB(38, 38, 38)
    ColorChartText = RGB(235, 235, 235)
    ColorBorder = RGB(82, 82, 82)
    ColorGridline = RGB(110, 110, 110)
    ColorMTM = RGB(210, 85, 40)
'    ColorHCV = RGB(255, 212, 75)
    ColorHCV = RGB(210, 170, 40)
    ColorApproved = RGB(240, 50, 50)
    ColorApprovedArea = RGB(80, 145, 50)
    ColorDrawn = RGB(242, 242, 242)
    ColorLTV = RGB(80, 145, 50)
    ColorLTVArea = RGB(80, 145, 50)
    ColorHaircut = RGB(255, 179, 71)
    ColorDanger = RGB(192, 0, 0)
    ColorHCVOverlap = RGB(255, 145, 40)
    ColorMTMOverlap = RGB(255, 65, 65)
'    ColorDanger = RGB(255, 82, 82)
    ColorMCFill = RGB(255, 174, 90)
    ColorMCLine = RGB(255, 132, 40)
    ColorSFFill = RGB(255, 110, 120)
    ColorInactiveFill = RGB(130, 130, 130)
    ColorInactiveLine = RGB(180, 180, 180)
    ColorThreshold = RGB(255, 99, 99)
    ColorLabelBackground = ColorChartText
    ColorLabelText = ColorChartBackground
    ColorTechnicalLabelText = ColorChartText
    ColorTechnicalLabelBorder = RGB(255, 255, 255)

    TransparencyGridline = 0.35
    TransparencyApprovedArea = 0.82
    TransparencyLTVArea = 0.82
    TransparencyHCVBand = 0.8
    TransparencyMTMBand = 0.62
    TransparencyHCVOverlap = 0.3
    TransparencyMTMOverlap = 0.15
    TransparencyRiskBand = 0.85
    TransparencyRiskBandAccent = 0.85
    TransparencyInactiveBand = 0.65
    TransparencyInactiveLine = 0.25
    TransparencyLabel = 0
End Sub




