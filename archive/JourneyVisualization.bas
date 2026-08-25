Attribute VB_Name = "JourneyVisualization"
'Option Explicit
'
'Private Const SOURCE_SHEET As String = "NDG Journey"
'Private Const TIME_SERIES_SHEET As String = "NDG Journey TS"
'Private Const DASHBOARD_SHEET As String = "Journey Dashboard"
'Private Const FIRST_DATA_ROW As Long = 2
'Private Const NDG_COLUMN As Long = 2
'Private Const USE_DARK_THEME As Boolean = True
'
'Private Const LABEL_MIN_WIDTH As Double = 16
'Private Const LABEL_HEIGHT As Double = 11
'Private Const LABEL_CHARACTER_WIDTH As Double = 8
'Private Const LABEL_HORIZONTAL_MARGIN As Double = 0
'Private Const COLLATERAL_DATE_SCALE As Double = 100
'Private Const COLLATERAL_DATE_ORIGIN As Double = 1000
'Private Const RISK_BAND_SHAPE_OFFSET As Double = 0
'Private Const RISK_BAND_SHAPE_HEIGHT As Double = 7
'Private Const RISK_BAND_LABEL_GAP As Double = 9
'Private Const CHART_BOTTOM_MARGIN As Double = 5
'
'Private Type JourneyColumns
'    DateCol As Long
'    EventCol As Long
'    ReasonCol As Long
'    ApprovedCol As Long
'    DrawnCol As Long
'    MTMCol As Long
'    HCVCol As Long
'    LTVCol As Long
'    MCCol As Long
'    SFCol As Long
'    MCClearedCol As Long
'    DeltaApprovedCol As Long
'    DeltaDrawnCol As Long
'    DeltaHCVCol As Long
'    DeltaLTVCol As Long
'End Type
'
'Private ColorChartBackground As Long
'Private ColorPlotBackground As Long
'Private ColorChartText As Long
'Private ColorBorder As Long
'Private ColorGridline As Long
'Private ColorMTM As Long
'Private ColorHCV As Long
'Private ColorApproved As Long
'Private ColorApprovedArea As Long
'Private ColorDrawn As Long
'Private ColorLTV As Long
'Private ColorLTVArea As Long
'Private ColorHaircut As Long
'Private ColorDanger As Long
'Private ColorHCVOverlap As Long
'Private ColorMTMOverlap As Long
'Private ColorMCFill As Long
'Private ColorMCLine As Long
'Private ColorSFFill As Long
'Private ColorInactiveFill As Long
'Private ColorInactiveLine As Long
'Private ColorThreshold As Long
'Private ColorLabelBackground As Long
'Private ColorLabelText As Long
'Private ColorTechnicalLabelText As Long
'Private ColorTechnicalLabelBorder As Long
'
'Private TransparencyGridline As Double
'Private TransparencyApprovedArea As Double
'Private TransparencyLTVArea As Double
'Private TransparencyHCVBand As Double
'Private TransparencyMTMBand As Double
'Private TransparencyHCVOverlap As Double
'Private TransparencyMTMOverlap As Double
'Private TransparencyRiskBand As Double
'Private TransparencyRiskBandAccent As Double
'Private TransparencyInactiveBand As Double
'Private TransparencyInactiveLine As Double
'Private TransparencyLabel As Double
'
'Public Sub GenerateJourneyDashboard()
'    Dim PreviousCalculation As XlCalculation
'    Dim PreviousScreenUpdating As Boolean
'    Dim PreviousEnableEvents As Boolean
'    Dim ErrorNumber As Long
'    Dim ErrorDescription As String
'    Dim wsData As Worksheet
'    Dim wsDash As Worksheet
'    Dim Columns As JourneyColumns
'    Dim LastRow As Long
'    Dim ChartLastRow As Long
'    Dim MainChart As ChartObject
'    Dim LTVChart As ChartObject
'
'    On Error GoTo ExitHandler
'
'    PreviousScreenUpdating = Application.ScreenUpdating
'    PreviousCalculation = Application.Calculation
'    PreviousEnableEvents = Application.EnableEvents
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.EnableEvents = False
'
'    InitialiseDashboardTheme
'    BuildJourneyTimeSeries
'
'    Set wsData = Worksheets(TIME_SERIES_SHEET)
'    Set wsDash = CreateOrReplaceSheet(DASHBOARD_SHEET)
'    LoadDashboardColumns wsData, Columns
'
'    LastRow = GetLastRow(wsData, "B")
'    If LastRow < FIRST_DATA_ROW Then _
'        Err.Raise vbObjectError + 1000, "GenerateJourneyDashboard", _
'                  TIME_SERIES_SHEET & " contains no data."
'
'    ChartLastRow = LastRow
'    If IsLoanEndedEvent(SafeCellText(wsData.Cells(LastRow, Columns.EventCol))) Then _
'        ChartLastRow = LastRow - 1
'
'    If ChartLastRow < FIRST_DATA_ROW Then _
'        Err.Raise vbObjectError + 1001, "GenerateJourneyDashboard", _
'                  "No active Journey row is available for plotting."
'
'    Set MainChart = wsDash.ChartObjects.Add(20, 20, 800, 320)
'    Set LTVChart = wsDash.ChartObjects.Add(20, 335, 800, 320)
'
'    BuildMainChart MainChart, wsData, Columns, ChartLastRow
'    RebuildMainChartLegend MainChart
'    PositionTitleAndLegend MainChart
'
'    BuildLTVChart LTVChart, wsData, Columns, ChartLastRow
'    ReserveRiskBandShapeSpace MainChart
'    AlignChartPlotAreas MainChart, LTVChart
'
'    AddRiskBandAccentShapes MainChart, wsData, Columns, ChartLastRow, _
'                            AbovePlotArea:=False
'    AddRiskBandAccentShapes LTVChart, wsData, Columns, ChartLastRow, _
'                            AbovePlotArea:=True
'    AddMCDurationLabels MainChart, wsData, Columns, ChartLastRow
'
'ExitHandler:
'    ErrorNumber = Err.Number
'    ErrorDescription = Err.Description
'
'    Application.ScreenUpdating = PreviousScreenUpdating
'    Application.Calculation = PreviousCalculation
'    Application.EnableEvents = PreviousEnableEvents
'
'    If ErrorNumber <> 0 Then
'        MsgBox "Error " & ErrorNumber & vbCrLf & ErrorDescription, _
'               vbExclamation, "Journey Dashboard"
'    End If
'End Sub
'
'Public Sub BuildJourneyTimeSeries()
'    Dim wsSource As Worksheet
'    Dim wsTarget As Worksheet
'    Dim Columns As JourneyColumns
'    Dim RowsByDate As Object
'    Dim LastRow As Long
'    Dim LastCol As Long
'    Dim SourceRow As Long
'    Dim OutputRow As Long
'    Dim CurrentDate As Date
'    Dim EndDate As Date
'    Dim DateKey As String
'    Dim EventText As String
'    Dim LoanIsActive As Boolean
'    Dim HasSnapshot As Boolean
'    Dim r As Long
'
'    Set wsSource = Worksheets(SOURCE_SHEET)
'    Set wsTarget = CreateOrReplaceSheet(TIME_SERIES_SHEET)
'
'    LastRow = GetLastRow(wsSource, "B")
'    If LastRow < FIRST_DATA_ROW Then _
'        Err.Raise vbObjectError + 1100, "BuildJourneyTimeSeries", _
'                  SOURCE_SHEET & " contains no data."
'
'    LoadSourceColumns wsSource, Columns
'    LastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
'
'    wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(1, LastCol)).Value = _
'        wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, LastCol)).Value
'
'    Set RowsByDate = CreateObject("Scripting.Dictionary")
'    For r = FIRST_DATA_ROW To LastRow
'        If IsDate(wsSource.Cells(r, Columns.DateCol).Value) Then
'            RowsByDate(DateKeyFor(wsSource.Cells(r, Columns.DateCol).Value)) = r
'        End If
'    Next r
'
'    CurrentDate = CDate(wsSource.Cells(FIRST_DATA_ROW, Columns.DateCol).Value)
'    EndDate = CDate(wsSource.Cells(LastRow, Columns.DateCol).Value)
'    OutputRow = FIRST_DATA_ROW
'    LoanIsActive = True
'
'    Do While CurrentDate <= EndDate
'        DateKey = DateKeyFor(CurrentDate)
'        HasSnapshot = RowsByDate.Exists(DateKey)
'
'        If HasSnapshot Then
'            SourceRow = CLng(RowsByDate(DateKey))
'            CopyRowValues wsSource, SourceRow, wsTarget, OutputRow, LastCol
'
'            EventText = SafeCellText(wsSource.Cells(SourceRow, Columns.EventCol))
'            If IsLoanEndedEvent(EventText) Then
'                LoanIsActive = False
'            ElseIf IsLoanRestartedEvent(EventText) Then
'                LoanIsActive = True
'            End If
'        Else
'            CopyRowValues wsTarget, OutputRow - 1, wsTarget, OutputRow, LastCol
'            wsTarget.Cells(OutputRow, Columns.DateCol).Value = CurrentDate
'            wsTarget.Cells(OutputRow, Columns.EventCol).ClearContents
'
'            If Columns.MCClearedCol > 0 Then _
'                wsTarget.Cells(OutputRow, Columns.MCClearedCol).Value = 0
'        End If
'
'        If Not LoanIsActive Then
'            ClearColumns wsTarget, OutputRow, Array( _
'                Columns.ApprovedCol, _
'                Columns.DrawnCol, _
'                Columns.MTMCol, _
'                Columns.HCVCol, _
'                Columns.LTVCol)
'        End If
'
'        If Not HasSnapshot Or Not LoanIsActive Then
'            ClearColumns wsTarget, OutputRow, Array( _
'                Columns.DeltaApprovedCol, _
'                Columns.DeltaDrawnCol, _
'                Columns.DeltaHCVCol, _
'                Columns.DeltaLTVCol)
'        End If
'
'        OutputRow = OutputRow + 1
'        CurrentDate = CurrentDate + 1
'    Loop
'
'    wsTarget.Columns(Columns.DateCol).NumberFormat = "dd/mm/yyyy"
'End Sub
'
'Private Sub BuildMainChart( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long)
'
'    Dim ApprovedSeries As Series
'    Dim RiskBandHeight As Double
'
'    With ch.Chart
'        .ChartType = xlLine
'        .DisplayBlanksAs = xlNotPlotted
'        .HasTitle = True
'        .ChartTitle.Text = _
'            " Lombard Loan - NDG " & _
'            SafeCellText(wsData.Cells(FIRST_DATA_ROW, NDG_COLUMN))
'        .HasLegend = True
'        .Legend.Position = xlLegendPositionTop
'    End With
'
'    AddLineSeries ch, wsData, Columns.DateCol, Columns.MTMCol, LastRow, _
'                  "Collateral MV", ColorMTM
'    AddLineSeries ch, wsData, Columns.DateCol, Columns.HCVCol, LastRow, _
'                  "HCV", ColorHCV
''    AddLineSeries ch, wsData, Columns.DateCol, Columns.DrawnCol, LastRow, _
''                  "Drawn", ColorDrawn
'    Set ApprovedSeries = AddLineSeries( _
'        ch, wsData, Columns.DateCol, Columns.ApprovedCol, LastRow, _
'        "Approved / LTV", ColorLTV, msoLineDash)
'    AddAreaSeries ch, wsData, Columns.DateCol, Columns.ApprovedCol, _
'                  LastRow, "Approved Area", _
'                  ColorApprovedArea, TransparencyApprovedArea
'
'    FormatAmountAxis ch, wsData, Columns, LastRow
'    FormatDateAxis ch, wsData, Columns.DateCol, LastRow, False
'    FormatJourneyChart ch
'
'    AddCollateralBandSeries ch, wsData, Columns, LastRow
'    NormaliseCollateralBandGroup ch
'    RiskBandHeight = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
'    AddRiskBandSeries ch, wsData, Columns, LastRow, RiskBandHeight
'    MatchSecondaryAxisToPrimary ch
'
'    ' Excel can regroup combo-chart series when area bands are added.
'    FormatLineSeries ApprovedSeries, ColorLTV, msoLineDash
''    AddLTVLegendOnlySeries ch
'End Sub
'
'Public Sub FormatJourneyChart(ByVal ch As ChartObject)
'    FormatChartFrame ch
'    FormatPrimaryGridlines ch
'End Sub
'
'Private Sub BuildLTVChart( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long)
'
'    Dim LTVSeries As Series
'    Dim HCVRatioSeries As Series
'    Dim ThresholdSeries As Series
'    Dim RiskBandHeight As Double
'
'    With ch.Chart
'        .ChartType = xlLine
'        .DisplayBlanksAs = xlNotPlotted
'        .HasTitle = False
'        .HasLegend = False
'    End With
'
'    Set LTVSeries = AddLineSeries( _
'        ch, wsData, Columns.DateCol, Columns.LTVCol, LastRow, _
'        "LTV", ColorLTV)
'    AddAreaSeries ch, wsData, Columns.DateCol, Columns.LTVCol, _
'                  LastRow, "LTV Area", ColorLTVArea, TransparencyLTVArea
'    Set HCVRatioSeries = AddRatioLineSeries( _
'        ch, wsData, Columns.DateCol, Columns.HCVCol, _
'        Columns.MTMCol, LastRow, "HCV / MTM", ColorHCV)
'
'    FormatDateAxis ch, wsData, Columns.DateCol, LastRow, True
'    FormatLTVPrimaryAxis ch
'    RiskBandHeight = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
'    AddRiskBandSeries ch, wsData, Columns, LastRow, RiskBandHeight
'
'    Set ThresholdSeries = ch.Chart.SeriesCollection.NewSeries
'    With ThresholdSeries
'        .name = "100% Threshold"
'        .XValues = DataRange(wsData, Columns.DateCol, LastRow)
'        .Values = ConstantValues(LastRow - FIRST_DATA_ROW + 1, 1)
'    End With
'    FormatLineSeries ThresholdSeries, ColorMTM
'
'    ' Excel can regroup combo-chart series when area bands are added.
'    FormatLineSeries LTVSeries, ColorLTV, msoLineDash
'    FormatLineSeries HCVRatioSeries, ColorHCV
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        .MaximumScaleIsAuto = False
'        .MaximumScale = RiskBandHeight
'    End With
'    FormatChartFrame ch
'End Sub
'
'Private Sub AddAreaSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal ValueCol As Long, _
'    ByVal LastRow As Long, _
'    ByVal SeriesName As String, _
'    ByVal FillColour As Long, _
'    ByVal FillTransparency As Double)
'
'    Dim AreaSeries As Series
'
'    Set AreaSeries = ch.Chart.SeriesCollection.NewSeries
'    With AreaSeries
'        .name = SeriesName
'        .XValues = DataRange(wsData, DateCol, LastRow)
'        .Values = DataRange(wsData, ValueCol, LastRow)
'        .ChartType = xlArea
'        .AxisGroup = xlPrimary
'        .Format.Line.Visible = msoFalse
'
'        With .Format.Fill
'            .Visible = msoTrue
'            .Solid
'            .ForeColor.RGB = FillColour
'            .Transparency = FillTransparency
'        End With
'    End With
'End Sub
'
'Private Function AddLineSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal ValueCol As Long, _
'    ByVal LastRow As Long, _
'    ByVal SeriesName As String, _
'    ByVal LineColour As Long, _
'    Optional ByVal DashStyle As MsoLineDashStyle = msoLineSolid) As Series
'
'    Set AddLineSeries = ch.Chart.SeriesCollection.NewSeries
'    With AddLineSeries
'        .name = SeriesName
'        .XValues = DataRange(wsData, DateCol, LastRow)
'        .Values = DataRange(wsData, ValueCol, LastRow)
'    End With
'    FormatLineSeries AddLineSeries, LineColour, DashStyle
'End Function
'
'Private Function AddRatioLineSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal NumeratorCol As Long, _
'    ByVal DenominatorCol As Long, _
'    ByVal LastRow As Long, _
'    ByVal SeriesName As String, _
'    ByVal LineColour As Long) As Series
'
'    Dim RatioValues() As Variant
'    Dim PointIndex As Long
'    Dim SourceRow As Long
'    Dim NumeratorValue As Variant
'    Dim DenominatorValue As Variant
'
'    ReDim RatioValues(1 To LastRow - FIRST_DATA_ROW + 1)
'
'    For SourceRow = FIRST_DATA_ROW To LastRow
'        PointIndex = SourceRow - FIRST_DATA_ROW + 1
'        NumeratorValue = wsData.Cells(SourceRow, NumeratorCol).Value2
'        DenominatorValue = wsData.Cells(SourceRow, DenominatorCol).Value2
'        RatioValues(PointIndex) = CVErr(xlErrNA)
'
'        If Not IsError(NumeratorValue) And _
'           Not IsError(DenominatorValue) Then
'            If IsNumeric(NumeratorValue) And _
'               IsNumeric(DenominatorValue) Then
'                If CDbl(DenominatorValue) <> 0 Then
'                    RatioValues(PointIndex) = _
'                        CDbl(NumeratorValue) / CDbl(DenominatorValue)
'                End If
'            End If
'        End If
'    Next SourceRow
'
'    Set AddRatioLineSeries = ch.Chart.SeriesCollection.NewSeries
'    With AddRatioLineSeries
'        .name = SeriesName
'        .XValues = DataRange(wsData, DateCol, LastRow)
'        .Values = RatioValues
'    End With
'    FormatLineSeries AddRatioLineSeries, LineColour
'End Function
'
'Private Sub FormatLineSeries( _
'    ByVal DataSeries As Series, _
'    ByVal LineColour As Long, _
'    Optional ByVal DashStyle As MsoLineDashStyle = msoLineSolid, _
'    Optional ByVal LineWeight As Double = 1.25)
'
'    With DataSeries
'        .ChartType = xlLine
'        .AxisGroup = xlPrimary
'        .MarkerStyle = xlMarkerStyleNone
'        With .Format.Line
'            .Visible = msoTrue
'            .Weight = LineWeight
'            .ForeColor.RGB = LineColour
'            .DashStyle = DashStyle
'        End With
'    End With
'End Sub
'
'Private Sub FormatAmountAxis( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long)
'
'    Const APPROVED_MIN_POSITION As Double = 0.3
'
'    Dim DataMinimum As Double
'    Dim ApprovedMinimum As Double
'    Dim AxisMinimum As Double
'    Dim AxisMaximum As Double
'    Dim TargetMinimum As Double
'
'    DataMinimum = WorksheetFunction.Min( _
'        DataRange(wsData, Columns.MTMCol, LastRow), _
'        DataRange(wsData, Columns.HCVCol, LastRow), _
'        DataRange(wsData, Columns.ApprovedCol, LastRow), _
'        DataRange(wsData, Columns.DrawnCol, LastRow))
'
'    ApprovedMinimum = WorksheetFunction.Min( _
'        DataRange(wsData, Columns.ApprovedCol, LastRow))
'
'    If DataMinimum >= 0 Then
'        AxisMinimum = DataMinimum * 0.95
'    Else
'        AxisMinimum = DataMinimum * 1.05
'    End If
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        .MinimumScaleIsAuto = False
'        .MinimumScale = AxisMinimum
'        .MaximumScaleIsAuto = True
'        .TickLabels.NumberFormat = "#,##0,"
'        .HasTitle = True
'        .AxisTitle.Text = "Amount (thousand Euros)"
'    End With
'
'    ch.Activate
'    DoEvents
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        AxisMaximum = .MaximumScale
'        .MaximumScaleIsAuto = False
'        .MaximumScale = AxisMaximum
'    End With
'
'    If AxisMaximum > ApprovedMinimum Then
'        TargetMinimum = _
'            (ApprovedMinimum - APPROVED_MIN_POSITION * AxisMaximum) / _
'            (1 - APPROVED_MIN_POSITION)
'
'        ' Non-negative data should not force the axis below zero.
'        If DataMinimum >= 0 Then _
'            TargetMinimum = Application.Max(0, TargetMinimum)
'
'        If TargetMinimum < AxisMinimum Then
'            ch.Chart.Axes(xlValue, xlPrimary).MinimumScale = TargetMinimum
'        End If
'    End If
'End Sub
'
'Private Sub FormatLTVPrimaryAxis(ByVal ch As ChartObject)
'    Const MINIMUM_MAXIMUM As Double = 1
'    Const MAXIMUM_MAXIMUM As Double = 2
'    Dim AutoMaximum As Double
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        .MinimumScaleIsAuto = False
'        .MinimumScale = 0
'        .MaximumScaleIsAuto = True
'        .MajorUnitIsAuto = True
'        .TickLabels.NumberFormat = "0%"
'    End With
'
'    ch.Activate
'    DoEvents
'    AutoMaximum = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        .HasTitle = True
'        .AxisTitle.Text = "LTV (% of Collateral MV)"
'        .MaximumScaleIsAuto = False
'        .MaximumScale = Application.Max( _
'            MINIMUM_MAXIMUM, Application.Min(MAXIMUM_MAXIMUM, AutoMaximum))
'        On Error Resume Next
'        .MajorGridlines.Format.Line.Visible = msoFalse
'        On Error GoTo 0
'    End With
'End Sub
'
'Private Sub FormatDateAxis( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal LastRow As Long, _
'    ByVal ShowLabels As Boolean)
'
'    With ch.Chart.Axes(xlCategory, xlPrimary)
'        .CategoryType = xlTimeScale
'        .AxisBetweenCategories = True
'        .BaseUnit = xlDays
'        .MinimumScaleIsAuto = False
'        .MinimumScale = CDbl(CDate(wsData.Cells(FIRST_DATA_ROW, DateCol).Value))
'        .MaximumScaleIsAuto = False
'        .MaximumScale = CDbl(CDate(wsData.Cells(LastRow, DateCol).Value))
'        .TickLabels.NumberFormat = "dd-mmm-yyyy"
'        If ShowLabels Then
'            .TickLabelPosition = xlTickLabelPositionNextToAxis
'        Else
'            .TickLabelPosition = xlTickLabelPositionNone
'        End If
'    End With
'End Sub
'
'Private Sub FormatChartFrame(ByVal ch As ChartObject)
'    With ch.Chart.ChartArea.Format
'        .Fill.Visible = msoTrue
'        .Fill.Solid
'        .Fill.ForeColor.RGB = ColorChartBackground
'        .Line.Visible = msoFalse
''        .Line.ForeColor.RGB = ColorBorder
'    End With
'
'    With ch.Chart.PlotArea.Format
'        .Fill.Visible = msoTrue
'        .Fill.Solid
'        .Fill.ForeColor.RGB = ColorPlotBackground
'        .Line.Visible = msoFalse
''        .Line.ForeColor.RGB = ColorBorder
'    End With
'
'    FormatChartText ch
'End Sub
'
'Private Sub FormatChartText(ByVal ch As ChartObject)
'    On Error Resume Next
'
'    With ch.Chart
'        If .HasTitle Then .ChartTitle.Font.Color = ColorChartText
'        If .HasLegend Then .Legend.Font.Color = ColorChartText
'
'        With .Axes(xlValue, xlPrimary)
'            .TickLabels.Font.Color = ColorChartText
'            .Format.Line.ForeColor.RGB = ColorBorder
'            If .HasTitle Then .AxisTitle.Font.Color = ColorChartText
'        End With
'
'        With .Axes(xlCategory, xlPrimary)
'            .TickLabels.Font.Color = ColorChartText
'            .Format.Line.ForeColor.RGB = ColorBorder
'            If .HasTitle Then .AxisTitle.Font.Color = ColorChartText
'        End With
'    End With
'
'    On Error GoTo 0
'End Sub
'
'Private Sub FormatPrimaryGridlines(ByVal ch As ChartObject)
'    On Error Resume Next
'    With ch.Chart.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line
'        .Visible = msoTrue
'        .ForeColor.RGB = ColorGridline
'        .Weight = 0.4
'        .Transparency = TransparencyGridline
'        .DashStyle = msoLineDash
'    End With
'    On Error GoTo 0
'End Sub
'
'Private Sub AddCollateralBandSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long)
'
'    Dim BaseValues() As Variant
'    Dim HCVOverlapValues() As Variant
'    Dim HaircutRestValues() As Variant
'    Dim MTMOverlapValues() As Variant
'    Dim DangerRestValues() As Variant
'    Dim Dates() As Variant
'    Dim AxisMaximum As Double
'    Dim OriginalPointCount As Long
'    Dim MaximumPointCount As Long
'    Dim OutputPointIndex As Long
'    Dim SourceRow As Long
'    Dim DateValue As Double
'    Dim ApprovedValue As Double
'    Dim MTMValue As Double
'    Dim HCVValue As Double
'    Dim NextDateValue As Double
'    Dim NextApprovedValue As Double
'    Dim NextMTMValue As Double
'    Dim NextHCVValue As Double
'    Dim HasApproved As Boolean
'    Dim HasCollateral As Boolean
'    Dim CrossingFractions(1 To 2) As Double
'    Dim CrossingCount As Long
'    Dim CrossingIndex As Long
'    Dim CrossingValue As Double
'
'    ch.Activate
'    DoEvents
'    AxisMaximum = ch.Chart.Axes(xlValue, xlPrimary).MaximumScale
'
'    OriginalPointCount = LastRow - FIRST_DATA_ROW + 1
'    If OriginalPointCount < 1 Then Exit Sub
'    MaximumPointCount = OriginalPointCount * 3
'
'    ReDim Dates(1 To MaximumPointCount)
'    ReDim BaseValues(1 To MaximumPointCount)
'    ReDim HCVOverlapValues(1 To MaximumPointCount)
'    ReDim HaircutRestValues(1 To MaximumPointCount)
'    ReDim MTMOverlapValues(1 To MaximumPointCount)
'    ReDim DangerRestValues(1 To MaximumPointCount)
'
'    For SourceRow = FIRST_DATA_ROW To LastRow
'        DateValue = CDbl(CDate( _
'            wsData.Cells(SourceRow, Columns.DateCol).Value))
'        HasApproved = IsNumeric( _
'            wsData.Cells(SourceRow, Columns.ApprovedCol).Value2)
'        HasCollateral = _
'            IsNumeric(wsData.Cells(SourceRow, Columns.MTMCol).Value2) _
'            And IsNumeric(wsData.Cells(SourceRow, Columns.HCVCol).Value2)
'
'        ApprovedValue = 0
'        MTMValue = 0
'        HCVValue = 0
'        If HasApproved Then _
'            ApprovedValue = CDbl( _
'                wsData.Cells(SourceRow, Columns.ApprovedCol).Value2)
'        If HasCollateral Then
'            MTMValue = CDbl(wsData.Cells(SourceRow, Columns.MTMCol).Value2)
'            HCVValue = CDbl(wsData.Cells(SourceRow, Columns.HCVCol).Value2)
'        End If
'
'        AppendCollateralBandPoint Dates, BaseValues, HCVOverlapValues, _
'            HaircutRestValues, MTMOverlapValues, DangerRestValues, _
'            OutputPointIndex, DateValue, HasApproved, ApprovedValue, _
'            HasCollateral, MTMValue, HCVValue, AxisMaximum
'
'        If SourceRow < LastRow And HasApproved And HasCollateral _
'           And IsDate(wsData.Cells(SourceRow + 1, Columns.DateCol).Value) _
'           And IsNumeric( _
'               wsData.Cells(SourceRow + 1, Columns.ApprovedCol).Value2) _
'           And IsNumeric( _
'               wsData.Cells(SourceRow + 1, Columns.MTMCol).Value2) _
'           And IsNumeric( _
'               wsData.Cells(SourceRow + 1, Columns.HCVCol).Value2) Then
'
'            NextDateValue = CDbl(CDate( _
'                wsData.Cells(SourceRow + 1, Columns.DateCol).Value))
'            NextApprovedValue = CDbl( _
'                wsData.Cells(SourceRow + 1, Columns.ApprovedCol).Value2)
'            NextMTMValue = CDbl( _
'                wsData.Cells(SourceRow + 1, Columns.MTMCol).Value2)
'            NextHCVValue = CDbl( _
'                wsData.Cells(SourceRow + 1, Columns.HCVCol).Value2)
'
'            CrossingCount = 0
'            CrossingValue = LinearCrossingFraction( _
'                ApprovedValue - HCVValue, _
'                NextApprovedValue - NextHCVValue)
'            If CrossingValue > 0 And CrossingValue < 1 Then
'                CrossingCount = 1
'                CrossingFractions(1) = CrossingValue
'            End If
'
'            CrossingValue = LinearCrossingFraction( _
'                ApprovedValue - MTMValue, _
'                NextApprovedValue - NextMTMValue)
'            If CrossingValue > 0 And CrossingValue < 1 Then
'                If CrossingCount = 0 Then
'                    CrossingCount = 1
'                    CrossingFractions(1) = CrossingValue
'                ElseIf Abs(CrossingValue - CrossingFractions(1)) > 0.0000001 Then
'                    CrossingCount = 2
'                    CrossingFractions(2) = CrossingValue
'                    If CrossingFractions(2) < CrossingFractions(1) Then
'                        CrossingValue = CrossingFractions(1)
'                        CrossingFractions(1) = CrossingFractions(2)
'                        CrossingFractions(2) = CrossingValue
'                    End If
'                End If
'            End If
'
'            For CrossingIndex = 1 To CrossingCount
'                CrossingValue = CrossingFractions(CrossingIndex)
'                AppendCollateralBandPoint Dates, BaseValues, _
'                    HCVOverlapValues, HaircutRestValues, _
'                    MTMOverlapValues, DangerRestValues, OutputPointIndex, _
'                    DateValue + CrossingValue * (NextDateValue - DateValue), _
'                    True, ApprovedValue + CrossingValue * _
'                        (NextApprovedValue - ApprovedValue), _
'                    True, MTMValue + CrossingValue * _
'                        (NextMTMValue - MTMValue), _
'                    HCVValue + CrossingValue * _
'                        (NextHCVValue - HCVValue), AxisMaximum
'            Next CrossingIndex
'        End If
'    Next SourceRow
'
'    ReDim Preserve Dates(1 To OutputPointIndex)
'    ReDim Preserve BaseValues(1 To OutputPointIndex)
'    ReDim Preserve HCVOverlapValues(1 To OutputPointIndex)
'    ReDim Preserve HaircutRestValues(1 To OutputPointIndex)
'    ReDim Preserve MTMOverlapValues(1 To OutputPointIndex)
'    ReDim Preserve DangerRestValues(1 To OutputPointIndex)
'
'    ScaleCollateralDates Dates, _
'        ch.Chart.Axes(xlCategory, xlPrimary).MinimumScale
'
'    AddStackedAreaSeries ch, Dates, _
'                         BaseValues, "Collateral Base", _
'                         ColorPlotBackground, 1, False
'    AddStackedAreaSeries ch, Dates, _
'                         HCVOverlapValues, "Approved in HCV Band", _
'                         ColorHCVOverlap, TransparencyHCVOverlap, True
'    AddStackedAreaSeries ch, Dates, _
'                         HaircutRestValues, "Haircut", ColorHaircut, _
'                         TransparencyHCVBand, True
'    AddStackedAreaSeries ch, Dates, _
'                         MTMOverlapValues, "Approved Above MTM", _
'                         ColorMTMOverlap, TransparencyMTMOverlap, True
'    AddStackedAreaSeries ch, Dates, _
'                         DangerRestValues, "Above MTM Risk", ColorDanger, _
'                         TransparencyMTMBand, True
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        .MaximumScaleIsAuto = False
'        .MaximumScale = AxisMaximum
'    End With
'End Sub
'
'Private Sub ScaleCollateralDates( _
'    ByRef Dates() As Variant, _
'    ByVal MinimumDate As Double)
'
'    Dim PointIndex As Long
'
'    For PointIndex = LBound(Dates) To UBound(Dates)
'        Dates(PointIndex) = COLLATERAL_DATE_ORIGIN + _
'            WorksheetFunction.Round( _
'                (CDbl(Dates(PointIndex)) - MinimumDate) * _
'                COLLATERAL_DATE_SCALE, 0)
'    Next PointIndex
'End Sub
'
'Private Sub AppendCollateralBandPoint( _
'    ByRef Dates() As Variant, _
'    ByRef BaseValues() As Variant, _
'    ByRef HCVOverlapValues() As Variant, _
'    ByRef HaircutRestValues() As Variant, _
'    ByRef MTMOverlapValues() As Variant, _
'    ByRef DangerRestValues() As Variant, _
'    ByRef PointIndex As Long, _
'    ByVal DateValue As Double, _
'    ByVal HasApproved As Boolean, _
'    ByVal ApprovedValue As Double, _
'    ByVal HasCollateral As Boolean, _
'    ByVal MTMValue As Double, _
'    ByVal HCVValue As Double, _
'    ByVal AxisMaximum As Double)
'
'    Dim HaircutValue As Double
'    Dim DangerValue As Double
'
'    PointIndex = PointIndex + 1
'    Dates(PointIndex) = DateValue
'
'    If Not HasCollateral Then
'        BaseValues(PointIndex) = CVErr(xlErrNA)
'        HCVOverlapValues(PointIndex) = CVErr(xlErrNA)
'        HaircutRestValues(PointIndex) = CVErr(xlErrNA)
'        MTMOverlapValues(PointIndex) = CVErr(xlErrNA)
'        DangerRestValues(PointIndex) = CVErr(xlErrNA)
'        Exit Sub
'    End If
'
'    HaircutValue = Abs(MTMValue - HCVValue)
'    DangerValue = Application.Max(0, AxisMaximum - MTMValue)
'    BaseValues(PointIndex) = HCVValue
'
'    If HasApproved Then
'        HCVOverlapValues(PointIndex) = Application.Max(0, _
'            Application.Min(HaircutValue, _
'                Application.Min(ApprovedValue, MTMValue) - HCVValue))
'        MTMOverlapValues(PointIndex) = Application.Max(0, _
'            Application.Min(DangerValue, _
'                Application.Min(ApprovedValue, AxisMaximum) - MTMValue))
'    Else
'        HCVOverlapValues(PointIndex) = 0
'        MTMOverlapValues(PointIndex) = 0
'    End If
'
'    HaircutRestValues(PointIndex) = _
'        HaircutValue - CDbl(HCVOverlapValues(PointIndex))
'    DangerRestValues(PointIndex) = _
'        DangerValue - CDbl(MTMOverlapValues(PointIndex))
'End Sub
'
'Private Function LinearCrossingFraction( _
'    ByVal StartDifference As Double, _
'    ByVal EndDifference As Double) As Double
'
'    LinearCrossingFraction = -1
'    If (StartDifference < 0 And EndDifference > 0) _
'       Or (StartDifference > 0 And EndDifference < 0) Then _
'        LinearCrossingFraction = _
'            -StartDifference / (EndDifference - StartDifference)
'End Function
'
'Private Sub AddStackedAreaSeries( _
'    ByVal ch As ChartObject, _
'    ByVal XValues As Variant, _
'    ByVal YValues As Variant, _
'    ByVal SeriesName As String, _
'    ByVal FillColour As Long, _
'    ByVal FillTransparency As Double, _
'    ByVal ShowFill As Boolean)
'
'    Dim DataSeries As Series
'
'    Set DataSeries = ch.Chart.SeriesCollection.NewSeries
'    With DataSeries
'        .name = SeriesName
'        .XValues = XValues
'        .Values = YValues
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .Format.Line.Visible = msoFalse
'        With .Format.Fill
'            .Visible = IIf(ShowFill, msoTrue, msoFalse)
'            If ShowFill Then
'                .Solid
'                .ForeColor.RGB = FillColour
'                .Transparency = FillTransparency
'            End If
'        End With
'    End With
'End Sub
'
'Private Sub NormaliseCollateralBandGroup(ByVal ch As ChartObject)
'    With ch.Chart.SeriesCollection("Collateral Base")
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .PlotOrder = 1
'    End With
'    With ch.Chart.SeriesCollection("Approved in HCV Band")
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .PlotOrder = 2
'    End With
'    With ch.Chart.SeriesCollection("Haircut")
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .PlotOrder = 3
'    End With
'    With ch.Chart.SeriesCollection("Approved Above MTM")
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .PlotOrder = 4
'    End With
'    With ch.Chart.SeriesCollection("Above MTM Risk")
'        .ChartType = xlAreaStacked
'        .AxisGroup = xlSecondary
'        .PlotOrder = 5
'    End With
'    ch.Activate
'    DoEvents
'End Sub
'
'Private Sub MatchSecondaryAxisToPrimary(ByVal ch As ChartObject)
'    Dim MinimumValue As Double
'    Dim MaximumValue As Double
'    Dim MinimumDate As Double
'    Dim MaximumDate As Double
'    Dim AxisBetweenCategories As Boolean
'    Dim DatePadding As Double
'
'    With ch.Chart.Axes(xlValue, xlPrimary)
'        MinimumValue = .MinimumScale
'        MaximumValue = .MaximumScale
'    End With
'
'    With ch.Chart.Axes(xlCategory, xlPrimary)
'        MinimumDate = .MinimumScale
'        MaximumDate = .MaximumScale
'        AxisBetweenCategories = .AxisBetweenCategories
'    End With
'
'    If AxisBetweenCategories Then DatePadding = 0.5
'
'    ch.Chart.HasAxis(xlValue, xlSecondary) = True
'    With ch.Chart.Axes(xlValue, xlSecondary)
'        .MinimumScaleIsAuto = False
'        .MinimumScale = MinimumValue
'        .MaximumScaleIsAuto = False
'        .MaximumScale = MaximumValue
'        .HasTitle = False
'        .TickLabelPosition = xlTickLabelPositionNone
'        .Format.Line.Visible = msoFalse
'        On Error Resume Next
'        .MajorGridlines.Format.Line.Visible = msoFalse
'        On Error GoTo 0
'    End With
'
'    ch.Chart.HasAxis(xlCategory, xlSecondary) = True
'    With ch.Chart.Axes(xlCategory, xlSecondary)
'        .CategoryType = xlTimeScale
'        .BaseUnit = xlDays
'        .AxisBetweenCategories = False
'        .MinimumScaleIsAuto = False
'        .MinimumScale = COLLATERAL_DATE_ORIGIN - _
'                        DatePadding * COLLATERAL_DATE_SCALE
'        .MaximumScaleIsAuto = False
'        .MaximumScale = COLLATERAL_DATE_ORIGIN + _
'                        (MaximumDate - MinimumDate + DatePadding) * _
'                        COLLATERAL_DATE_SCALE
'        .HasTitle = False
'        .TickLabelPosition = xlTickLabelPositionNone
'        .MajorTickMark = xlTickMarkNone
'        .MinorTickMark = xlTickMarkNone
'        .Format.Line.Visible = msoFalse
'        On Error Resume Next
'        .MajorGridlines.Format.Line.Visible = msoFalse
'        On Error GoTo 0
'    End With
'End Sub
'
'Private Sub AddRiskBandSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long, _
'    ByVal BandHeight As Double)
'
'    AddInactiveLoanBandSeries ch, wsData, Columns, LastRow, BandHeight
'    AddStatusBandSeries ch, wsData, Columns, LastRow, _
'                        Columns.MCCol, Columns.MCClearedCol, "MC", _
'                        ColorMCFill, BandHeight
'    AddStatusBandSeries ch, wsData, Columns, LastRow, _
'                        Columns.SFCol, 0, "SF", _
'                        ColorSFFill, BandHeight
'End Sub
'
'Private Sub AddRiskBandAccentShapes( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long, _
'    ByVal AbovePlotArea As Boolean)
'
'    Dim Episodes As Collection
'
'    ch.Activate
'    DoEvents
'
'    Set Episodes = GetInactiveLoanEpisodes(wsData, Columns, LastRow)
'    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
'                           ColorInactiveLine, AbovePlotArea
'
'    Set Episodes = GetStatusEpisodes( _
'        wsData, Columns.MCCol, Columns.MCClearedCol, LastRow)
'    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
'                           ColorMCLine, AbovePlotArea
'
'    Set Episodes = GetStatusEpisodes( _
'        wsData, Columns.SFCol, 0, LastRow)
'    AddEpisodeAccentShapes ch, wsData, Columns.DateCol, Episodes, _
'                           ColorSFFill, AbovePlotArea
'End Sub
'
'Private Sub AddEpisodeAccentShapes( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal Episodes As Collection, _
'    ByVal FillColour As Long, _
'    ByVal AbovePlotArea As Boolean)
'
'    Dim Episode As Variant
'
'    For Each Episode In Episodes
'        AddRiskBandAccentShape ch, wsData, DateCol, _
'            CLng(Episode(0)), CLng(Episode(1)), FillColour, _
'            AbovePlotArea
'    Next Episode
'End Sub
'
'Private Sub AddRiskBandAccentShape( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal StartRow As Long, _
'    ByVal EndRow As Long, _
'    ByVal FillColour As Long, _
'    ByVal AbovePlotArea As Boolean)
'
'    Dim AxisMinimum As Double
'    Dim AxisMaximum As Double
'    Dim AxisPadding As Double
'    Dim VisibleMinimum As Double
'    Dim VisibleMaximum As Double
'    Dim StartDate As Double
'    Dim EndDate As Double
'    Dim ShapeLeft As Double
'    Dim ShapeRight As Double
'    Dim ShapeTop As Double
'    Dim ShapeHeight As Double
'    Dim shp As Shape
'
'    With ch.Chart.Axes(xlCategory, xlPrimary)
'        AxisMinimum = .MinimumScale
'        AxisMaximum = .MaximumScale
'        If .AxisBetweenCategories Then AxisPadding = 0.5
'    End With
'
'    VisibleMinimum = AxisMinimum - AxisPadding
'    VisibleMaximum = AxisMaximum + AxisPadding
'    If VisibleMaximum <= VisibleMinimum Then Exit Sub
'
'    StartDate = CDbl(CDate(wsData.Cells(StartRow, DateCol).Value))
'    EndDate = CDbl(CDate(wsData.Cells(EndRow, DateCol).Value))
'
'    With ch.Chart.PlotArea
'        ShapeLeft = .InsideLeft + _
'            (StartDate - VisibleMinimum) / _
'            (VisibleMaximum - VisibleMinimum) * .InsideWidth
'        ShapeRight = .InsideLeft + _
'            (EndDate - VisibleMinimum) / _
'            (VisibleMaximum - VisibleMinimum) * .InsideWidth
'        If AbovePlotArea Then
'            ShapeTop = .InsideTop - RISK_BAND_SHAPE_OFFSET - _
'                       RISK_BAND_SHAPE_HEIGHT
'        Else
'            ShapeTop = .InsideTop + .InsideHeight + _
'                       RISK_BAND_SHAPE_OFFSET
'        End If
'    End With
'
'    ShapeLeft = Application.Max( _
'        ch.Chart.PlotArea.InsideLeft, ShapeLeft)
'    ShapeRight = Application.Min( _
'        ch.Chart.PlotArea.InsideLeft + ch.Chart.PlotArea.InsideWidth, _
'        ShapeRight)
'    If ShapeRight <= ShapeLeft Then ShapeRight = ShapeLeft + 1
'
'    ShapeHeight = RISK_BAND_SHAPE_HEIGHT
'    If ShapeTop < 1 Then _
'        Err.Raise vbObjectError + 1300, "AddRiskBandAccentShape", _
'                  "No chart space is available above the plot area."
'    If ShapeTop + ShapeHeight > ch.Chart.ChartArea.Height - 1 Then _
'        Err.Raise vbObjectError + 1300, "AddRiskBandAccentShape", _
'                  "No chart space is available below the category axis."
'
'    Set shp = ch.Chart.Shapes.AddShape( _
'        msoShapeRectangle, ShapeLeft, ShapeTop, _
'        ShapeRight - ShapeLeft, ShapeHeight)
'
'    With shp
'        .name = "RiskBand_Accent_" & _
'                Format$(ch.Chart.Shapes.Count, "000")
'        With .Fill
'            .Visible = msoTrue
'            .Solid
'            .ForeColor.RGB = FillColour
'            .Transparency = TransparencyRiskBandAccent
'        End With
'        .Line.Visible = msoFalse
'        .ZOrder msoBringToFront
'    End With
'End Sub
'
'Private Sub AddStatusBandSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long, _
'    ByVal ActiveCol As Long, _
'    ByVal CloseCol As Long, _
'    ByVal LabelPrefix As String, _
'    ByVal FillColour As Long, _
'    ByVal BandHeight As Double)
'
'    Dim Episodes As Collection
'    Dim Episode As Variant
'
'    Set Episodes = GetStatusEpisodes(wsData, ActiveCol, CloseCol, LastRow)
'
'    For Each Episode In Episodes
'        AddSingleBandEpisodeSeries ch, wsData, Columns.DateCol, LastRow, _
'            CLng(Episode(0)), CLng(Episode(1)), BandHeight, LabelPrefix, _
'            FillColour, FillColour, TransparencyRiskBand, 1
'    Next Episode
'End Sub
'
'Private Function GetStatusEpisodes( _
'    ByVal wsData As Worksheet, _
'    ByVal ActiveCol As Long, _
'    ByVal CloseCol As Long, _
'    ByVal LastRow As Long) As Collection
'
'    Dim Result As New Collection
'    Dim StartRow As Long
'    Dim InBand As Boolean
'    Dim r As Long
'
'    For r = FIRST_DATA_ROW To LastRow
'        If Not InBand Then
'            If GetNumericValue(wsData.Cells(r, ActiveCol).Value2) > 0 Then
'                StartRow = r
'                InBand = True
'            End If
'        End If
'
'        If InBand And StatusEpisodeClosed(wsData, r, ActiveCol, CloseCol) Then
'            Result.Add Array(StartRow, r, True)
'            InBand = False
'        End If
'    Next r
'
'    If InBand Then Result.Add Array(StartRow, LastRow, False)
'    Set GetStatusEpisodes = Result
'End Function
'
'Private Function StatusEpisodeClosed( _
'    ByVal wsData As Worksheet, _
'    ByVal RowNumber As Long, _
'    ByVal ActiveCol As Long, _
'    ByVal CloseCol As Long) As Boolean
'
'    If CloseCol > 0 Then
'        StatusEpisodeClosed = _
'            (GetNumericValue(wsData.Cells(RowNumber, CloseCol).Value2) = 1)
'    Else
'        StatusEpisodeClosed = _
'            (GetNumericValue(wsData.Cells(RowNumber, ActiveCol).Value2) = 0)
'    End If
'End Function
'
'Private Function GetInactiveLoanEpisodes( _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long) As Collection
'
'    Dim Result As New Collection
'    Dim StartRow As Long
'    Dim InInactivePeriod As Boolean
'    Dim EventText As String
'    Dim r As Long
'
'    For r = FIRST_DATA_ROW To LastRow
'        EventText = SafeCellText(wsData.Cells(r, Columns.EventCol))
'
'        If Not InInactivePeriod Then
'            If IsLoanEndedEvent(EventText) Then
'                StartRow = r
'                InInactivePeriod = True
'            End If
'        ElseIf IsLoanRestartedEvent(EventText) Then
'            Result.Add Array(StartRow, r)
'            InInactivePeriod = False
'        End If
'    Next r
'
'    Set GetInactiveLoanEpisodes = Result
'End Function
'
'Private Sub AddInactiveLoanBandSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long, _
'    ByVal BandHeight As Double)
'
'    Dim Episodes As Collection
'    Dim Episode As Variant
'
'    Set Episodes = GetInactiveLoanEpisodes(wsData, Columns, LastRow)
'
'    For Each Episode In Episodes
'        AddSingleBandEpisodeSeries ch, wsData, Columns.DateCol, LastRow, _
'            CLng(Episode(0)), CLng(Episode(1)), BandHeight, _
'            "Inactive Episode", ColorInactiveFill, ColorInactiveLine, _
'            TransparencyInactiveBand, TransparencyInactiveLine
'    Next Episode
'End Sub
'
'Private Sub AddSingleBandEpisodeSeries( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByVal DateCol As Long, _
'    ByVal LastRow As Long, _
'    ByVal StartRow As Long, _
'    ByVal EndRow As Long, _
'    ByVal BandHeight As Double, _
'    ByVal SeriesName As String, _
'    ByVal FillColour As Long, _
'    ByVal LineColour As Long, _
'    ByVal FillTransparency As Double, _
'    ByVal LineTransparency As Double)
'
'    Dim BandValues() As Variant
'    Dim DataSeries As Series
'    Dim PointCount As Long
'    Dim PointIndex As Long
'    Dim SourceRow As Long
'
'    PointCount = LastRow - FIRST_DATA_ROW + 1
'    ReDim BandValues(1 To PointCount)
'
'    For SourceRow = FIRST_DATA_ROW To LastRow
'        PointIndex = SourceRow - FIRST_DATA_ROW + 1
'        If SourceRow >= StartRow And SourceRow <= EndRow Then
'            BandValues(PointIndex) = BandHeight
'        Else
'            BandValues(PointIndex) = CVErr(xlErrNA)
'        End If
'    Next SourceRow
'
'    Set DataSeries = ch.Chart.SeriesCollection.NewSeries
'    With DataSeries
'        .name = SeriesName
'        .XValues = DataRange(wsData, DateCol, LastRow)
'        .Values = BandValues
'        .AxisGroup = xlPrimary
'        .ChartType = xlArea
'        .MarkerStyle = xlMarkerStyleNone
'
'        With .Format.Fill
'            .Visible = msoTrue
'            .Solid
'            .ForeColor.RGB = FillColour
'            .Transparency = FillTransparency
'        End With
'
'        With .Format.Line
'            If LineTransparency >= 1 Then
'                .Visible = msoFalse
'            Else
'                .Visible = msoTrue
'                .ForeColor.RGB = LineColour
'                .Transparency = LineTransparency
'                .Weight = 1
'                .DashStyle = msoLineSolid
'            End If
'        End With
'    End With
'End Sub
'
'Private Sub AddMCDurationLabels( _
'    ByVal ch As ChartObject, _
'    ByVal wsData As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal LastRow As Long)
'
'    Dim Episodes As Collection
'    Dim Episode As Variant
'    Dim DurationDays As Long
'    Dim LabelText As String
'    Dim ReasonRow As Long
'    Dim IsTechnical As Boolean
'
'    Set Episodes = GetStatusEpisodes( _
'        wsData, Columns.MCCol, Columns.MCClearedCol, LastRow)
'
'    ch.Activate
'    DoEvents
'
'    For Each Episode In Episodes
'        DurationDays = DateDiff( _
'            "d", _
'            CDate(wsData.Cells(CLng(Episode(0)), Columns.DateCol).Value), _
'            CDate(wsData.Cells(CLng(Episode(1)), Columns.DateCol).Value))
'
'        If CBool(Episode(2)) Then
'            LabelText = DurationDays & "d"
'        Else
'            LabelText = (DurationDays + 1) & "d (open)"
'        End If
'
'        ReasonRow = CLng(Episode(1))
'        If CBool(Episode(2)) Then ReasonRow = ReasonRow - 1
'        If ReasonRow < CLng(Episode(0)) Then ReasonRow = CLng(Episode(0))
'        IsTechnical = IsTechnicalReason( _
'            SafeCellText(wsData.Cells(ReasonRow, Columns.ReasonCol)))
'
'        AddMCTextLabelByRows ch, LastRow, CLng(Episode(0)), _
'                             CLng(Episode(1)), LabelText, IsTechnical
'    Next Episode
'End Sub
'
'Private Sub AddMCTextLabelByRows( _
'    ByVal ch As ChartObject, _
'    ByVal LastRow As Long, _
'    ByVal StartRow As Long, _
'    ByVal EndRow As Long, _
'    ByVal LabelText As String, _
'    ByVal IsTechnical As Boolean)
'
'    Dim PointCount As Long
'    Dim MidPoint As Double
'    Dim RelativeX As Double
'    Dim XPos As Double
'    Dim YPos As Double
'    Dim PlotLeft As Double
'    Dim PlotRight As Double
'    Dim LabelWidth As Double
'    Dim LabelFillColour As Long
'    Dim LabelFontColour As Long
'    Dim LabelFillTransparency As Double
'    Dim shp As Shape
'
'    PointCount = LastRow - FIRST_DATA_ROW + 1
'    If PointCount < 1 Then Exit Sub
'
'    MidPoint = ((StartRow - 1) + (EndRow - 1)) / 2
'
'    If ch.Chart.Axes(xlCategory, xlPrimary).AxisBetweenCategories Then
'        RelativeX = (MidPoint - 0.5) / PointCount
'    ElseIf PointCount > 1 Then
'        RelativeX = (MidPoint - 1) / (PointCount - 1)
'    Else
'        RelativeX = 0.5
'    End If
'    RelativeX = Application.Max(0, Application.Min(1, RelativeX))
'
'    With ch.Chart.PlotArea
'        PlotLeft = .InsideLeft
'        PlotRight = .InsideLeft + .InsideWidth
'        XPos = .InsideLeft + RelativeX * .InsideWidth
'        YPos = .InsideTop + .InsideHeight + _
'               RISK_BAND_SHAPE_OFFSET + RISK_BAND_SHAPE_HEIGHT + _
'               RISK_BAND_LABEL_GAP
'    End With
'
'    If YPos + LABEL_HEIGHT > _
'       ch.Chart.ChartArea.Height - CHART_BOTTOM_MARGIN Then _
'        Err.Raise vbObjectError + 1301, "AddMCTextLabelByRows", _
'                  "No chart space is available below the risk band strip."
'
'    LabelWidth = Application.Min( _
'        PlotRight - PlotLeft, _
'        Application.Max( _
'            LABEL_MIN_WIDTH, _
'            Len(LabelText) * LABEL_CHARACTER_WIDTH + _
'            2 * LABEL_HORIZONTAL_MARGIN))
'
'    XPos = Application.Max( _
'        PlotLeft + LabelWidth / 2, _
'        Application.Min(PlotRight - LabelWidth / 2, XPos))
'
'    If IsTechnical Then
'        LabelFontColour = ColorTechnicalLabelText
'    Else
'        LabelFillColour = ColorLabelBackground
'        LabelFontColour = ColorLabelText
'        LabelFillTransparency = TransparencyLabel
'    End If
'
'    Set shp = ch.Chart.Shapes.AddTextbox( _
'        msoTextOrientationHorizontal, XPos - LabelWidth / 2, YPos, _
'        LabelWidth, LABEL_HEIGHT)
'
'    With shp
'        .name = "MC_Duration_Label_" & Format$(ch.Chart.Shapes.Count, "000")
'        .Visible = msoTrue
'        .TextFrame.Characters.Text = LabelText
'        .TextFrame.MarginLeft = LABEL_HORIZONTAL_MARGIN
'        .TextFrame.MarginRight = LABEL_HORIZONTAL_MARGIN
'        .TextFrame.HorizontalAlignment = xlCenter
'        .TextFrame.VerticalAlignment = xlCenter
'        With .TextFrame.Characters.Font
'            .Bold = True
'            .Size = 9
'            .Color = LabelFontColour
'        End With
'        With .Fill
'            If IsTechnical Then
'                .Visible = msoFalse
'            Else
'                .Visible = msoTrue
'                .Solid
'                .ForeColor.RGB = LabelFillColour
'                .Transparency = LabelFillTransparency
'            End If
'        End With
'        With .Line
'            If IsTechnical Then
'                .Visible = msoTrue
'                .ForeColor.RGB = ColorTechnicalLabelBorder
'                .Transparency = 0
'                .Weight = 0.5
''                .DashStyle = msoLineRoundDot
'            Else
'                .Visible = msoFalse
'            End If
'        End With
'        .ZOrder msoBringToFront
'    End With
'End Sub
'
'Private Sub AddLTVLegendOnlySeries(ByVal ch As ChartObject)
'    Dim LegendSeries As Series
'
'    Set LegendSeries = ch.Chart.SeriesCollection.NewSeries
'    With LegendSeries
'        .name = "LTV"
'        .XValues = Array(0)
'        .Values = Array(0)
'    End With
'    FormatLineSeries LegendSeries, ColorLTV
'End Sub
'
'Private Sub RebuildMainChartLegend(ByVal ch As ChartObject)
'    With ch.Chart
'        .HasLegend = False
'        .HasLegend = True
'    End With
'
'    ch.Activate
'    DoEvents
'
'    RemoveHelperAreaLegendEntries ch
'    RemoveDuplicateRiskBandLegendEntries ch
'End Sub
'
'Private Sub RemoveHelperAreaLegendEntries(ByVal ch As ChartObject)
'    Dim i As Long
'
'    For i = ch.Chart.Legend.LegendEntries.Count To 1 Step -1
'        If IsHelperAreaLegendEntry( _
'                ch.Chart.Legend.LegendEntries(i)) Then _
'            ch.Chart.Legend.LegendEntries(i).Delete
'    Next i
'End Sub
'
'Private Function IsHelperAreaLegendEntry(ByVal Entry As Object) As Boolean
'    On Error GoTo ExitFunction
'
'    With Entry.LegendKey.Format
'        If .Fill.Visible <> msoTrue And .Line.Visible <> msoTrue Then
'            IsHelperAreaLegendEntry = True
'        ElseIf .Fill.Visible = msoTrue And .Line.Visible <> msoTrue Then
'            Select Case .Fill.ForeColor.RGB
'                Case ColorApprovedArea, ColorHaircut, ColorDanger, _
'                     ColorHCVOverlap, ColorMTMOverlap
'                    IsHelperAreaLegendEntry = True
'            End Select
'        End If
'    End With
'
'ExitFunction:
'End Function
'
'Private Sub RemoveDuplicateRiskBandLegendEntries(ByVal ch As ChartObject)
'    Dim SeenKeys As Object
'    Dim LegendKey As String
'    Dim i As Long
'
'    Set SeenKeys = CreateObject("Scripting.Dictionary")
'
'    For i = ch.Chart.Legend.LegendEntries.Count To 1 Step -1
'        LegendKey = RiskBandLegendKey(ch.Chart.Legend.LegendEntries(i))
'
'        If Len(LegendKey) > 0 Then
'            If SeenKeys.Exists(LegendKey) Then
'                ch.Chart.Legend.LegendEntries(i).Delete
'            Else
'                SeenKeys.Add LegendKey, True
'            End If
'        End If
'    Next i
'End Sub
'
'Private Function RiskBandLegendKey(ByVal Entry As Object) As String
'    Dim FillColour As Long
'
'    On Error GoTo ExitFunction
'    With Entry.LegendKey.Format.Fill
'        If .Visible <> msoTrue Then Exit Function
'        FillColour = .ForeColor.RGB
'    End With
'
'    Select Case FillColour
'        Case ColorMCFill, ColorSFFill, ColorInactiveFill
'            RiskBandLegendKey = CStr(FillColour)
'    End Select
'
'ExitFunction:
'End Function
'
'Private Sub ReserveRiskBandShapeSpace(ByVal ch As ChartObject)
'    Dim RequiredBottomSpace As Double
'    Dim MaximumInsideHeight As Double
'    Dim Pass As Long
'
'    RequiredBottomSpace = RISK_BAND_SHAPE_OFFSET + _
'                          RISK_BAND_SHAPE_HEIGHT + _
'                          RISK_BAND_LABEL_GAP + LABEL_HEIGHT + _
'                          CHART_BOTTOM_MARGIN
'
'    For Pass = 1 To 2
'        ch.Activate
'        DoEvents
'
'        With ch.Chart.PlotArea
'            MaximumInsideHeight = ch.Chart.ChartArea.Height - _
'                                  .InsideTop - RequiredBottomSpace
'            If MaximumInsideHeight <= 0 Then _
'                Err.Raise vbObjectError + 1299, _
'                          "ReserveRiskBandShapeSpace", _
'                          "The main chart is too short for the risk band strip."
'            If .InsideHeight > MaximumInsideHeight Then _
'                .InsideHeight = MaximumInsideHeight
'        End With
'    Next Pass
'End Sub
'
'Private Sub AlignChartPlotAreas( _
'    ByVal UpperChart As ChartObject, _
'    ByVal LowerChart As ChartObject)
'
'    Dim TargetInsideLeft As Double
'    Dim TargetInsideWidth As Double
'    Dim Pass As Long
'
'    UpperChart.Activate
'    DoEvents
'    LowerChart.Activate
'    DoEvents
'
'    TargetInsideLeft = UpperChart.Chart.PlotArea.InsideLeft
'    TargetInsideWidth = UpperChart.Chart.PlotArea.InsideWidth
'
'    For Pass = 1 To 2
'        With LowerChart.Chart.PlotArea
'            .Left = .Left + TargetInsideLeft - .InsideLeft
'            .Width = .Width + TargetInsideWidth - .InsideWidth
'        End With
'        LowerChart.Activate
'        DoEvents
'    Next Pass
'End Sub
'
'Private Sub PositionTitleAndLegend(ByVal ch As ChartObject)
'    Const OUTER_MARGIN As Double = 12
'    Const VERTICAL_MARGIN As Double = 6
'
'    With ch.Chart
'        .HasTitle = True
'        .HasLegend = True
'        .Legend.Position = xlLegendPositionTop
'
'        ch.Activate
'        DoEvents
'
'        .ChartTitle.Left = OUTER_MARGIN
'        .ChartTitle.Top = VERTICAL_MARGIN
'        .Legend.Left = .ChartArea.Width - .Legend.Width - OUTER_MARGIN
'        .Legend.Top = VERTICAL_MARGIN
'    End With
'
'    FormatChartText ch
'End Sub
'
'Private Sub LoadSourceColumns( _
'    ByVal ws As Worksheet, _
'    ByRef Columns As JourneyColumns)
'
'    LoadCoreColumns ws, Columns, "BuildJourneyTimeSeries"
'    Columns.LTVCol = FindColumnByHeader(ws, "LTV")
'    Columns.MCClearedCol = FindColumnByHeader(ws, "MC Cleared")
'    Columns.DeltaApprovedCol = FindColumnByHeader(ws, "Delta Approved")
'    Columns.DeltaDrawnCol = FindColumnByHeader(ws, "Delta Drawn")
'    Columns.DeltaHCVCol = FindColumnByHeader(ws, "Delta HCV")
'    Columns.DeltaLTVCol = FindColumnByHeader(ws, "Delta LTV")
'End Sub
'
'Private Sub LoadDashboardColumns( _
'    ByVal ws As Worksheet, _
'    ByRef Columns As JourneyColumns)
'
'    LoadCoreColumns ws, Columns, "GenerateJourneyDashboard"
'    Columns.LTVCol = RequireColumn(ws, "LTV", "GenerateJourneyDashboard")
'    Columns.MCCol = RequireColumn(ws, "MC", "GenerateJourneyDashboard")
'    Columns.SFCol = RequireColumn(ws, "SF", "GenerateJourneyDashboard")
'    Columns.MCClearedCol = RequireColumn(ws, "MC Cleared", "GenerateJourneyDashboard")
'    Columns.ReasonCol = RequireColumn(ws, "Reason MC/SF", "GenerateJourneyDashboard")
'End Sub
'
'Private Sub LoadCoreColumns( _
'    ByVal ws As Worksheet, _
'    ByRef Columns As JourneyColumns, _
'    ByVal CallerName As String)
'
'    Columns.DateCol = RequireColumn(ws, "Snapshot Date", CallerName)
'    Columns.EventCol = RequireColumn(ws, "Event", CallerName)
'    Columns.ApprovedCol = RequireColumn(ws, "Approved", CallerName)
'    Columns.DrawnCol = RequireColumn(ws, "Drawn", CallerName)
'    Columns.MTMCol = RequireColumn(ws, "MTM", CallerName)
'    Columns.HCVCol = RequireColumn(ws, "HCV", CallerName)
'End Sub
'
'Private Function RequireColumn( _
'    ByVal ws As Worksheet, _
'    ByVal HeaderText As String, _
'    ByVal CallerName As String) As Long
'
'    RequireColumn = FindColumnByHeader(ws, HeaderText)
'    If RequireColumn = 0 Then _
'        Err.Raise vbObjectError + 1200, CallerName, _
'                  "Required column not found: " & HeaderText
'End Function
'
'Private Function DataRange( _
'    ByVal ws As Worksheet, _
'    ByVal ColumnNumber As Long, _
'    ByVal LastRow As Long) As Range
'
'    Set DataRange = ws.Range( _
'        ws.Cells(FIRST_DATA_ROW, ColumnNumber), ws.Cells(LastRow, ColumnNumber))
'End Function
'
'Private Function ConstantValues( _
'    ByVal Count As Long, _
'    ByVal Value As Double) As Variant
'
'    Dim Result() As Double
'    Dim i As Long
'
'    ReDim Result(1 To Count)
'    For i = 1 To Count
'        Result(i) = Value
'    Next i
'    ConstantValues = Result
'End Function
'
'Private Function DateKeyFor(ByVal Value As Variant) As String
'    DateKeyFor = Format$(CDate(Value), "yyyymmdd")
'End Function
'
'Private Sub CopyRowValues( _
'    ByVal SourceSheet As Worksheet, _
'    ByVal SourceRow As Long, _
'    ByVal TargetSheet As Worksheet, _
'    ByVal TargetRow As Long, _
'    ByVal LastCol As Long)
'
'    TargetSheet.Range(TargetSheet.Cells(TargetRow, 1), _
'                      TargetSheet.Cells(TargetRow, LastCol)).Value = _
'        SourceSheet.Range(SourceSheet.Cells(SourceRow, 1), _
'                          SourceSheet.Cells(SourceRow, LastCol)).Value
'End Sub
'
'Private Sub ClearColumns( _
'    ByVal ws As Worksheet, _
'    ByVal RowNumber As Long, _
'    ByVal ColumnNumbers As Variant)
'
'    Dim ColumnNumber As Variant
'
'    For Each ColumnNumber In ColumnNumbers
'        If CLng(ColumnNumber) > 0 Then _
'            ws.Cells(RowNumber, CLng(ColumnNumber)).ClearContents
'    Next ColumnNumber
'End Sub
'
'Private Sub InitialiseDashboardTheme()
'    If USE_DARK_THEME Then
'        InitialiseDarkDashboardTheme
'    Else
'        InitialiseLightDashboardTheme
'    End If
'End Sub
'
'Private Sub InitialiseLightDashboardTheme()
'    ColorChartBackground = RGB(255, 255, 255)
'    ColorPlotBackground = RGB(255, 255, 255)
'    ColorChartText = RGB(60, 60, 60)
'    ColorBorder = RGB(220, 220, 220)
'    ColorGridline = RGB(225, 225, 225)
'    ColorMTM = RGB(231, 171, 120)
'    ColorHCV = RGB(243, 214, 112)
'    ColorApproved = RGB(45, 45, 45)
'    ColorApprovedArea = RGB(66, 133, 244)
'    ColorDrawn = RGB(164, 139, 193)
'    ColorLTV = RGB(66, 133, 244)
'    ColorLTVArea = RGB(66, 133, 244)
'    ColorHaircut = RGB(245, 166, 35)
'    ColorDanger = RGB(220, 65, 65)
'    ColorHCVOverlap = RGB(245, 130, 32)
'    ColorMTMOverlap = RGB(200, 0, 0)
'    ColorMCFill = RGB(255, 165, 80)
'    ColorMCLine = RGB(255, 100, 0)
'    ColorSFFill = RGB(255, 150, 150)
'    ColorInactiveFill = RGB(190, 190, 190)
'    ColorInactiveLine = RGB(135, 135, 135)
'    ColorThreshold = RGB(220, 0, 0)
'    ColorLabelBackground = RGB(255, 255, 255)
'    ColorLabelText = ColorMCLine
'    ColorTechnicalLabelText = RGB(255, 255, 255)
'    ColorTechnicalLabelBorder = RGB(255, 255, 255)
'
'    TransparencyGridline = 0.15
'    TransparencyApprovedArea = 0.82
'    TransparencyLTVArea = 0.82
'    TransparencyHCVBand = 0.8
'    TransparencyMTMBand = 0.5
'    TransparencyHCVOverlap = 0.35
'    TransparencyMTMOverlap = 0.2
'    TransparencyRiskBand = 0.85
'    TransparencyRiskBandAccent = 0.08
'    TransparencyInactiveBand = 0.72
'    TransparencyInactiveLine = 0.35
'    TransparencyLabel = 0.15
'End Sub
'
'Private Sub InitialiseDarkDashboardTheme()
'    ColorChartBackground = RGB(30, 30, 30)
'    ColorPlotBackground = RGB(38, 38, 38)
'    ColorChartText = RGB(235, 235, 235)
'    ColorBorder = RGB(82, 82, 82)
'    ColorGridline = RGB(110, 110, 110)
'    ColorMTM = RGB(210, 85, 40)
''    ColorHCV = RGB(255, 212, 75)
'    ColorHCV = RGB(210, 170, 40)
'    ColorApproved = RGB(240, 50, 50)
'    ColorApprovedArea = RGB(80, 145, 50)
'    ColorDrawn = RGB(242, 242, 242)
'    ColorLTV = RGB(80, 145, 50)
'    ColorLTVArea = RGB(80, 145, 50)
'    ColorHaircut = RGB(255, 179, 71)
'    ColorDanger = RGB(192, 0, 0)
'    ColorHCVOverlap = RGB(255, 145, 40)
'    ColorMTMOverlap = RGB(255, 65, 65)
''    ColorDanger = RGB(255, 82, 82)
'    ColorMCFill = RGB(255, 174, 90)
'    ColorMCLine = RGB(255, 132, 40)
'    ColorSFFill = RGB(255, 110, 120)
'    ColorInactiveFill = RGB(130, 130, 130)
'    ColorInactiveLine = RGB(180, 180, 180)
'    ColorThreshold = RGB(255, 99, 99)
'    ColorLabelBackground = ColorChartText
'    ColorLabelText = ColorChartBackground
'    ColorTechnicalLabelText = ColorChartText
'    ColorTechnicalLabelBorder = RGB(255, 255, 255)
'
'    TransparencyGridline = 0.35
'    TransparencyApprovedArea = 0.82
'    TransparencyLTVArea = 0.82
'    TransparencyHCVBand = 0.8
'    TransparencyMTMBand = 0.62
'    TransparencyHCVOverlap = 0.3
'    TransparencyMTMOverlap = 0.15
'    TransparencyRiskBand = 0.85
'    TransparencyRiskBandAccent = 0
'    TransparencyInactiveBand = 0.65
'    TransparencyInactiveLine = 0.25
'    TransparencyLabel = 0
'End Sub
'
