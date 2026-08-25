Attribute VB_Name = "JourneyPositionAnalysis"
Option Explicit

' RELEASE: ONE_CLICK_EMBEDDED_DASHBOARD_20260813_V2
' Shared report styles and highlights are consumed by Position Analysis and Dashboard.

' Sole implementation for position-analysis buttons and output layout.

Private Const POSITION_BUTTON_HEADER As String = "Position Analysis"
Private Const POSITION_BUTTON_PREFIX As String = "PositionAnalysis_"
Private Const LEGACY_JOURNEY_BUTTON_PREFIX As String = "JourneyPositionAnalysis_"
Private Const POSITION_BUTTON_CONTEXT As String = "POSITION_ANALYSIS_V1"
Private Const POSITION_FILE_SUFFIX As String = "_Lombard_Loans_ITA_Positions.csv"
Private Const POSITION_ANALYSIS_SHEET As String = "Position Change Analysis"
Private Const MIN_MEANINGFUL_MTM As Double = 0.00001
Private Const LTV_NOT_MEANINGFUL As String = "N/M"

Private Const POS_GUARANTEE As Long = 0
Private Const POS_ISIN As Long = 1
Private Const POS_NAME As Long = 2
Private Const POS_ASSET_TYPE As Long = 3
Private Const POS_CCY As Long = 4
Private Const POS_ISSUER As Long = 5
Private Const POS_QUANTITY As Long = 6
Private Const POS_PRICE_EUR As Long = 7
Private Const POS_VALUE As Long = 8
Private Const POS_WEIGHT As Long = 9
Private Const POS_MAX_LTV As Long = 10
Private Const POS_ELIGIBLE_MV As Long = 11
Private Const POS_HCV As Long = 12
Private Const POS_ABOVE_LIMIT As Long = 13
Private Const POS_COMMENT As Long = 14
Private Const POS_FIELD_COUNT As Long = 15

Public Sub ApplyUnifiedReportTitle(ByVal TargetRange As Range)
    With TargetRange
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 78, 121)
        .Font.name = "Arial"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
End Sub

Public Sub ApplyUnifiedReportSubtitle(ByVal TargetRange As Range)
    With TargetRange
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(221, 235, 247)
        .Font.name = "Arial"
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
End Sub

Public Sub ApplyUnifiedSectionTitle(ByVal TargetRange As Range, _
                                    Optional ByVal TitleText As String = "")
    With TargetRange
        If TitleText <> "" Then .Value = TitleText
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(68, 114, 196)
        .Font.name = "Arial"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub ApplyUnifiedSignalCell(ByVal TargetCell As Range, _
                                  ByVal FillColor As Long, _
                                  ByVal FontColor As Long)
    With TargetCell
        .Interior.Pattern = xlSolid
        .Interior.Color = FillColor
        .Font.Color = FontColor
        .Font.Bold = True
    End With
End Sub

Public Sub ApplyUnifiedStatusHighlight(ByVal TargetCell As Range, _
                                       ByVal StatusText As String)
    TargetCell.HorizontalAlignment = xlCenter

    Select Case StatusText
        Case "Ended"
            ApplyUnifiedSignalCell _
                TargetCell, RGB(217, 217, 217), RGB(89, 89, 89)
        Case "Shortfall"
            ApplyUnifiedSignalCell _
                TargetCell, RGB(255, 120, 120), RGB(150, 0, 0)
        Case "Margin Call"
            ApplyUnifiedSignalCell _
                TargetCell, RGB(255, 205, 80), RGB(120, 65, 0)
        Case Else
            ApplyUnifiedSignalCell _
                TargetCell, RGB(198, 239, 206), RGB(84, 130, 53)
    End Select
End Sub

Public Sub ApplyUnifiedDeltaHighlight(ByVal TargetCell As Range, _
                                      ByVal DeltaValue As Double)
    If DeltaValue > 0 Then
        ApplyUnifiedSignalCell _
            TargetCell, RGB(198, 239, 206), RGB(84, 130, 53)
    ElseIf DeltaValue < 0 Then
        ApplyUnifiedSignalCell _
            TargetCell, RGB(255, 199, 206), RGB(156, 0, 6)
    End If
End Sub

Public Sub AddPositionAnalysisButtons(ByVal ws As Worksheet)
    Dim SourceTable As ListObject
    Dim ButtonColumn As ListColumn
    Dim DateCol As Long
    Dim LastRow As Long
    Dim r As Long

    If FindColumnByHeader(ws, POSITION_BUTTON_HEADER) = 0 Then
        If ws.ListObjects.Count > 0 Then Set SourceTable = ws.ListObjects(1)

        If SourceTable Is Nothing Then
            ws.Columns(1).Insert Shift:=xlToRight
            ws.Cells(1, 1).Value = POSITION_BUTTON_HEADER
        ElseIf SourceTable.Range.Column = 1 Then
            Set ButtonColumn = SourceTable.ListColumns.Add(1)
            ButtonColumn.name = POSITION_BUTTON_HEADER
        Else
            ws.Columns(1).Insert Shift:=xlToRight
            ws.Cells(1, 1).Value = POSITION_BUTTON_HEADER
        End If
    End If

    ws.Columns(1).ColumnWidth = 12
    DateCol = FindColumnByHeader(ws, "Snapshot Date")

    If DateCol = 0 Then
        Err.Raise vbObjectError + 1805, "AddPositionAnalysisButtons", _
                  "The Snapshot Date column could not be found."
    End If

    LastRow = ws.Cells(ws.Rows.Count, DateCol).End(xlUp).Row
    ClearPositionAnalysisButtons ws

    For r = 3 To LastRow
        AddPositionAnalysisButton ws.Cells(r, 1), ws, r
    Next r
End Sub

Private Sub ClearPositionAnalysisButtons(ByVal ws As Worksheet)
    Dim ButtonColumn As Long
    Dim i As Long

    ButtonColumn = FindColumnByHeader(ws, POSITION_BUTTON_HEADER)

    For i = ws.Shapes.Count To 1 Step -1
        If IsPositionAnalysisButton(ws.Shapes(i)) Or _
           IsShapeInPositionButtonColumn(ws.Shapes(i), ButtonColumn) Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Private Function IsShapeInPositionButtonColumn(ByVal Candidate As Shape, _
                                               ByVal ButtonColumn As Long) As Boolean
    Dim ShapeRow As Long
    Dim ShapeColumn As Long

    If ButtonColumn = 0 Then Exit Function

    On Error Resume Next
    ShapeRow = Candidate.TopLeftCell.Row
    ShapeColumn = Candidate.TopLeftCell.Column
    On Error GoTo 0

    IsShapeInPositionButtonColumn = _
        ShapeRow >= 2 And ShapeColumn = ButtonColumn
End Function

Private Function IsPositionAnalysisButton(ByVal Candidate As Shape) As Boolean
    Dim ActionName As String
    Dim ContextText As String

    If Left$(Candidate.name, Len(POSITION_BUTTON_PREFIX)) = _
       POSITION_BUTTON_PREFIX Then
        IsPositionAnalysisButton = True
        Exit Function
    End If

    If Left$(Candidate.name, Len(LEGACY_JOURNEY_BUTTON_PREFIX)) = _
       LEGACY_JOURNEY_BUTTON_PREFIX Then
        IsPositionAnalysisButton = True
        Exit Function
    End If

    On Error Resume Next
    ActionName = Candidate.OnAction
    ContextText = Candidate.AlternativeText
    On Error GoTo 0

    If InStr(1, ActionName, "AnalyzePositionChanges", vbTextCompare) > 0 Then
        IsPositionAnalysisButton = True
        Exit Function
    End If

    IsPositionAnalysisButton = _
        Left$(ContextText, Len("POSITION_ANALYSIS_")) = "POSITION_ANALYSIS_"
End Function

Public Sub AddPositionAnalysisButton(ByVal AnchorCell As Range, _
                                     ByVal SourceSheet As Worksheet, _
                                     ByVal SourceRow As Long)
    Dim ButtonShape As Shape
    Dim ButtonName As String
    Dim DateCol As Long
    Dim NDGCol As Long
    Dim TargetNDG As String
    Dim PreviousDate As Date
    Dim CurrentDate As Date

    DateCol = FindColumnByHeader(SourceSheet, "Snapshot Date")
    NDGCol = FindColumnByHeader(SourceSheet, "NDG")

    If SourceRow < 3 Or DateCol = 0 Or NDGCol = 0 Then
        Err.Raise vbObjectError + 1802, "AddPositionAnalysisButton", _
                  "The source row cannot be compared with a previous snapshot."
    End If

    If Not IsDate(SourceSheet.Cells(SourceRow - 1, DateCol).Value) Or _
       Not IsDate(SourceSheet.Cells(SourceRow, DateCol).Value) Then
        Err.Raise vbObjectError + 1802, "AddPositionAnalysisButton", _
                  "The source row cannot be compared with a previous snapshot."
    End If

    TargetNDG = SafeCellText(SourceSheet.Cells(SourceRow, NDGCol))
    PreviousDate = CDate(SourceSheet.Cells(SourceRow - 1, DateCol).Value)
    CurrentDate = CDate(SourceSheet.Cells(SourceRow, DateCol).Value)

    ButtonName = POSITION_BUTTON_PREFIX & AnchorCell.Row & "_" & AnchorCell.Column

    On Error Resume Next
    AnchorCell.Worksheet.Shapes(ButtonName).Delete
    On Error GoTo 0

    Set ButtonShape = AnchorCell.Worksheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, AnchorCell.Left + 1, AnchorCell.Top + 1, _
        AnchorCell.Width - 2, AnchorCell.Height - 2)

    With ButtonShape
        .name = ButtonName
        .OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & _
                    "'!AnalyzePositionChanges"
        .Placement = xlMoveAndSize
        .AlternativeText = ButtonContext( _
            TargetNDG, PreviousDate, CurrentDate, SourceSheet.name, SourceRow)
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Line.Visible = msoFalse

        With .TextFrame2
            .TextRange.Text = "Analyze"
            .TextRange.Font.name = "Arial"
            .TextRange.Font.Size = 8
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
    End With
End Sub

Public Sub AnalyzePositionChanges()
    Dim PreviousScreenUpdating As Boolean
    Dim CallerShape As Shape
    Dim Context As Object
    Dim AccountMetrics As Object
    Dim PreviousPositions As Object
    Dim CurrentPositions As Object
    Dim PreviousDate As Date
    Dim CurrentDate As Date
    Dim TargetNDG As String
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    On Error GoTo CleanUp

    PreviousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    If TypeName(Application.Caller) <> "String" Then
        Err.Raise vbObjectError + 1800, "AnalyzePositionChanges", _
                  "Run this procedure from a position comparison button."
    End If

    Set CallerShape = ActiveSheet.Shapes(CStr(Application.Caller))
    Set Context = ParseButtonContext(CallerShape.AlternativeText)
    PreviousDate = Context("PreviousDate")
    CurrentDate = Context("CurrentDate")
    TargetNDG = Context("NDG")

    If TargetNDG = "" Then
        Err.Raise vbObjectError + 1806, "AnalyzePositionChanges", _
                  "The selected comparison has no NDG."
    End If

    If Context("SourceSheet") <> "" And Context("SourceRow") > 2 Then
        Set AccountMetrics = ReadAccountMetrics( _
            ThisWorkbook.Worksheets(Context("SourceSheet")), Context("SourceRow"))
    End If

    Set PreviousPositions = ReadPositionsForNDG(PreviousDate, TargetNDG)
    Set CurrentPositions = ReadPositionsForNDG(CurrentDate, TargetNDG)
    BuildPositionChangeAnalysis _
        PreviousDate, CurrentDate, TargetNDG, PreviousPositions, _
        CurrentPositions, AccountMetrics

CleanUp:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
    Application.ScreenUpdating = PreviousScreenUpdating

    If ErrorNumber <> 0 Then Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Sub

Private Function ButtonContext(ByVal TargetNDG As String, _
                               ByVal PreviousDate As Date, _
                               ByVal CurrentDate As Date, _
                               ByVal SourceSheetName As String, _
                               ByVal SourceRow As Long) As String
    ButtonContext = Join( _
        Array(POSITION_BUTTON_CONTEXT, TargetNDG, CStr(CDbl(PreviousDate)), _
              CStr(CDbl(CurrentDate)), SourceSheetName, CStr(SourceRow)), Chr$(9))
End Function

Private Function ParseButtonContext(ByVal ContextText As String) As Object
    Dim Context As Object
    Dim Parts As Variant

    Parts = Split(ContextText, Chr$(9))
    If UBound(Parts) <> 5 Or Parts(0) <> POSITION_BUTTON_CONTEXT Then
        Err.Raise vbObjectError + 1801, "AnalyzePositionChanges", _
                  "The position comparison button has invalid context."
    End If

    If Not IsNumeric(Parts(2)) Or Not IsNumeric(Parts(3)) Or _
       Not IsNumeric(Parts(5)) Then
        Err.Raise vbObjectError + 1802, "AnalyzePositionChanges", _
                  "The position comparison button has invalid dates or source row."
    End If

    Set Context = CreateObject("Scripting.Dictionary")
    Context("NDG") = Parts(1)
    Context("PreviousDate") = CDate(CDbl(Parts(2)))
    Context("CurrentDate") = CDate(CDbl(Parts(3)))
    Context("SourceSheet") = Parts(4)
    Context("SourceRow") = CLng(Parts(5))
    Set ParseButtonContext = Context
End Function

Private Function ReadPositionsForNDG(ByVal SnapshotDate As Date, _
                                     ByVal TargetNDG As String) As Object
    Dim Positions As Object
    Dim Lines As Variant
    Dim Header As Variant
    Dim Fields As Variant
    Dim Record As Variant
    Dim BasePath As String
    Dim FilePath As String
    Dim Key As String
    Dim RowNDG As String
    Dim RowQuantity As Double
    Dim RowPrice As Double
    Dim RowValue As Double
    Dim RowMaxLTV As Double
    Dim QuantityWeight As Double
    Dim ValueWeight As Double
    Dim idxNDG As Long
    Dim idxGuarantee As Long
    Dim idxISIN As Long
    Dim idxName As Long
    Dim idxAssetType As Long
    Dim idxCcy As Long
    Dim idxIssuer As Long
    Dim idxQuantity As Long
    Dim idxPriceEUR As Long
    Dim idxValue As Long
    Dim idxWeight As Long
    Dim idxMaxLTV As Long
    Dim idxEligibleMV As Long
    Dim idxHCV As Long
    Dim idxAboveLimit As Long
    Dim idxComment As Long
    Dim r As Long

    BasePath = PathSelection()
    If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
    FilePath = BasePath & Format$(SnapshotDate, "yyyymmdd") & POSITION_FILE_SUFFIX

    If Dir(FilePath) = "" Then
        Err.Raise vbObjectError + 1803, "ReadPositionsForNDG", _
                  "Positions file not found:" & vbCrLf & FilePath
    End If

    If FileLen(FilePath) = 0 Then
        Err.Raise vbObjectError + 1804, "ReadPositionsForNDG", _
                  "Positions file is empty:" & vbCrLf & FilePath
    End If

    Lines = ReadAllLines(FilePath)
    Header = Split(Lines(0), ";")

    idxNDG = RequiredPositionHeader(Header, "NDG", FilePath)
    idxGuarantee = RequiredPositionHeader(Header, "CO_FT_GAR", FilePath)
    idxISIN = RequiredPositionHeader(Header, "ISIN", FilePath)
    idxName = RequiredPositionHeader(Header, "Security Name", FilePath)
    idxAssetType = RequiredPositionHeader(Header, "Asset Type / Classification", FilePath)
    idxCcy = RequiredPositionHeader(Header, "Pricing Currency", FilePath)
    idxQuantity = RequiredPositionHeader(Header, "No. Securities", FilePath)
    idxPriceEUR = RequiredPositionHeader(Header, "Price EUR", FilePath)
    idxValue = RequiredPositionHeader(Header, "Position Value", FilePath)
    idxWeight = RequiredPositionHeader(Header, "Weight %", FilePath)
    idxMaxLTV = RequiredPositionHeader(Header, "Max LTV %", FilePath)
    idxEligibleMV = RequiredPositionHeader(Header, "Eligible MV", FilePath)
    idxHCV = RequiredPositionHeader(Header, "Max LTV Value", FilePath)

    idxIssuer = FindHeaderIndex(Header, "Issuer")
    idxAboveLimit = FindHeaderIndex(Header, "Position MV above limit (Haircut to Zero)")
    idxComment = FindHeaderIndex(Header, "Additional Comment")

    Set Positions = CreateObject("Scripting.Dictionary")
    Positions.CompareMode = vbTextCompare

    For r = 1 To UBound(Lines)
        If Len(Trim$(CStr(Lines(r)))) > 0 Then
            Fields = Split(Lines(r), ";")
            RowNDG = CsvText(Fields, idxNDG)

            If CLng(Val(RowNDG)) = CLng(Val(TargetNDG)) Then
                Key = PositionKey( _
                    CsvText(Fields, idxGuarantee), _
                    CsvText(Fields, idxISIN), _
                    CsvText(Fields, idxName), _
                    CsvText(Fields, idxCcy))

                If Positions.Exists(Key) Then
                    Record = Positions(Key)
                Else
                    Record = EmptyPositionRecord()
                    Record(POS_GUARANTEE) = CsvText(Fields, idxGuarantee)
                    Record(POS_ISIN) = CsvText(Fields, idxISIN)
                    Record(POS_NAME) = CsvText(Fields, idxName)
                    Record(POS_ASSET_TYPE) = CsvText(Fields, idxAssetType)
                    Record(POS_CCY) = CsvText(Fields, idxCcy)
                    Record(POS_ISSUER) = CsvText(Fields, idxIssuer)
                End If

                RowQuantity = ParseCsvDouble(CsvText(Fields, idxQuantity))
                RowPrice = ParseCsvDouble(CsvText(Fields, idxPriceEUR))
                RowValue = ParseCsvDouble(CsvText(Fields, idxValue))
                RowMaxLTV = CsvRate(CsvText(Fields, idxMaxLTV))
                QuantityWeight = Abs(CDbl(Record(POS_QUANTITY))) + Abs(RowQuantity)
                ValueWeight = Abs(CDbl(Record(POS_VALUE))) + Abs(RowValue)

                If QuantityWeight > 0 Then
                    Record(POS_PRICE_EUR) = _
                        (CDbl(Record(POS_PRICE_EUR)) * Abs(CDbl(Record(POS_QUANTITY))) + _
                         RowPrice * Abs(RowQuantity)) / QuantityWeight
                Else
                    Record(POS_PRICE_EUR) = RowPrice
                End If

                If ValueWeight > 0 Then
                    Record(POS_MAX_LTV) = _
                        (CDbl(Record(POS_MAX_LTV)) * Abs(CDbl(Record(POS_VALUE))) + _
                         RowMaxLTV * Abs(RowValue)) / ValueWeight
                Else
                    Record(POS_MAX_LTV) = RowMaxLTV
                End If

                Record(POS_QUANTITY) = CDbl(Record(POS_QUANTITY)) + RowQuantity
                Record(POS_VALUE) = CDbl(Record(POS_VALUE)) + RowValue
                Record(POS_WEIGHT) = CDbl(Record(POS_WEIGHT)) + _
                                     CsvRate(CsvText(Fields, idxWeight))
                Record(POS_ELIGIBLE_MV) = CDbl(Record(POS_ELIGIBLE_MV)) + _
                                          ParseCsvDouble(CsvText(Fields, idxEligibleMV))
                Record(POS_HCV) = CDbl(Record(POS_HCV)) + _
                                  ParseCsvDouble(CsvText(Fields, idxHCV))
                Record(POS_ABOVE_LIMIT) = CDbl(Record(POS_ABOVE_LIMIT)) + _
                                          ParseCsvDouble(CsvText(Fields, idxAboveLimit))
                Record(POS_ASSET_TYPE) = AppendUniqueText( _
                    CStr(Record(POS_ASSET_TYPE)), CsvText(Fields, idxAssetType))
                Record(POS_CCY) = AppendUniqueText( _
                    CStr(Record(POS_CCY)), CsvText(Fields, idxCcy))
                Record(POS_COMMENT) = AppendUniqueText( _
                    CStr(Record(POS_COMMENT)), CsvText(Fields, idxComment))

                If CStr(Record(POS_ISSUER)) = "" Then
                    Record(POS_ISSUER) = CsvText(Fields, idxIssuer)
                End If

                Positions(Key) = Record
            End If
        End If
    Next r

    Set ReadPositionsForNDG = Positions
End Function

Private Function RequiredPositionHeader(ByRef Header As Variant, _
                                        ByVal HeaderName As String, _
                                        ByVal FilePath As String) As Long
    RequiredPositionHeader = FindHeaderIndex(Header, HeaderName)

    If RequiredPositionHeader < 0 Then
        Err.Raise vbObjectError + 1805, "ReadPositionsForNDG", _
                  "Column '" & HeaderName & "' not found in:" & vbCrLf & FilePath
    End If
End Function

Private Function EmptyPositionRecord() As Variant
    Dim Record(0 To POS_FIELD_COUNT - 1) As Variant
    EmptyPositionRecord = Record
End Function

Private Function PositionKey(ByVal Guarantee As String, ByVal ISIN As String, _
                             ByVal SecurityName As String, _
                             ByVal PricingCcy As String) As String
    If ISIN = "" Then ISIN = SecurityName
    PositionKey = UCase$( _
        Trim$(Guarantee) & "|" & Trim$(ISIN) & "|" & Trim$(PricingCcy))
End Function

Private Function CsvText(ByRef Fields As Variant, ByVal FieldIndex As Long) As String
    Dim Text As String

    If FieldIndex < LBound(Fields) Or FieldIndex > UBound(Fields) Then Exit Function
    Text = Trim$(CStr(Fields(FieldIndex)))

    If Len(Text) >= 2 Then
        If Left$(Text, 1) = Chr$(34) And Right$(Text, 1) = Chr$(34) Then
            Text = Mid$(Text, 2, Len(Text) - 2)
            Text = Replace(Text, Chr$(34) & Chr$(34), Chr$(34))
        End If
    End If

    CsvText = Trim$(Text)
End Function

Private Function CsvRate(ByVal ValueText As String) As Double
    ValueText = Trim$(ValueText)

    If Right$(ValueText, 1) = "%" Then
        CsvRate = ParseCsvDouble(Left$(ValueText, Len(ValueText) - 1)) / 100
    Else
        CsvRate = ParseCsvDouble(ValueText)
        If Abs(CsvRate) > 1.5 Then CsvRate = CsvRate / 100
    End If
End Function

Private Sub AlignCurrencyTransitions(ByVal PreviousPositions As Object, _
                                     ByVal CurrentPositions As Object)
    Dim PreviousKey As Variant
    Dim CurrentKey As String
    Dim CurrentRecord As Variant

    For Each PreviousKey In PreviousPositions.Keys
        If Not CurrentPositions.Exists(CStr(PreviousKey)) Then
            CurrentKey = UniqueCurrencyTransitionCurrentKey( _
                CStr(PreviousKey), PreviousPositions, CurrentPositions)

            If CurrentKey <> "" Then
                CurrentRecord = CurrentPositions(CurrentKey)
                CurrentPositions.Remove CurrentKey
                CurrentPositions.Add CStr(PreviousKey), CurrentRecord
            End If
        End If
    Next PreviousKey
End Sub

Private Function UniqueCurrencyTransitionCurrentKey( _
        ByVal PreviousKey As String, _
        ByVal PreviousPositions As Object, _
        ByVal CurrentPositions As Object) As String
    Dim CandidateKey As String
    Dim CandidateCount As Long
    Dim ReverseMatchCount As Long
    Dim Key As Variant

    For Each Key In CurrentPositions.Keys
        If Not PreviousPositions.Exists(CStr(Key)) Then
            If IsCurrencyTransitionCandidate( _
                    PreviousPositions(PreviousKey), _
                    CurrentPositions(CStr(Key))) Then
                CandidateCount = CandidateCount + 1
                CandidateKey = CStr(Key)
            End If
        End If
    Next Key

    If CandidateCount <> 1 Then Exit Function

    For Each Key In PreviousPositions.Keys
        If Not CurrentPositions.Exists(CStr(Key)) Then
            If IsCurrencyTransitionCandidate( _
                    PreviousPositions(CStr(Key)), _
                    CurrentPositions(CandidateKey)) Then
                ReverseMatchCount = ReverseMatchCount + 1
            End If
        End If
    Next Key

    If ReverseMatchCount = 1 Then
        UniqueCurrencyTransitionCurrentKey = CandidateKey
    End If
End Function

Private Function IsCurrencyTransitionCandidate( _
        ByVal PreviousRecord As Variant, _
        ByVal CurrentRecord As Variant) As Boolean
    Dim PreviousISIN As String
    Dim CurrentISIN As String
    Dim PreviousCurrency As String
    Dim CurrentCurrency As String

    PreviousISIN = Trim$(CStr(PreviousRecord(POS_ISIN)))
    CurrentISIN = Trim$(CStr(CurrentRecord(POS_ISIN)))
    PreviousCurrency = Trim$(CStr(PreviousRecord(POS_CCY)))
    CurrentCurrency = Trim$(CStr(CurrentRecord(POS_CCY)))

    If PreviousISIN = "" Or CurrentISIN = "" Then Exit Function
    If PreviousCurrency = "" Or CurrentCurrency = "" Then Exit Function

    If StrComp(Trim$(CStr(PreviousRecord(POS_GUARANTEE))), _
               Trim$(CStr(CurrentRecord(POS_GUARANTEE))), _
               vbTextCompare) <> 0 Then Exit Function

    If StrComp(PreviousISIN, CurrentISIN, vbTextCompare) <> 0 Then Exit Function
    If StrComp(PreviousCurrency, CurrentCurrency, vbTextCompare) = 0 Then Exit Function

    If ValuesDiffer(CDbl(PreviousRecord(POS_QUANTITY)), _
                    CDbl(CurrentRecord(POS_QUANTITY))) Then Exit Function

    IsCurrencyTransitionCandidate = True
End Function

Private Function AppendUniqueText(ByVal ExistingText As String, _
                                  ByVal NewText As String) As String
    AppendUniqueText = ExistingText
    If NewText = "" Then Exit Function

    If ExistingText = "" Then
        AppendUniqueText = NewText
    ElseIf InStr(1, ExistingText, NewText, vbTextCompare) = 0 Then
        AppendUniqueText = ExistingText & "; " & NewText
    End If
End Function

Private Function ReadAccountMetrics(ByVal ws As Worksheet, _
                                    ByVal CurrentRow As Long) As Object
    Dim Metrics As Object
    Dim MetricNames As Variant
    Dim MetricName As Variant
    Dim ColumnNumber As Long
    Dim EventCol As Long

    If CurrentRow < 3 Then
        Err.Raise vbObjectError + 1806, "ReadAccountMetrics", _
                  "The source row has no previous account snapshot."
    End If

    Set Metrics = CreateObject("Scripting.Dictionary")
    MetricNames = Array("Approved", "MTM", "HCV", "LTV", "MC", "SF")

    For Each MetricName In MetricNames
        ColumnNumber = FindColumnByHeader(ws, CStr(MetricName))
        If ColumnNumber = 0 Then
            Err.Raise vbObjectError + 1806, "ReadAccountMetrics", _
                      "Source column '" & CStr(MetricName) & "' could not be found."
        End If

        Metrics("Previous" & CStr(MetricName)) = _
            WorksheetDouble(ws.Cells(CurrentRow - 1, ColumnNumber).Value2)
        Metrics("Current" & CStr(MetricName)) = _
            WorksheetDouble(ws.Cells(CurrentRow, ColumnNumber).Value2)
    Next MetricName

    Metrics("Event") = ""
    EventCol = FindColumnByHeader(ws, "Event")
    If EventCol > 0 Then Metrics("Event") = SafeCellText(ws.Cells(CurrentRow, EventCol))
    Set ReadAccountMetrics = Metrics
End Function

Private Sub BuildPositionChangeAnalysis(ByVal PreviousDate As Date, _
                                        ByVal CurrentDate As Date, _
                                        ByVal TargetNDG As String, _
                                        ByVal PreviousPositions As Object, _
                                        ByVal CurrentPositions As Object, _
                                        ByVal AccountMetrics As Object)
    Const DETAIL_HEADER_ROW As Long = 3
    Const DETAIL_FIRST_COL As Long = 6
    Const DETAIL_COLUMN_COUNT As Long = 39
    Const DETAIL_HCV_DELTA_COL As Long = 36

    Dim ws As Worksheet
    Dim AnalysisTable As ListObject
    Dim Keys As Object
    Dim Insights As Collection
    Dim ClassificationChanges As Collection
    Dim CurrencyChanges As Collection
    Dim Headers As Variant
    Dim Results() As Variant
    Dim PreviousRecord As Variant
    Dim CurrentRecord As Variant
    Dim Key As Variant
    Dim PreviousExists As Boolean
    Dim CurrentExists As Boolean
    Dim ChangeType As String
    Dim Drivers As String
    Dim AssetLabel As String
    Dim ChangeIdentifier As String
    Dim PreviousAssetTypeText As String
    Dim CurrentAssetTypeText As String
    Dim PreviousAssetClass As String
    Dim CurrentAssetClass As String
    Dim PreviousCurrency As String
    Dim CurrentCurrency As String
    Dim ClassificationChanged As Boolean
    Dim CurrencyChanged As Boolean
    Dim EventText As String
    Dim CommentText As String
    Dim PreviousQuantity As Double
    Dim CurrentQuantity As Double
    Dim PreviousPrice As Double
    Dim CurrentPrice As Double
    Dim PreviousValue As Double
    Dim CurrentValue As Double
    Dim PreviousWeight As Double
    Dim CurrentWeight As Double
    Dim PreviousMaxLTV As Double
    Dim CurrentMaxLTV As Double
    Dim PreviousEligible As Double
    Dim CurrentEligible As Double
    Dim PreviousHCV As Double
    Dim CurrentHCV As Double
    Dim PreviousAboveLimit As Double
    Dim CurrentAboveLimit As Double
    Dim PriceChange As Double
    Dim PositionMultiplier As Double
    Dim PriceEffect As Double
    Dim QuantityEffect As Double
    Dim ResidualEffect As Double
    Dim HCVDelta As Double
    Dim PriceHCVEffect As Double
    Dim QuantityHCVEffect As Double
    Dim CompositionHCVEffect As Double
    Dim OtherHCVEffect As Double
    Dim TotalPriceEffect As Double
    Dim TotalQuantityEffect As Double
    Dim CompositionValueEffect As Double
    Dim TotalResidualEffect As Double
    Dim PreviousPositionHCV As Double
    Dim CurrentPositionHCV As Double
    Dim TopNegativeHCV As Double
    Dim TopPositiveHCV As Double
    Dim TopNegativeAsset As String
    Dim TopPositiveAsset As String
    Dim CurrentApproved As Double
    Dim PreviousMTM As Double
    Dim CurrentMTM As Double
    Dim PreviousAccountHCV As Double
    Dim CurrentAccountHCV As Double
    Dim CurrentLTV As Double
    Dim PreviousMC As Double
    Dim CurrentMC As Double
    Dim PreviousSF As Double
    Dim CurrentSF As Double
    Dim PositionHCVDelta As Double
    Dim AccountHCVDelta As Double
    Dim ReconciliationGap As Double
    Dim HasAccountMetrics As Boolean
    Dim NewCount As Long
    Dim RemovedCount As Long
    Dim ChangedCount As Long
    Dim RowCount As Long
    Dim InsightLastRow As Long
    Dim i As Long

    AlignCurrencyTransitions PreviousPositions, CurrentPositions

    Set Keys = CreateObject("Scripting.Dictionary")
    Keys.CompareMode = vbTextCompare
    Set ClassificationChanges = New Collection
    Set CurrencyChanges = New Collection

    For Each Key In PreviousPositions.Keys
        Keys(Key) = True
    Next Key

    For Each Key In CurrentPositions.Keys
        Keys(Key) = True
    Next Key

    RowCount = Keys.Count
    If RowCount > 0 Then ReDim Results(1 To RowCount, 1 To DETAIL_COLUMN_COUNT)

    For Each Key In Keys.Keys
        i = i + 1
        PreviousExists = PreviousPositions.Exists(Key)
        CurrentExists = CurrentPositions.Exists(Key)

        If PreviousExists Then
            PreviousRecord = PreviousPositions(Key)
        Else
            PreviousRecord = EmptyPositionRecord()
        End If

        If CurrentExists Then
            CurrentRecord = CurrentPositions(Key)
        Else
            CurrentRecord = EmptyPositionRecord()
        End If

        PreviousAssetTypeText = CStr(PreviousRecord(POS_ASSET_TYPE))
        CurrentAssetTypeText = CStr(CurrentRecord(POS_ASSET_TYPE))
        PreviousAssetClass = AssetClassText(PreviousAssetTypeText)
        CurrentAssetClass = AssetClassText(CurrentAssetTypeText)
        PreviousCurrency = CStr(PreviousRecord(POS_CCY))
        CurrentCurrency = CStr(CurrentRecord(POS_CCY))
        ClassificationChanged = PreviousExists And CurrentExists And _
            TextValuesDiffer(PreviousAssetTypeText, CurrentAssetTypeText)
        CurrencyChanged = PreviousExists And CurrentExists And _
            TextValuesDiffer(PreviousCurrency, CurrentCurrency)

        PreviousQuantity = CDbl(PreviousRecord(POS_QUANTITY))
        CurrentQuantity = CDbl(CurrentRecord(POS_QUANTITY))
        PreviousPrice = CDbl(PreviousRecord(POS_PRICE_EUR))
        CurrentPrice = CDbl(CurrentRecord(POS_PRICE_EUR))
        PreviousValue = CDbl(PreviousRecord(POS_VALUE))
        CurrentValue = CDbl(CurrentRecord(POS_VALUE))
        PreviousWeight = CDbl(PreviousRecord(POS_WEIGHT))
        CurrentWeight = CDbl(CurrentRecord(POS_WEIGHT))
        PreviousMaxLTV = CDbl(PreviousRecord(POS_MAX_LTV))
        CurrentMaxLTV = CDbl(CurrentRecord(POS_MAX_LTV))
        PreviousEligible = CDbl(PreviousRecord(POS_ELIGIBLE_MV))
        CurrentEligible = CDbl(CurrentRecord(POS_ELIGIBLE_MV))
        PreviousHCV = CDbl(PreviousRecord(POS_HCV))
        CurrentHCV = CDbl(CurrentRecord(POS_HCV))
        PreviousAboveLimit = CDbl(PreviousRecord(POS_ABOVE_LIMIT))
        CurrentAboveLimit = CDbl(CurrentRecord(POS_ABOVE_LIMIT))
        HCVDelta = CurrentHCV - PreviousHCV

        If PreviousExists And CurrentExists And PreviousPrice <> 0 Then
            PriceChange = CurrentPrice / PreviousPrice - 1
        Else
            PriceChange = 0
        End If

        PriceEffect = 0
        QuantityEffect = 0

        If PreviousExists And CurrentExists And _
           PreviousQuantity <> 0 And PreviousPrice <> 0 Then
            PositionMultiplier = PreviousValue / (PreviousQuantity * PreviousPrice)
            PriceEffect = (CurrentPrice - PreviousPrice) * _
                          (PreviousQuantity + CurrentQuantity) / 2 * PositionMultiplier
            QuantityEffect = (CurrentQuantity - PreviousQuantity) * _
                             (PreviousPrice + CurrentPrice) / 2 * PositionMultiplier
        End If

        ResidualEffect = CurrentValue - PreviousValue - PriceEffect - QuantityEffect
        ChangeType = PositionChangeType( _
            PreviousExists, CurrentExists, _
            PreviousAssetTypeText, CurrentAssetTypeText, _
            PreviousCurrency, CurrentCurrency, _
            PreviousQuantity, CurrentQuantity, _
            PreviousPrice, CurrentPrice, PreviousMaxLTV, CurrentMaxLTV, _
            PreviousEligible, CurrentEligible, PreviousValue, CurrentValue, _
            PreviousHCV, CurrentHCV)
        Drivers = PositionDrivers( _
            PreviousExists, CurrentExists, _
            PreviousAssetTypeText, CurrentAssetTypeText, _
            PreviousAssetClass, CurrentAssetClass, _
            PreviousCurrency, CurrentCurrency, _
            PreviousQuantity, CurrentQuantity, _
            PreviousPrice, CurrentPrice, PreviousMaxLTV, CurrentMaxLTV, _
            PreviousEligible, CurrentEligible, HCVDelta)

        If ChangeType = "New Asset" Then
            NewCount = NewCount + 1
            CompositionHCVEffect = CompositionHCVEffect + HCVDelta
            CompositionValueEffect = CompositionValueEffect + CurrentValue - PreviousValue
        ElseIf ChangeType = "Removed Asset" Then
            RemovedCount = RemovedCount + 1
            CompositionHCVEffect = CompositionHCVEffect + HCVDelta
            CompositionValueEffect = CompositionValueEffect + CurrentValue - PreviousValue
        Else
            TotalPriceEffect = TotalPriceEffect + PriceEffect
            TotalQuantityEffect = TotalQuantityEffect + QuantityEffect
            TotalResidualEffect = TotalResidualEffect + ResidualEffect
            PriceHCVEffect = PriceHCVEffect + PriceEffect * PreviousMaxLTV
            QuantityHCVEffect = QuantityHCVEffect + QuantityEffect * PreviousMaxLTV
        End If

        If ChangeType <> "No Material Change" Then ChangedCount = ChangedCount + 1
        PreviousPositionHCV = PreviousPositionHCV + PreviousHCV
        CurrentPositionHCV = CurrentPositionHCV + CurrentHCV

        AssetLabel = PreferText(CStr(CurrentRecord(POS_NAME)), CStr(PreviousRecord(POS_NAME)))
        ChangeIdentifier = PreferText( _
            CStr(CurrentRecord(POS_ISIN)), CStr(PreviousRecord(POS_ISIN)))
        If ChangeIdentifier <> "" Then
            AssetLabel = AssetLabel & " (" & _
                         ChangeIdentifier & ")"
        Else
            ChangeIdentifier = AssetLabel
        End If

        If ClassificationChanged Then
            ClassificationChanges.Add AttributeChangeDetail( _
                ChangeIdentifier, _
                ClassificationTransitionText( _
                    PreviousAssetTypeText, CurrentAssetTypeText, _
                    PreviousAssetClass, CurrentAssetClass), _
                PreviousEligible, CurrentEligible, PreviousHCV, CurrentHCV)
        End If

        If CurrencyChanged Then
            CurrencyChanges.Add AttributeChangeDetail( _
                ChangeIdentifier, _
                DisplayText(PreviousCurrency) & " to " & _
                DisplayText(CurrentCurrency), _
                PreviousEligible, CurrentEligible, PreviousHCV, CurrentHCV)
        End If

        If HCVDelta < TopNegativeHCV Then
            TopNegativeHCV = HCVDelta
            TopNegativeAsset = AssetLabel
        End If

        If HCVDelta > TopPositiveHCV Then
            TopPositiveHCV = HCVDelta
            TopPositiveAsset = AssetLabel
        End If

        CommentText = PreferText( _
            CStr(CurrentRecord(POS_COMMENT)), CStr(PreviousRecord(POS_COMMENT)))

        Results(i, 1) = ChangeType
        Results(i, 2) = Drivers
        Results(i, 3) = PreferText(CStr(CurrentRecord(POS_GUARANTEE)), CStr(PreviousRecord(POS_GUARANTEE)))
        Results(i, 4) = PreferText(CStr(CurrentRecord(POS_ISIN)), CStr(PreviousRecord(POS_ISIN)))
        Results(i, 5) = PreferText(CStr(CurrentRecord(POS_NAME)), CStr(PreviousRecord(POS_NAME)))
        Results(i, 6) = PreviousAssetTypeText
        Results(i, 7) = CurrentAssetTypeText
        Results(i, 8) = PreviousAssetClass
        Results(i, 9) = CurrentAssetClass
        Results(i, 10) = PreferText(CStr(CurrentRecord(POS_ISSUER)), CStr(PreviousRecord(POS_ISSUER)))
        Results(i, 11) = PreviousCurrency
        Results(i, 12) = CurrentCurrency
        Results(i, 13) = PreviousQuantity
        Results(i, 14) = CurrentQuantity
        Results(i, 15) = CurrentQuantity - PreviousQuantity
        Results(i, 16) = PreviousPrice
        Results(i, 17) = CurrentPrice
        If PreviousExists And CurrentExists And PreviousPrice <> 0 Then
            Results(i, 18) = PriceChange
        Else
            Results(i, 18) = ""
        End If
        Results(i, 19) = PreviousValue
        Results(i, 20) = CurrentValue
        Results(i, 21) = CurrentValue - PreviousValue
        Results(i, 22) = PriceEffect
        Results(i, 23) = QuantityEffect
        Results(i, 24) = ResidualEffect
        Results(i, 25) = PreviousWeight
        Results(i, 26) = CurrentWeight
        Results(i, 27) = CurrentWeight - PreviousWeight
        Results(i, 28) = PreviousMaxLTV
        Results(i, 29) = CurrentMaxLTV
        Results(i, 30) = CurrentMaxLTV - PreviousMaxLTV
        Results(i, 31) = PreviousEligible
        Results(i, 32) = CurrentEligible
        Results(i, 33) = CurrentEligible - PreviousEligible
        Results(i, 34) = PreviousHCV
        Results(i, 35) = CurrentHCV
        Results(i, 36) = HCVDelta
        Results(i, 37) = PreviousAboveLimit
        Results(i, 38) = CurrentAboveLimit
        Results(i, 39) = CommentText
    Next Key

    If AccountMetrics Is Nothing Then
        HasAccountMetrics = False
    Else
        HasAccountMetrics = True
        CurrentApproved = AccountMetrics("CurrentApproved")
        PreviousMTM = AccountMetrics("PreviousMTM")
        CurrentMTM = AccountMetrics("CurrentMTM")
        PreviousAccountHCV = AccountMetrics("PreviousHCV")
        CurrentAccountHCV = AccountMetrics("CurrentHCV")
        CurrentLTV = AccountMetrics("CurrentLTV")
        PreviousMC = AccountMetrics("PreviousMC")
        CurrentMC = AccountMetrics("CurrentMC")
        PreviousSF = AccountMetrics("PreviousSF")
        CurrentSF = AccountMetrics("CurrentSF")
        EventText = AccountMetrics("Event")
    End If

    PositionHCVDelta = CurrentPositionHCV - PreviousPositionHCV
    OtherHCVEffect = PositionHCVDelta - PriceHCVEffect - _
                     QuantityHCVEffect - CompositionHCVEffect

    Set Insights = New Collection
    If HasAccountMetrics Then
        AccountHCVDelta = CurrentAccountHCV - PreviousAccountHCV
        ReconciliationGap = PositionHCVDelta - AccountHCVDelta
        AddRiskInsights _
            Insights, EventText, PreviousMC, CurrentMC, PreviousSF, CurrentSF, _
            CurrentApproved, PreviousAccountHCV, CurrentAccountHCV, _
            PreviousMTM, CurrentMTM, CurrentLTV
    End If

    If ClassificationChanges.Count > 0 Then
        Insights.Add "Asset classification changes: " & _
                     CollectionText(ClassificationChanges) & "."
    End If

    If CurrencyChanges.Count > 0 Then
        Insights.Add "Pricing currency changes: " & _
                     CollectionText(CurrencyChanges) & "."
    End If

'    Insights.Add BuildAttributionInsight( _
'        "Position MV attribution", _
'        TotalPriceEffect, TotalQuantityEffect, _
'        CompositionValueEffect, TotalResidualEffect, _
'        "residual")
'
'    Insights.Add BuildAttributionInsight( _
'        "Position-level HCV attribution", _
'        PriceHCVEffect, QuantityHCVEffect, _
'        CompositionHCVEffect, OtherHCVEffect, _
'        "LTV, eligibility, and residual factors")

'    If HasAccountMetrics And PreviousMC = 0 And CurrentMC > 0 Then
'        Insights.Add PrimaryNegativeDriver( _
'            "Main negative HCV driver", PriceHCVEffect, QuantityHCVEffect, _
'            CompositionHCVEffect, OtherHCVEffect)
'    ElseIf HasAccountMetrics And PreviousSF = 0 And CurrentSF > 0 Then
'        Insights.Add PrimaryNegativeDriver( _
'            "Main negative MTM driver", TotalPriceEffect, TotalQuantityEffect, _
'            CompositionValueEffect, TotalResidualEffect)
'    End If

    If TopNegativeAsset <> "" Then
        Insights.Add "Largest negative HCV contributor: " & TopNegativeAsset & _
                     " (" & SignedAmount(TopNegativeHCV) & ")."
    End If

    If NewCount > 0 Or RemovedCount > 0 Then
        Insights.Add AssetChangeInsight( _
            NewCount, RemovedCount, CompositionHCVEffect)
    End If

    If TopPositiveAsset <> "" Then
        Insights.Add "Largest positive HCV contributor: " & TopPositiveAsset & _
                     " (" & SignedAmount(TopPositiveHCV) & ")."
    End If

    If HasAccountMetrics And _
       ValuesDiffer(PositionHCVDelta, AccountHCVDelta, 0.000001) Then
        Insights.Add "HCV reconciliation: positions " & _
                     SignedAmount(PositionHCVDelta) & "; account " & _
                     SignedAmount(AccountHCVDelta) & "; unexplained difference " & _
                     SignedAmount(ReconciliationGap) & "."
    End If

    Set ws = CreateOrReplaceSheet(POSITION_ANALYSIS_SHEET)
    Headers = Array( _
        "Change Type", "Drivers", "CO_FT_GAR", "ISIN", "Security Name", _
        "Previous Asset Type", "Current Asset Type", _
        "Previous Asset Class", "Current Asset Class", "Issuer", _
        "Previous Currency", "Current Currency", _
        "Previous Quantity", "Current Quantity", _
        "Delta Quantity", "Previous Price EUR", "Current Price EUR", "Price Change %", _
        "Previous Position Value", "Current Position Value", "Delta Position Value", _
        "Price Effect", "Quantity Effect", "Residual Effect", _
        "Previous Weight %", "Current Weight %", "Delta Weight", _
        "Previous Max LTV %", "Current Max LTV %", "Delta Max LTV", _
        "Previous Eligible MV", "Current Eligible MV", "Delta Eligible MV", _
        "Previous HCV Contribution", "Current HCV Contribution", _
        "Delta HCV Contribution", "Previous Above Limit", "Current Above Limit", "Comment")

    ws.Cells(1, 1).Value = "Position Change Analysis"
    ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(1, DETAIL_FIRST_COL + DETAIL_COLUMN_COUNT - 1)).Merge
    ws.Cells(2, 1).Value = _
        "NDG " & TargetNDG & " | " & Format$(PreviousDate, "dd/mm/yyyy") & _
        " to " & Format$(CurrentDate, "dd/mm/yyyy") & _
        " | " & ChangedAssetText(ChangedCount)
    ws.Range( _
        ws.Cells(2, 1), _
        ws.Cells(2, DETAIL_FIRST_COL + DETAIL_COLUMN_COUNT - 1)).Merge

    WriteAccountSummary ws, PreviousDate, CurrentDate, AccountMetrics
    InsightLastRow = WriteInsights(ws, Insights)

    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(DETAIL_HEADER_ROW, DETAIL_FIRST_COL + i).Value = Headers(i)
    Next i

    If RowCount > 0 Then
        ws.Cells(DETAIL_HEADER_ROW + 1, DETAIL_FIRST_COL) _
          .Resize(RowCount, DETAIL_COLUMN_COUNT).Value = Results
        ws.Cells(DETAIL_HEADER_ROW + 1, DETAIL_FIRST_COL) _
          .Resize(RowCount, DETAIL_COLUMN_COUNT).Sort _
            Key1:=ws.Cells( _
                DETAIL_HEADER_ROW + 1, _
                DETAIL_FIRST_COL + DETAIL_HCV_DELTA_COL - 1), _
            Order1:=xlAscending, Header:=xlNo

        Set AnalysisTable = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range( _
                ws.Cells(DETAIL_HEADER_ROW, DETAIL_FIRST_COL), _
                ws.Cells( _
                    DETAIL_HEADER_ROW + RowCount, _
                    DETAIL_FIRST_COL + DETAIL_COLUMN_COUNT - 1)), _
            XlListObjectHasHeaders:=xlYes)
        AnalysisTable.name = "tblPositionChangeAnalysis"
        AnalysisTable.TableStyle = "TableStyleMedium2"

'        If ChangedCount > 0 And ChangedCount < RowCount Then
'            AnalysisTable.Range.AutoFilter Field:=1, Criteria1:="<>No Material Change"
'        End If
    Else
        ws.Cells(DETAIL_HEADER_ROW + 1, DETAIL_FIRST_COL).Value = _
            "No positions found for this NDG in either snapshot."
    End If

    FormatPositionChangeAnalysis _
        ws, DETAIL_HEADER_ROW, DETAIL_FIRST_COL, RowCount, DETAIL_COLUMN_COUNT

    AddCollateralOverviewChart _
        ws, InsightLastRow + 2, _
        DETAIL_FIRST_COL + DETAIL_COLUMN_COUNT + 2, _
        PreviousDate, CurrentDate, PreviousPositions, CurrentPositions, _
        AccountMetrics
End Sub

Private Function PositionChangeType(ByVal PreviousExists As Boolean, _
                                    ByVal CurrentExists As Boolean, _
                                    ByVal PreviousAssetType As String, _
                                    ByVal CurrentAssetType As String, _
                                    ByVal PreviousCurrency As String, _
                                    ByVal CurrentCurrency As String, _
                                    ByVal PreviousQuantity As Double, _
                                    ByVal CurrentQuantity As Double, _
                                    ByVal PreviousPrice As Double, _
                                    ByVal CurrentPrice As Double, _
                                    ByVal PreviousMaxLTV As Double, _
                                    ByVal CurrentMaxLTV As Double, _
                                    ByVal PreviousEligible As Double, _
                                    ByVal CurrentEligible As Double, _
                                    ByVal PreviousValue As Double, _
                                    ByVal CurrentValue As Double, _
                                    ByVal PreviousHCV As Double, _
                                    ByVal CurrentHCV As Double) As String
    If Not PreviousExists Then
        PositionChangeType = "New Asset"
    ElseIf Not CurrentExists Then
        PositionChangeType = "Removed Asset"
    ElseIf TextValuesDiffer(PreviousAssetType, CurrentAssetType) And _
           TextValuesDiffer(PreviousCurrency, CurrentCurrency) Then
        PositionChangeType = "Classification / Currency Change"
    ElseIf TextValuesDiffer(PreviousAssetType, CurrentAssetType) Then
        PositionChangeType = "Asset Classification Change"
    ElseIf TextValuesDiffer(PreviousCurrency, CurrentCurrency) Then
        PositionChangeType = "Currency Change"
    ElseIf ValuesDiffer(PreviousQuantity, CurrentQuantity) Then
        If CurrentQuantity > PreviousQuantity Then
            PositionChangeType = "Position Increased"
        Else
            PositionChangeType = "Position Reduced"
        End If
    ElseIf ValuesDiffer(PreviousPrice, CurrentPrice) Then
        PositionChangeType = "Price Movement"
    ElseIf ValuesDiffer(PreviousMaxLTV, CurrentMaxLTV) Or _
           ValuesDiffer(PreviousEligible, CurrentEligible) Then
        PositionChangeType = "LTV / Eligibility"
    ElseIf ValuesDiffer(PreviousValue, CurrentValue) Or _
           ValuesDiffer(PreviousHCV, CurrentHCV) Then
        PositionChangeType = "Valuation Change"
    Else
        PositionChangeType = "No Material Change"
    End If
End Function

Private Function PositionDrivers(ByVal PreviousExists As Boolean, _
                                 ByVal CurrentExists As Boolean, _
                                 ByVal PreviousAssetType As String, _
                                 ByVal CurrentAssetType As String, _
                                 ByVal PreviousAssetClass As String, _
                                 ByVal CurrentAssetClass As String, _
                                 ByVal PreviousCurrency As String, _
                                 ByVal CurrentCurrency As String, _
                                 ByVal PreviousQuantity As Double, _
                                 ByVal CurrentQuantity As Double, _
                                 ByVal PreviousPrice As Double, _
                                 ByVal CurrentPrice As Double, _
                                 ByVal PreviousMaxLTV As Double, _
                                 ByVal CurrentMaxLTV As Double, _
                                 ByVal PreviousEligible As Double, _
                                 ByVal CurrentEligible As Double, _
                                 ByVal HCVDelta As Double) As String
    Dim Drivers As String

    If Not PreviousExists Then
        PositionDrivers = "New collateral added; HCV " & SignedAmount(HCVDelta)
        Exit Function
    ElseIf Not CurrentExists Then
        PositionDrivers = "Collateral removed; HCV " & SignedAmount(HCVDelta)
        Exit Function
    End If

    If TextValuesDiffer(PreviousAssetType, CurrentAssetType) Then
        If TextValuesDiffer(PreviousAssetClass, CurrentAssetClass) Then
            AddDriver Drivers, _
                "Asset class " & DisplayText(PreviousAssetClass) & " to " & _
                DisplayText(CurrentAssetClass)
        Else
            AddDriver Drivers, _
                "Asset type " & DisplayText(PreviousAssetType) & " to " & _
                DisplayText(CurrentAssetType)
        End If
    End If

    If TextValuesDiffer(PreviousCurrency, CurrentCurrency) Then
        AddDriver Drivers, _
            "Currency " & DisplayText(PreviousCurrency) & " to " & _
            DisplayText(CurrentCurrency)
    End If

    If ValuesDiffer(PreviousPrice, CurrentPrice) And PreviousPrice <> 0 Then
        AddDriver Drivers, "Price " & SignedPercent(CurrentPrice / PreviousPrice - 1)
    End If

    If ValuesDiffer(PreviousQuantity, CurrentQuantity) Then
        AddDriver Drivers, "Quantity " & SignedAmount(CurrentQuantity - PreviousQuantity)
    End If

    If ValuesDiffer(PreviousMaxLTV, CurrentMaxLTV) Then
        AddDriver Drivers, "Max LTV " & SignedPercent(CurrentMaxLTV - PreviousMaxLTV)
    End If

    If ValuesDiffer(PreviousEligible, CurrentEligible) Then
        AddDriver Drivers, "Eligible MV " & SignedAmount(CurrentEligible - PreviousEligible)
    End If

    If ValuesDiffer(HCVDelta, 0) Then AddDriver Drivers, "HCV " & SignedAmount(HCVDelta)
    If Drivers = "" Then Drivers = "No material position-level change"
    PositionDrivers = Drivers
End Function

Private Sub AddDriver(ByRef Drivers As String, ByVal DriverText As String)
    If Drivers <> "" Then Drivers = Drivers & "; "
    Drivers = Drivers & DriverText
End Sub

Private Function TextValuesDiffer(ByVal PreviousText As String, _
                                  ByVal CurrentText As String) As Boolean
    TextValuesDiffer = StrComp( _
        Trim$(PreviousText), Trim$(CurrentText), vbTextCompare) <> 0
End Function

Private Function AssetClassText(ByVal AssetType As String) As String
    Dim AssetTypes As Variant
    Dim AssetTypeItem As Variant

    If Trim$(AssetType) = "" Then Exit Function
    AssetTypes = Split(AssetType, ";")

    For Each AssetTypeItem In AssetTypes
        AssetClassText = AppendUniqueText( _
            AssetClassText, GetAssetClass(Trim$(CStr(AssetTypeItem))))
    Next AssetTypeItem
End Function

Private Function DisplayText(ByVal ValueText As String) As String
    DisplayText = Trim$(ValueText)
    If DisplayText = "" Then DisplayText = "none"
End Function

Private Function ClassificationTransitionText( _
    ByVal PreviousAssetType As String, _
    ByVal CurrentAssetType As String, _
    ByVal PreviousAssetClass As String, _
    ByVal CurrentAssetClass As String) As String

    If TextValuesDiffer(PreviousAssetClass, CurrentAssetClass) Then
        ClassificationTransitionText = _
            DisplayText(PreviousAssetClass) & " to " & _
            DisplayText(CurrentAssetClass)
    Else
        ClassificationTransitionText = _
            DisplayText(PreviousAssetType) & " to " & _
            DisplayText(CurrentAssetType)
    End If
End Function

Private Function AttributeChangeDetail(ByVal Identifier As String, _
                                       ByVal TransitionText As String, _
                                       ByVal PreviousEligible As Double, _
                                       ByVal CurrentEligible As Double, _
                                       ByVal PreviousHCV As Double, _
                                       ByVal CurrentHCV As Double) As String
    AttributeChangeDetail = DisplayText(Identifier) & " (" & _
        TransitionText & "; " & _
        MetricMovement("eligible MV", PreviousEligible, CurrentEligible) & "; " & _
        MetricMovement("HCV", PreviousHCV, CurrentHCV) & ")"
End Function

Private Function MetricMovement(ByVal MetricName As String, _
                                ByVal PreviousValue As Double, _
                                ByVal CurrentValue As Double) As String
    Dim Difference As Double

    Difference = CurrentValue - PreviousValue
    If Not ValuesDiffer(PreviousValue, CurrentValue) Then
        MetricMovement = MetricName & " unchanged"
    ElseIf Difference > 0 Then
        MetricMovement = MetricName & " increased by " & AmountText(Difference)
    Else
        MetricMovement = MetricName & " decreased by " & AmountText(Abs(Difference))
    End If
End Function

Private Function CollectionText(ByVal Items As Collection) As String
    Dim Item As Variant

    For Each Item In Items
        If CollectionText <> "" Then CollectionText = CollectionText & "; "
        CollectionText = CollectionText & CStr(Item)
    Next Item
End Function

Private Function BuildAttributionInsight(ByVal InsightTitle As String, _
                                         ByVal PriceEffect As Double, _
                                         ByVal QuantityEffect As Double, _
                                         ByVal CompositionEffect As Double, _
                                         ByVal OtherEffect As Double, _
                                         ByVal OtherLabel As String) As String
    Dim EffectsText As String

    AddMaterialEffect EffectsText, "price", PriceEffect
    AddMaterialEffect EffectsText, "quantity", QuantityEffect
    AddMaterialEffect EffectsText, "asset additions/removals", CompositionEffect
    AddMaterialEffect EffectsText, OtherLabel, OtherEffect

    If EffectsText = "" Then
        BuildAttributionInsight = InsightTitle & ": no material effects."
    Else
        BuildAttributionInsight = InsightTitle & ": " & EffectsText & "."
    End If
End Function

Private Sub AddMaterialEffect(ByRef EffectsText As String, _
                              ByVal EffectLabel As String, _
                              ByVal EffectValue As Double)
    If Not ValuesDiffer(EffectValue, 0) Then Exit Sub
    If EffectsText <> "" Then EffectsText = EffectsText & "; "
    EffectsText = EffectsText & EffectLabel & " " & SignedAmount(EffectValue)
End Sub

Private Function AssetChangeInsight(ByVal NewCount As Long, _
                                    ByVal RemovedCount As Long, _
                                    ByVal HCVEffect As Double) As String
    Dim ChangeText As String

    If NewCount > 0 Then
        ChangeText = CStr(NewCount) & _
                     IIf(NewCount = 1, " asset added", " assets added")
    End If

    If RemovedCount > 0 Then
        If ChangeText <> "" Then ChangeText = ChangeText & "; "
        ChangeText = ChangeText & CStr(RemovedCount) & _
                     IIf(RemovedCount = 1, " asset removed", " assets removed")
    End If

    AssetChangeInsight = "Asset changes: " & ChangeText & _
                         "; net HCV impact " & SignedAmount(HCVEffect) & "."
End Function

Private Function ChangedAssetText(ByVal ChangedCount As Long) As String
    ChangedAssetText = CStr(ChangedCount) & _
                       IIf(ChangedCount = 1, " asset changed", " assets changed")
End Function

Private Function ValuesDiffer(ByVal PreviousValue As Double, _
                              ByVal CurrentValue As Double, _
                              Optional ByVal Tolerance As Double = 0.0000001) As Boolean
    Dim PreviousAbs As Double

    PreviousAbs = Abs(PreviousValue)
    If Abs(CurrentValue) > PreviousAbs Then PreviousAbs = Abs(CurrentValue)
    If PreviousAbs < 1 Then PreviousAbs = 1
    ValuesDiffer = Abs(CurrentValue - PreviousValue) > Tolerance * PreviousAbs
End Function

Private Function PreferText(ByVal CurrentText As String, ByVal PreviousText As String) As String
    If CurrentText <> "" Then PreferText = CurrentText Else PreferText = PreviousText
End Function

Private Sub AddRiskInsights(ByVal Insights As Collection, ByVal EventText As String, _
                            ByVal PreviousMC As Double, ByVal CurrentMC As Double, _
                            ByVal PreviousSF As Double, ByVal CurrentSF As Double, _
                            ByVal CurrentApproved As Double, _
                            ByVal PreviousHCV As Double, ByVal CurrentHCV As Double, _
                            ByVal PreviousMTM As Double, ByVal CurrentMTM As Double, _
                            ByVal CurrentLTV As Double)
    Dim OtherEvents As String

    OtherEvents = NonRiskEventText(EventText)
    If OtherEvents <> "" Then Insights.Add "Account events: " & OtherEvents & "."

    If PreviousMC = 0 And CurrentMC > 0 Then
        Insights.Add "Margin call triggered: HCV fell from " & _
                     AmountText(PreviousHCV) & " to " & AmountText(CurrentHCV) & _
                     "; approved limit " & AmountText(CurrentApproved) & _
                     "; deficit " & AmountText(CurrentApproved - CurrentHCV) & "."
    ElseIf PreviousMC > 0 And CurrentMC = 0 Then
        Insights.Add "Margin call cleared: HCV exceeds the approved limit by " & _
                     AmountText(CurrentHCV - CurrentApproved) & "."
    ElseIf CurrentMC > 0 Then
        Insights.Add "Margin call remains active: HCV is " & _
                     AmountText(CurrentApproved - CurrentHCV) & _
                     " below the approved limit."
    End If

    If PreviousSF = 0 And CurrentSF > 0 Then
        If Not HasMeaningfulMTM(CurrentMTM) Then
            Insights.Add "Shortfall triggered: MTM fell from " & _
                         AmountText(PreviousMTM) & _
                         " to an effectively zero balance; approved limit " & _
                         AmountText(CurrentApproved) & _
                         "; LTV is not meaningful."
        Else
            Insights.Add "Shortfall triggered: MTM fell from " & _
                         AmountText(PreviousMTM) & " to " & AmountText(CurrentMTM) & _
                         "; approved limit " & AmountText(CurrentApproved) & _
                         "; LTV " & Format$(CurrentLTV, "0.0%") & "."
        End If
    ElseIf PreviousSF > 0 And CurrentSF = 0 Then
        Insights.Add "Shortfall cleared: MTM exceeds the approved limit by " & _
                     AmountText(CurrentMTM - CurrentApproved) & "."
    ElseIf CurrentSF > 0 Then
        If Not HasMeaningfulMTM(CurrentMTM) Then
            Insights.Add "Shortfall remains active: MTM is effectively zero; " & _
                         "approved limit " & AmountText(CurrentApproved) & _
                         "; LTV is not meaningful."
        ElseIf CurrentMTM < CurrentApproved Then
            Insights.Add "Shortfall remains active: MTM is " & _
                         AmountText(CurrentApproved - CurrentMTM) & _
                         " below the approved limit."
        Else
            Insights.Add "Shortfall remains active despite MTM headroom of " & _
                         AmountText(CurrentMTM - CurrentApproved) & "."
        End If
    End If
End Sub

Private Function HasMeaningfulMTM(ByVal MTMValue As Double) As Boolean
    HasMeaningfulMTM = MTMValue >= MIN_MEANINGFUL_MTM
End Function

Private Function NonRiskEventText(ByVal EventText As String) As String
    Dim Events As Variant
    Dim EventItem As Variant
    Dim CleanItem As String

    Events = Split(EventText, ";")
    For Each EventItem In Events
        CleanItem = Trim$(CStr(EventItem))

        Select Case LCase$(CleanItem)
            Case "", "margin call triggered", "margin call cleared", _
                 "shortfall triggered", "shortfall cleared"
            Case Else
                If NonRiskEventText <> "" Then _
                    NonRiskEventText = NonRiskEventText & "; "
                NonRiskEventText = NonRiskEventText & CleanItem
        End Select
    Next EventItem
End Function

Private Function AmountText(ByVal Value As Double) As String
    AmountText = Format$(Value, "#,##0.00;-#,##0.00;0.00")
End Function

Private Function SignedAmount(ByVal Value As Double) As String
    SignedAmount = Format$(Value, "+#,##0.00;-#,##0.00;0.00")
End Function

Private Function SignedPercent(ByVal Value As Double) As String
    SignedPercent = Format$(Value, "+0.0%;-0.0%;0.0%")
End Function

Private Function PrimaryNegativeDriver(ByVal InsightTitle As String, _
                                       ByVal PriceEffect As Double, _
                                       ByVal QuantityEffect As Double, _
                                       ByVal CompositionEffect As Double, _
                                       ByVal OtherEffect As Double) As String
    Dim DriverName As String
    Dim DriverValue As Double

    If PriceEffect < DriverValue Then
        DriverName = "price"
        DriverValue = PriceEffect
    End If

    If QuantityEffect < DriverValue Then
        DriverName = "quantity"
        DriverValue = QuantityEffect
    End If

    If CompositionEffect < DriverValue Then
        DriverName = "asset additions/removals"
        DriverValue = CompositionEffect
    End If

    If OtherEffect < DriverValue Then
        DriverName = "LTV, eligibility, or residual factors"
        DriverValue = OtherEffect
    End If

    If DriverName = "" Then
        PrimaryNegativeDriver = _
            InsightTitle & ": no material negative driver identified."
    Else
        PrimaryNegativeDriver = InsightTitle & ": " & DriverName & _
                                " (" & SignedAmount(DriverValue) & ")."
    End If
End Function

Private Sub WriteAccountSummary(ByVal ws As Worksheet, _
                                ByVal PreviousDate As Date, _
                                ByVal CurrentDate As Date, _
                                ByVal Metrics As Object)
    ws.Range("A3").Value = "Account Metric"
    ws.Range("B3").Value = PreviousDate
    ws.Range("C3").Value = CurrentDate
    ws.Range("D3").Value = "Change"

    If Metrics Is Nothing Then
        ws.Range("A4:D4").Merge
        ws.Range("A4").Value = "No account-level source supplied."
        Exit Sub
    End If

    ws.Range("A4:A9").Value = Application.Transpose( _
        Array("Approved", "MTM", "HCV", "LTV", "Margin Call", "Shortfall"))

    ws.Range("B4").Value = Metrics("PreviousApproved")
    ws.Range("C4").Value = Metrics("CurrentApproved")
    ws.Range("D4").Value = Metrics("CurrentApproved") - Metrics("PreviousApproved")
    ws.Range("B5").Value = Metrics("PreviousMTM")
    ws.Range("C5").Value = Metrics("CurrentMTM")
    ws.Range("D5").Value = Metrics("CurrentMTM") - Metrics("PreviousMTM")
    ws.Range("B6").Value = Metrics("PreviousHCV")
    ws.Range("C6").Value = Metrics("CurrentHCV")
    ws.Range("D6").Value = Metrics("CurrentHCV") - Metrics("PreviousHCV")
    If HasMeaningfulMTM(CDbl(Metrics("PreviousMTM"))) Then
        ws.Range("B7").Value = Metrics("PreviousLTV")
    Else
        ws.Range("B7").Value = LTV_NOT_MEANINGFUL
    End If

    If HasMeaningfulMTM(CDbl(Metrics("CurrentMTM"))) Then
        ws.Range("C7").Value = Metrics("CurrentLTV")
    Else
        ws.Range("C7").Value = LTV_NOT_MEANINGFUL
    End If

    If IsNumeric(ws.Range("B7").Value) And IsNumeric(ws.Range("C7").Value) Then
        ws.Range("D7").Value = Metrics("CurrentLTV") - Metrics("PreviousLTV")
    Else
        ws.Range("D7").Value = LTV_NOT_MEANINGFUL
    End If
    ws.Range("B8").Value = IIf(Metrics("PreviousMC") > 0, "Yes", "No")
    ws.Range("C8").Value = IIf(Metrics("CurrentMC") > 0, "Yes", "No")
    ws.Range("B9").Value = IIf(Metrics("PreviousSF") > 0, "Yes", "No")
    ws.Range("C9").Value = IIf(Metrics("CurrentSF") > 0, "Yes", "No")

    ApplyUnifiedDeltaHighlight ws.Range("D4"), CDbl(ws.Range("D4").Value2)
    ApplyUnifiedDeltaHighlight ws.Range("D5"), CDbl(ws.Range("D5").Value2)
    ApplyUnifiedDeltaHighlight ws.Range("D6"), CDbl(ws.Range("D6").Value2)
    ApplyUnifiedStatusHighlight ws.Range("B8"), _
        IIf(Metrics("PreviousMC") > 0, "Margin Call", "Normal")
    ApplyUnifiedStatusHighlight ws.Range("C8"), _
        IIf(Metrics("CurrentMC") > 0, "Margin Call", "Normal")
    ApplyUnifiedStatusHighlight ws.Range("B9"), _
        IIf(Metrics("PreviousSF") > 0, "Shortfall", "Normal")
    ApplyUnifiedStatusHighlight ws.Range("C9"), _
        IIf(Metrics("CurrentSF") > 0, "Shortfall", "Normal")
End Sub

Private Function WriteInsights(ByVal ws As Worksheet, _
                               ByVal Insights As Collection) As Long
    Const FIRST_INSIGHT_ROW As Long = 12
    Const INSIGHT_CHARACTERS_PER_ROW As Long = 80
    Const INSIGHT_ROW_HEIGHT As Double = 18
    Const INSIGHT_SHARED_PADDING_ROWS As Long = 1

    Dim InsightRange As Range
    Dim ItemText As String
    Dim InsightText As String
    Dim RequiredRows As Long
    Dim i As Long

    ws.Range("A11:D11").Merge
    ws.Range("A11").Value = "Attribution Insights"

    For i = 1 To Insights.Count
        ItemText = Chr$(149) & " " & CStr(Insights(i))
        If InsightText <> "" Then InsightText = InsightText & vbLf
        InsightText = InsightText & ItemText
        RequiredRows = RequiredRows + EstimatedInsightRows( _
            ItemText, INSIGHT_CHARACTERS_PER_ROW)
    Next i

    If RequiredRows = 0 Then
        InsightText = "No material insights identified."
        RequiredRows = 1
    End If
    RequiredRows = RequiredRows + INSIGHT_SHARED_PADDING_ROWS

    Set InsightRange = ws.Range( _
        ws.Cells(FIRST_INSIGHT_ROW, 1), _
        ws.Cells(FIRST_INSIGHT_ROW + RequiredRows - 1, 4))
    InsightRange.Merge
    InsightRange.Value = InsightText
    InsightRange.WrapText = True
    InsightRange.HorizontalAlignment = xlLeft
    InsightRange.VerticalAlignment = xlTop
    InsightRange.Rows.RowHeight = INSIGHT_ROW_HEIGHT
    WriteInsights = FIRST_INSIGHT_ROW + RequiredRows - 1
End Function

Private Function EstimatedInsightRows(ByVal InsightText As String, _
                                      ByVal CharactersPerRow As Long) As Long
    EstimatedInsightRows = _
        (Len(InsightText) + CharactersPerRow - 1) \ CharactersPerRow
    If EstimatedInsightRows < 1 Then EstimatedInsightRows = 1
End Function

Private Sub AddCollateralOverviewChart(ByVal ws As Worksheet, _
                                       ByVal ChartTopRow As Long, _
                                       ByVal HelperFirstColumn As Long, _
                                       ByVal PreviousDate As Date, _
                                       ByVal CurrentDate As Date, _
                                       ByVal PreviousPositions As Object, _
                                       ByVal CurrentPositions As Object, _
                                       ByVal AccountMetrics As Object)
    Const CHART_HEIGHT As Double = 230

    Dim PreviousMV As Object
    Dim CurrentMV As Object
    Dim AssetClasses As Collection
    Dim IncludedClasses As Object
    Dim ClassOrder As Variant
    Dim ClassName As Variant
    Dim ChartBox As ChartObject
    Dim OverviewChart As Chart
    Dim ChartSeries As Series
    Dim CategoryRange As Range
    Dim HelperColumn As Long
    Dim HelperLastColumn As Long
    Dim PreviousAmount As Double
    Dim CurrentAmount As Double

    Set PreviousMV = AssetClassMVTotals(PreviousPositions)
    Set CurrentMV = AssetClassMVTotals(CurrentPositions)
    Set AssetClasses = New Collection
    Set IncludedClasses = CreateObject("Scripting.Dictionary")
    IncludedClasses.CompareMode = vbTextCompare

    ClassOrder = Array( _
        "Cash", "Bonds", "Equity", "Funds", "Certificates", _
        "Insurance", "GP", "Non Eligible Asset", "UNKNOWN")

    For Each ClassName In ClassOrder
        PreviousAmount = DictionaryAmount(PreviousMV, CStr(ClassName))
        CurrentAmount = DictionaryAmount(CurrentMV, CStr(ClassName))
        If ValuesDiffer(PreviousAmount, 0) Or _
           ValuesDiffer(CurrentAmount, 0) Then
            AssetClasses.Add CStr(ClassName)
            IncludedClasses(CStr(ClassName)) = True
        End If
    Next ClassName

    For Each ClassName In PreviousMV.Keys
        If Not IncludedClasses.Exists(CStr(ClassName)) Then
            AssetClasses.Add CStr(ClassName)
            IncludedClasses(CStr(ClassName)) = True
        End If
    Next ClassName

    For Each ClassName In CurrentMV.Keys
        If Not IncludedClasses.Exists(CStr(ClassName)) Then
            AssetClasses.Add CStr(ClassName)
            IncludedClasses(CStr(ClassName)) = True
        End If
    Next ClassName

    ws.Cells(1, HelperFirstColumn).Value = "Snapshot Date"
    ws.Cells(2, HelperFirstColumn).Value = PreviousDate
    ws.Cells(3, HelperFirstColumn).Value = CurrentDate
    ws.Range(ws.Cells(2, HelperFirstColumn), _
             ws.Cells(3, HelperFirstColumn)).NumberFormat = "dd/mm/yyyy"
    Set CategoryRange = ws.Range( _
        ws.Cells(2, HelperFirstColumn), ws.Cells(3, HelperFirstColumn))

    Set ChartBox = ws.ChartObjects.Add( _
        Left:=ws.Cells(ChartTopRow, 1).Left, _
        Top:=ws.Cells(ChartTopRow, 1).Top, _
        Width:=ws.Range(ws.Cells(ChartTopRow, 1), _
                        ws.Cells(ChartTopRow, 4)).Width, _
        Height:=CHART_HEIGHT)
    ChartBox.name = "chtCollateralOverview"
    ChartBox.Placement = xlMoveAndSize
    Set OverviewChart = ChartBox.Chart

    With OverviewChart
        .ChartType = xlAreaStacked
        .HasTitle = True
        .ChartTitle.Text = "Collateral Mix & Coverage"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .PlotVisibleOnly = False
        .DisplayBlanksAs = xlNotPlotted
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    HelperColumn = HelperFirstColumn + 1
    For Each ClassName In AssetClasses
        ws.Cells(1, HelperColumn).Value = CStr(ClassName)
        ws.Cells(2, HelperColumn).Value = _
            DictionaryAmount(PreviousMV, CStr(ClassName))
        ws.Cells(3, HelperColumn).Value = _
            DictionaryAmount(CurrentMV, CStr(ClassName))

        Set ChartSeries = OverviewChart.SeriesCollection.NewSeries
        With ChartSeries
            .name = CStr(ClassName)
            .XValues = CategoryRange
            .Values = ws.Range( _
                ws.Cells(2, HelperColumn), ws.Cells(3, HelperColumn))
            .ChartType = xlAreaStacked
            .AxisGroup = xlPrimary
            .Format.Fill.ForeColor.RGB = AssetClassChartColor(CStr(ClassName))
            .Format.Fill.Transparency = 0.08
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            .Format.Line.Transparency = 0.35
            .Format.Line.Weight = 0.75
        End With

        HelperColumn = HelperColumn + 1
    Next ClassName

    If Not AccountMetrics Is Nothing Then
        ws.Cells(1, HelperColumn).Value = "HCV"
        ws.Cells(2, HelperColumn).Value = AccountMetrics("PreviousHCV")
        ws.Cells(3, HelperColumn).Value = AccountMetrics("CurrentHCV")
        AddOverviewLineSeries _
            OverviewChart, "HCV", CategoryRange, _
            ws.Range(ws.Cells(2, HelperColumn), ws.Cells(3, HelperColumn)), _
            RGB(31, 78, 121), xlMarkerStyleCircle
        HelperColumn = HelperColumn + 1

        ws.Cells(1, HelperColumn).Value = "Approved Loan"
        ws.Cells(2, HelperColumn).Value = AccountMetrics("PreviousApproved")
        ws.Cells(3, HelperColumn).Value = AccountMetrics("CurrentApproved")
        AddOverviewLineSeries _
            OverviewChart, "Approved Loan", CategoryRange, _
            ws.Range(ws.Cells(2, HelperColumn), ws.Cells(3, HelperColumn)), _
            RGB(156, 0, 6), xlMarkerStyleDiamond
        HelperColumn = HelperColumn + 1
    End If

    HelperLastColumn = HelperColumn - 1
    With OverviewChart
        On Error Resume Next
        .Axes(xlCategory).TickLabels.NumberFormat = "dd/mm/yyyy"
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0.0,,""m"""
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = _
            RGB(225, 225, 225)
        On Error GoTo 0
    End With

    ws.Range(ws.Cells(1, HelperFirstColumn), _
             ws.Cells(1, HelperLastColumn)).EntireColumn.Hidden = True
    ws.Range("A1").Select
End Sub

Private Function AssetClassMVTotals(ByVal Positions As Object) As Object
    Dim Totals As Object
    Dim PositionKey As Variant
    Dim PositionRecord As Variant
    Dim ClassName As String

    Set Totals = CreateObject("Scripting.Dictionary")
    Totals.CompareMode = vbTextCompare

    If Positions Is Nothing Then
        Set AssetClassMVTotals = Totals
        Exit Function
    End If

    For Each PositionKey In Positions.Keys
        PositionRecord = Positions(PositionKey)
        ClassName = AssetClassText(CStr(PositionRecord(POS_ASSET_TYPE)))
        If ClassName = "" Then ClassName = "UNKNOWN"

        If Totals.Exists(ClassName) Then
            Totals(ClassName) = CDbl(Totals(ClassName)) + _
                                CDbl(PositionRecord(POS_VALUE))
        Else
            Totals.Add ClassName, CDbl(PositionRecord(POS_VALUE))
        End If
    Next PositionKey

    Set AssetClassMVTotals = Totals
End Function

Private Function DictionaryAmount(ByVal Amounts As Object, _
                                  ByVal ItemName As String) As Double
    If Amounts.Exists(ItemName) Then _
        DictionaryAmount = CDbl(Amounts(ItemName))
End Function

Private Sub AddOverviewLineSeries(ByVal TargetChart As Chart, _
                                  ByVal SeriesName As String, _
                                  ByVal CategoryRange As Range, _
                                  ByVal ValueRange As Range, _
                                  ByVal LineColor As Long, _
                                  ByVal MarkerStyle As XlMarkerStyle)
    Dim ChartSeries As Series

    Set ChartSeries = TargetChart.SeriesCollection.NewSeries
    With ChartSeries
        .name = SeriesName
        .XValues = CategoryRange
        .Values = ValueRange
        .ChartType = xlLineMarkers
        .AxisGroup = xlPrimary
        .MarkerStyle = MarkerStyle
        .MarkerSize = 7
        .MarkerForegroundColor = LineColor
        .MarkerBackgroundColor = RGB(255, 255, 255)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = LineColor
        .Format.Line.Weight = 2.25
    End With
End Sub

Private Function AssetClassChartColor(ByVal ClassName As String) As Long
    Select Case ClassName
        Case "Cash"
            AssetClassChartColor = RGB(91, 155, 213)
        Case "Bonds"
            AssetClassChartColor = RGB(112, 173, 71)
        Case "Equity"
            AssetClassChartColor = RGB(237, 125, 49)
        Case "Funds"
            AssetClassChartColor = RGB(165, 165, 165)
        Case "Certificates"
            AssetClassChartColor = RGB(255, 192, 0)
        Case "Insurance"
            AssetClassChartColor = RGB(112, 48, 160)
        Case "GP"
            AssetClassChartColor = RGB(0, 176, 240)
        Case "Non Eligible Asset"
            AssetClassChartColor = RGB(192, 80, 77)
        Case Else
            AssetClassChartColor = RGB(128, 128, 128)
    End Select
End Function

Private Sub ApplyPositiveFontColor(ByVal TargetRange As Range)
    Dim TargetCell As Range

    For Each TargetCell In TargetRange.Cells
        If IsNumeric(TargetCell.Value2) Then
            If CDbl(TargetCell.Value2) > 0 Then
                TargetCell.Font.Color = RGB(84, 130, 53)
            End If
        End If
    Next TargetCell
End Sub

Private Sub FormatPositionChangeAnalysis(ByVal ws As Worksheet, _
                                         ByVal DetailHeaderRow As Long, _
                                         ByVal DetailFirstColumn As Long, _
                                         ByVal RowCount As Long, _
                                         ByVal DetailColumnCount As Long)
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim r As Long

    LastRow = DetailHeaderRow + RowCount
    LastColumn = DetailFirstColumn + DetailColumnCount - 1
    ws.Cells.Font.name = "Arial"

    ApplyUnifiedReportTitle _
        ws.Range(ws.Cells(1, 1), ws.Cells(1, LastColumn))
    ApplyUnifiedReportSubtitle _
        ws.Range(ws.Cells(2, 1), ws.Cells(2, LastColumn))
    ApplyUnifiedSectionTitle ws.Range("A3:D3")
    ApplyUnifiedSectionTitle ws.Range("A11:D11")

    ws.Range("B3:C3").NumberFormat = "dd/mm/yyyy"
    ws.Range("B4:D6").NumberFormat = "#,##0.00;[Red]-#,##0.00;-"
    ws.Range("B7:D7").NumberFormat = "+0.0%;[Red]-0.0%;-"

    If RowCount > 0 Then
        ws.Rows(CStr(DetailHeaderRow + 1) & ":" & CStr(LastRow)).RowHeight = 18
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 12), _
                 ws.Cells(LastRow, DetailFirstColumn + 13)).NumberFormat = _
                    "#,##0.0000;[Red]-#,##0.0000;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 14), _
                 ws.Cells(LastRow, DetailFirstColumn + 14)).NumberFormat = _
                    "+#,##0.0000;[Red]-#,##0.0000;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 15), _
                 ws.Cells(LastRow, DetailFirstColumn + 16)).NumberFormat = _
                    "#,##0.0000;[Red]-#,##0.0000;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 17), _
                 ws.Cells(LastRow, DetailFirstColumn + 17)).NumberFormat = _
                    "+0.0%;[Red]-0.0%;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 18), _
                 ws.Cells(LastRow, DetailFirstColumn + 19)).NumberFormat = _
                    "#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 20), _
                 ws.Cells(LastRow, DetailFirstColumn + 23)).NumberFormat = _
                    "+#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 24), _
                 ws.Cells(LastRow, DetailFirstColumn + 25)).NumberFormat = _
                    "0.0%;[Red]-0.0%;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 26), _
                 ws.Cells(LastRow, DetailFirstColumn + 26)).NumberFormat = _
                    "+0.0%;[Red]-0.0%;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 27), _
                 ws.Cells(LastRow, DetailFirstColumn + 28)).NumberFormat = _
                    "0.0%;[Red]-0.0%;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 29), _
                 ws.Cells(LastRow, DetailFirstColumn + 29)).NumberFormat = _
                    "+0.0%;[Red]-0.0%;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 30), _
                 ws.Cells(LastRow, DetailFirstColumn + 31)).NumberFormat = _
                    "#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 32), _
                 ws.Cells(LastRow, DetailFirstColumn + 32)).NumberFormat = _
                    "+#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 33), _
                 ws.Cells(LastRow, DetailFirstColumn + 34)).NumberFormat = _
                    "#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 35), _
                 ws.Cells(LastRow, DetailFirstColumn + 35)).NumberFormat = _
                    "+#,##0.00;[Red]-#,##0.00;-"
        ws.Range(ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 36), _
                 ws.Cells(LastRow, DetailFirstColumn + 37)).NumberFormat = _
                    "#,##0.00;[Red]-#,##0.00;-"

        ApplyPositiveFontColor ws.Range( _
            ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 14), _
            ws.Cells(LastRow, DetailFirstColumn + 14))
        ApplyPositiveFontColor ws.Range( _
            ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 20), _
            ws.Cells(LastRow, DetailFirstColumn + 23))
        ApplyPositiveFontColor ws.Range( _
            ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 32), _
            ws.Cells(LastRow, DetailFirstColumn + 32))
        ApplyPositiveFontColor ws.Range( _
            ws.Cells(DetailHeaderRow + 1, DetailFirstColumn + 35), _
            ws.Cells(LastRow, DetailFirstColumn + 35))

        For r = DetailHeaderRow + 1 To LastRow
            Select Case CStr(ws.Cells(r, DetailFirstColumn).Value)
                Case "New Asset"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(198, 239, 206)
                Case "Removed Asset"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(255, 199, 206)
                Case "Position Increased"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(221, 235, 247)
                Case "Position Reduced"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(255, 235, 156)
                Case "LTV / Eligibility"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(244, 176, 132)
                Case "Asset Classification Change"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(248, 203, 173)
                Case "Currency Change"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(189, 215, 238)
                Case "Classification / Currency Change"
                    ws.Cells(r, DetailFirstColumn).Interior.Color = RGB(217, 210, 233)
                Case "No Material Change"
                    ws.Cells(r, DetailFirstColumn).Font.Color = RGB(128, 128, 128)
            End Select
        Next r
    End If

    ws.Columns.AutoFit
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 16
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 3
    ws.Columns(DetailFirstColumn).ColumnWidth = 20
    ws.Columns(DetailFirstColumn + 1).ColumnWidth = 42
    ws.Columns(DetailFirstColumn + 4).ColumnWidth = 30
    ws.Columns(DetailFirstColumn + 5).ColumnWidth = 28
    ws.Columns(DetailFirstColumn + 6).ColumnWidth = 28
    ws.Columns(DetailFirstColumn + 7).ColumnWidth = 18
    ws.Columns(DetailFirstColumn + 8).ColumnWidth = 18
    ws.Columns(DetailFirstColumn + 9).ColumnWidth = 24
    ws.Columns(DetailFirstColumn + 10).ColumnWidth = 12
    ws.Columns(DetailFirstColumn + 11).ColumnWidth = 12
    ws.Columns(LastColumn).ColumnWidth = 35
    ws.Range( _
        ws.Cells(DetailHeaderRow, DetailFirstColumn), _
        ws.Cells(LastRow, LastColumn)).VerticalAlignment = xlCenter

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(DetailHeaderRow + 1, DetailFirstColumn).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub




