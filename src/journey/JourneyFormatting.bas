Attribute VB_Name = "JourneyFormatting"
Option Explicit

Public Sub FormatJourneyTable()

    On Error GoTo ExitHandler

    Application.ScreenUpdating = False

    Dim ws As Worksheet

    Dim LastRow As Long
    Dim LastCol As Long

    Dim ApprovedCol As Long
    Dim MCCol As Long
    Dim SFCol As Long
    Dim DeltaApprovedCol As Long
    Dim DeltaDrawnCol As Long
    Dim LTVCol As Long
    Dim DeltaLTVCol As Long
    Dim MCClearedCol As Long
    Dim EventCol As Long
    Dim ReasonCol As Long
    Dim CommentCol As Long
    
    Dim TableRange As Range

    Set ws = _
        Worksheets("NDG Journey")

    LastRow = _
        GetLastRow(ws, "A")

    LastCol = _
        ws.Cells(1, ws.Columns.Count) _
          .End(xlToLeft).Column

    ApprovedCol = _
        FindColumnByHeader( _
            ws, _
            "Approved")

    MCCol = _
        FindColumnByHeader( _
            ws, _
            "MC")

    SFCol = _
        FindColumnByHeader( _
            ws, _
            "SF")
            
    DeltaApprovedCol = _
        FindColumnByHeader( _
            ws, _
            "Delta Approved")
    
    DeltaDrawnCol = _
        FindColumnByHeader( _
            ws, _
            "Delta Drawn")
            
    LTVCol = _
        FindColumnByHeader( _
            ws, _
            "LTV")
            
    DeltaLTVCol = _
        FindColumnByHeader( _
            ws, _
            "Delta LTV")

    MCClearedCol = _
        FindColumnByHeader( _
            ws, _
            "MC Cleared")

    EventCol = _
        FindColumnByHeader( _
            ws, _
            "Event")

    ReasonCol = _
        FindColumnByHeader( _
            ws, _
            "Reason MC/SF")

    CommentCol = _
        FindColumnByHeader( _
            ws, _
            "Comment")

    If ApprovedCol = 0 _
       Or MCCol = 0 _
       Or SFCol = 0 _
       Or LTVCol = 0 _
       Or DeltaLTVCol = 0 _
       Or MCClearedCol = 0 _
       Or EventCol = 0 _
       Or ReasonCol = 0 _
       Or CommentCol = 0 _
       Or DeltaApprovedCol = 0 _
       Or DeltaDrawnCol = 0 Then
    
        Err.Raise _
            vbObjectError + 1400, _
            "FormatJourneyTable", _
            "One or more required Journey columns could not be found."
    
    End If
    
    Set TableRange = _
        ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(LastRow, LastCol))

    ClearJourneyBandFormatting _
        ws, _
        2, _
        LastRow, _
        1, _
        LastCol
        
    FormatReportTable _
        TableRange, _
        HeaderRows:=1
        
    FormatReportTable _
        TableRange, _
        HeaderRows:=1
    
    FormatJourneyColumnSeparators _
        ws, _
        LastRow, _
        LastCol
        
'    With TableRange
'        .HorizontalAlignment = xlRight
'        .IndentLevel = 1
'    End With
    
    With ws.Columns(EventCol)
'        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    With ws.Columns(ReasonCol)
'        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    
    With ws.Columns(CommentCol)
'        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
        
    '
    ' AutoFit the whole table first.
    '
    TableRange.Columns.AutoFit

    '
    ' Entire-row MC formatting.
    '
    FormatMCJourneyBands _
        ws, _
        LastRow, _
        LastCol, _
        MCCol, _
        MCClearedCol, _
        ReasonCol

    '
    ' Entire-row SF formatting.
    '
    FormatSFJourneyBands _
        ws, _
        LastRow, _
        LastCol, _
        SFCol, _
        ReasonCol

    '
    ' Positive MC cells receive an additional
    ' fill based on MC / Approved.
    '
    FormatMCByApprovedRatio _
        ws, _
        LastRow, _
        MCCol, _
        ApprovedCol
        
    FormatLTVTrendCells _
        ws, _
        LastRow, _
        LTVCol, _
        DeltaLTVCol
        
    HighlightNonZeroDeltaCells _
        ws, _
        LastRow, _
        DeltaApprovedCol, _
        DeltaDrawnCol

'    With ws.Columns(ReasonCol)
'
'        .ColumnWidth = 22
'        .WrapText = True
'        .VerticalAlignment = xlTop
'
'    End With

'    With ws.Columns(CommentCol)
'
'        .ColumnWidth = 45
'        .WrapText = True
'        .VerticalAlignment = xlTop
'
'    End With

'    ws.Rows("2:" & LastRow).AutoFit

'    FormatJourneyColumns

ExitHandler:

    Application.ScreenUpdating = True

    If Err.Number <> 0 Then

        MsgBox _
            "Error " & Err.Number & vbCrLf & _
            Err.Description, _
            vbExclamation, _
            "Format Journey Table"

    End If

End Sub

Private Sub FormatMCJourneyBands( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal LastCol As Long, _
    ByVal MCCol As Long, _
    ByVal MCClearedCol As Long, _
    ByVal ReasonCol As Long)

    Dim StartRow As Long
    Dim EndRow As Long
    Dim LastActiveRow As Long

    Dim InBand As Boolean
    Dim IsTechnical As Boolean

    Dim ReasonText As String

    Dim r As Long

    InBand = False
    StartRow = 0

    For r = 2 To LastRow

        '
        ' Start an MC interval.
        '
        If Not InBand Then

            If GetNumericValue( _
                    ws.Cells(r, MCCol).Value2) > 0 Then

                StartRow = r
                InBand = True

            End If

        End If

        '
        ' MC Cleared row itself is not active and
        ' therefore is not filled.
        '
        If InBand Then

            If GetNumericValue( _
                    ws.Cells(r, MCClearedCol).Value2) = 1 Then

                EndRow = r
                LastActiveRow = EndRow - 1

                If LastActiveRow < StartRow Then
                    LastActiveRow = StartRow
                End If

                ReasonText = _
                    SafeCellText( _
                        ws.Cells( _
                            LastActiveRow, _
                            ReasonCol))

                IsTechnical = _
                    IsTechnicalReason(ReasonText)

                ApplyJourneyBandFormatting _
                    ws, _
                    StartRow, _
                    LastActiveRow, _
                    LastCol, _
                    IsTechnical, _
                    False

                InBand = False
                StartRow = 0

            End If

        End If

    Next r

    '
    ' MC still open at the end.
    '
    If InBand Then

        LastActiveRow = LastRow

        ReasonText = _
            SafeCellText( _
                ws.Cells( _
                    LastActiveRow, _
                    ReasonCol))

        IsTechnical = _
            IsTechnicalReason(ReasonText)

        ApplyJourneyBandFormatting _
            ws, _
            StartRow, _
            LastActiveRow, _
            LastCol, _
            IsTechnical, _
            False

    End If

End Sub

Private Sub FormatSFJourneyBands( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal LastCol As Long, _
    ByVal SFCol As Long, _
    ByVal ReasonCol As Long)

    Dim StartRow As Long
    Dim LastActiveRow As Long

    Dim InBand As Boolean
    Dim IsTechnical As Boolean

    Dim ReasonText As String

    Dim r As Long

    InBand = False
    StartRow = 0

    For r = 2 To LastRow

        If Not InBand Then

            If GetNumericValue( _
                    ws.Cells(r, SFCol).Value2) > 0 Then

                StartRow = r
                InBand = True

            End If

        Else

            '
            ' First SF = 0 row is the cleared row.
            ' The cleared row itself is not formatted.
            '
            If GetNumericValue( _
                    ws.Cells(r, SFCol).Value2) = 0 Then

                LastActiveRow = r - 1

                If LastActiveRow < StartRow Then
                    LastActiveRow = StartRow
                End If

                ReasonText = _
                    SafeCellText( _
                        ws.Cells( _
                            LastActiveRow, _
                            ReasonCol))

                IsTechnical = _
                    IsTechnicalReason(ReasonText)

                ApplyJourneyBandFormatting _
                    ws, _
                    StartRow, _
                    LastActiveRow, _
                    LastCol, _
                    IsTechnical, _
                    True

                InBand = False
                StartRow = 0

            End If

        End If

    Next r

    '
    ' SF remains open at the end.
    '
    If InBand Then

        LastActiveRow = LastRow

        ReasonText = _
            SafeCellText( _
                ws.Cells( _
                    LastActiveRow, _
                    ReasonCol))

        IsTechnical = _
            IsTechnicalReason(ReasonText)

        ApplyJourneyBandFormatting _
            ws, _
            StartRow, _
            LastActiveRow, _
            LastCol, _
            IsTechnical, _
            True

    End If

End Sub

Private Sub ApplyJourneyBandFormatting( _
    ByVal ws As Worksheet, _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    ByVal LastCol As Long, _
    ByVal IsTechnical As Boolean, _
    ByVal IsShortfall As Boolean)

    Dim TargetRange As Range

    If StartRow < 2 Then Exit Sub
    If EndRow < StartRow Then Exit Sub
    If LastCol < 1 Then Exit Sub

    Set TargetRange = _
        ws.Range( _
            ws.Cells(StartRow, 1), _
            ws.Cells(EndRow, LastCol))

    ApplyJourneyRangeStyle _
        TargetRange, _
        IsTechnical, _
        IsShortfall

End Sub

Public Sub FormatJourneyColumns()

    Dim ws As Worksheet

    Dim DateCol As Long
    Dim MCSinceCol As Long
    Dim SFSinceCol As Long
    Dim EventCol As Long
    Dim ReasonCol As Long
    Dim CommentCol As Long

    Set ws = Worksheets("NDG Journey")

    DateCol = _
        FindColumnByHeader( _
            ws, _
            "Snapshot Date")

    MCSinceCol = _
        FindColumnByHeader(ws, "MC Since")
    
    SFSinceCol = _
        FindColumnByHeader(ws, "SF Since")

    EventCol = _
        FindColumnByHeader( _
            ws, _
            "Event")

    ReasonCol = _
        FindColumnByHeader( _
            ws, _
            "Reason MC/SF")
            
    CommentCol = _
        FindColumnByHeader( _
            ws, _
            "Comment")

    If DateCol > 0 Then
        ws.Columns(DateCol).AutoFit
    End If

    If MCSinceCol > 0 Then
        ws.Columns(MCSinceCol).AutoFit
    End If

    If SFSinceCol > 0 Then
        ws.Columns(SFSinceCol).AutoFit
    End If

    If EventCol > 0 Then
        ws.Columns(EventCol).AutoFit
    End If

    If ReasonCol > 0 Then
        ws.Columns(ReasonCol).AutoFit
    End If
    
    If CommentCol > 0 Then
        ws.Columns(CommentCol).AutoFit
    End If

End Sub

Private Sub ApplyJourneyRangeStyle( _
    ByVal TargetRange As Range, _
    ByVal IsTechnical As Boolean, _
    ByVal IsShortfall As Boolean)

    Dim BorderStyle As XlLineStyle
    Dim BorderColor As Long
    Dim FillColor As Long
    Dim PatternColor As Long

    '
    ' Colour depends only on MC or SF.
    '
    If IsShortfall Then

        '
        ' SF: red.
        '
        BorderColor = RGB(210, 0, 0)
        PatternColor = RGB(255, 190, 190)
        FillColor = RGB(255, 190, 190)

    Else

        '
        ' MC: orange-yellow.
        '
        BorderColor = RGB(230, 145, 0)
        PatternColor = RGB(255, 220, 130)
        FillColor = RGB(255, 220, 130)

    End If

    '
    ' Technical status affects only pattern and line style.
    '
    If IsTechnical Then

        BorderStyle = xlDash

        With TargetRange.Interior

            .Pattern = xlPatternGray8
            .PatternColor = PatternColor
            .Color = RGB(255, 255, 255)

        End With

    Else

        BorderStyle = xlContinuous

        With TargetRange.Interior

            .Pattern = xlSolid
            .Color = FillColor

        End With

    End If

    With TargetRange.Borders(xlEdgeTop)

        .LineStyle = BorderStyle
        .Color = BorderColor
        .Weight = xlThin

    End With

    With TargetRange.Borders(xlEdgeBottom)

        .LineStyle = BorderStyle
        .Color = BorderColor
        .Weight = xlThin

    End With

End Sub

Private Sub ClearJourneyBandFormatting( _
    ByVal ws As Worksheet, _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    ByVal FirstCol As Long, _
    ByVal LastCol As Long)

    Dim TargetRange As Range

    Set TargetRange = _
        ws.Range( _
            ws.Cells(StartRow, FirstCol), _
            ws.Cells(EndRow, LastCol))

    ClearJourneyRangeStyle TargetRange

End Sub

Private Sub ClearJourneyRangeStyle( _
    ByVal TargetRange As Range)

    With TargetRange.Interior

        .Pattern = xlNone
        .ColorIndex = xlColorIndexNone

    End With

    '
    ' Clear all borders left by previous MC/SF intervals.
    '
    TargetRange.Borders.LineStyle = xlNone

End Sub

Public Function IsTechnicalReason( _
    ByVal ReasonText As String) As Boolean

    Dim NormalizedReason As String

    NormalizedReason = _
        LCase$(Trim$(ReasonText))

    IsTechnicalReason = _
        (InStr( _
            1, _
            NormalizedReason, _
            "technical reason", _
            vbTextCompare) > 0) _
        Or _
        (InStr( _
            1, _
            NormalizedReason, _
            "switch of collateral", _
            vbTextCompare) > 0)

End Function

Public Function SafeCellText( _
    ByVal TargetCell As Range) As String

    If IsError(TargetCell.Value) Then

        SafeCellText = ""

    Else

        SafeCellText = _
            Trim$(CStr(TargetCell.Value2))

    End If

End Function

Private Sub FormatMCByApprovedRatio( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal MCCol As Long, _
    ByVal ApprovedCol As Long)

    Dim ApprovedValue As Double
    Dim MCValue As Double
    Dim MCRatio As Double
    Dim FillColor As Long
    Dim r As Long

    '
    ' Remove MC-specific formatting from a previous run.
    ' Row-level MC/SF fill has already been applied.
    '
    With ws.Range( _
            ws.Cells(2, MCCol), _
            ws.Cells(LastRow, MCCol))

        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlGeneral

    End With

    For r = 2 To LastRow

        MCValue = _
            ParseCsvDouble( _
                ws.Cells(r, MCCol).Value2)

        ApprovedValue = _
            ParseCsvDouble( _
                ws.Cells(r, ApprovedCol).Value2)

        If MCValue > 0 _
           And ApprovedValue > 0 Then

            MCRatio = MCValue / ApprovedValue

            If MCRatio > 1 Then MCRatio = 1
            If MCRatio < 0 Then MCRatio = 0

            FillColor = _
                GetMCRatioFillColor(MCRatio)

            With ws.Cells(r, MCCol)

                .Interior.Pattern = xlSolid
                .Interior.Color = FillColor

                .Font.Color = _
                    GetMCRatioFontColor(MCRatio)

                .Font.Bold = True
                .HorizontalAlignment = xlRight
'                .IndentLevel = 1

            End With

        End If

    Next r

End Sub

Private Function GetMCRatioFillColor( _
    ByVal MCRatio As Double) As Long

    Dim RedValue As Long
    Dim GreenValue As Long
    Dim BlueValue As Long

    Dim LocalRatio As Double

    '
    ' 0% to 50%:
    ' pale yellow -> orange
    '
    If MCRatio <= 0.5 Then

        LocalRatio = _
            MCRatio / 0.5

        RedValue = 255

        GreenValue = _
            CLng(242 - _
                 LocalRatio * (242 - 165))

        BlueValue = _
            CLng(204 - _
                 LocalRatio * 204)

    '
    ' 50% to 100%:
    ' orange -> red
    '
    Else

        LocalRatio = _
            (MCRatio - 0.5) / 0.5

        RedValue = 255

        GreenValue = _
            CLng(165 - _
                 LocalRatio * 165)

        BlueValue = 0

    End If

    GetMCRatioFillColor = _
        RGB( _
            RedValue, _
            GreenValue, _
            BlueValue)

End Function

Private Function GetMCRatioFontColor( _
    ByVal MCRatio As Double) As Long

    If MCRatio >= 0.7 Then

        GetMCRatioFontColor = _
            RGB(255, 255, 255)

    Else

        GetMCRatioFontColor = _
            RGB(120, 40, 0)

    End If

End Function

Private Sub HighlightNonZeroDeltaCells( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal DeltaApprovedCol As Long, _
    ByVal DeltaDrawnCol As Long)

    Dim TargetCols As Variant
    Dim CurrentCol As Variant

    Dim CellValue As Double
    Dim r As Long

    TargetCols = _
        Array( _
            DeltaApprovedCol, _
            DeltaDrawnCol)

    For Each CurrentCol In TargetCols

        For r = 2 To LastRow

            CellValue = _
                ParseCsvDouble( _
                    ws.Cells( _
                        r, _
                        CLng(CurrentCol)).Value2)

            If CellValue <> 0 Then

                With ws.Cells( _
                        r, _
                        CLng(CurrentCol))

                    .Interior.Pattern = xlSolid
                    .Font.Bold = True

                    If CellValue > 0 Then

                        '
                        ' Positive delta: green.
                        '
                        .Interior.Color = _
                            RGB(198, 239, 206)

                        .Font.Color = _
                            RGB(0, 97, 0)

                    Else

                        '
                        ' Negative delta: red.
                        '
                        .Interior.Color = _
                            RGB(255, 199, 206)

                        .Font.Color = _
                            RGB(156, 0, 6)

                    End If

                End With

            End If

        Next r

    Next CurrentCol

End Sub

Private Sub FormatJourneyColumnSeparators( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal LastCol As Long)

    Dim TableRange As Range
    Dim c As Long

    Set TableRange = _
        ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(LastRow, LastCol))

    '
    ' Subtle separator between every column.
    '
    For c = 1 To LastCol - 1

        With ws.Range( _
                ws.Cells(1, c), _
                ws.Cells(LastRow, c)) _
                .Borders(xlEdgeRight)

            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = RGB(220, 220, 220)

        End With

    Next c

    '
    ' Slightly stronger separators between logical groups.
    '
    AddJourneyGroupSeparator _
        ws, LastRow, "Approved"

    AddJourneyGroupSeparator _
        ws, LastRow, "MC"

    AddJourneyGroupSeparator _
        ws, LastRow, "MC Since"

    AddJourneyGroupSeparator _
        ws, LastRow, "Delta Approved"

    AddJourneyGroupSeparator _
        ws, LastRow, "Event"

    AddJourneyGroupSeparator _
        ws, LastRow, "Reason MC/SF"

End Sub

Private Sub AddJourneyGroupSeparator( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal HeaderText As String)

    Dim ColNo As Long

    ColNo = _
        FindColumnByHeader( _
            ws, _
            HeaderText)

    If ColNo = 0 Then Exit Sub

    With ws.Range( _
            ws.Cells(1, ColNo), _
            ws.Cells(LastRow, ColNo)) _
            .Borders(xlEdgeLeft)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(160, 160, 160)

    End With

End Sub

Private Sub FormatLTVTrendCells( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal LTVCol As Long, _
    ByVal DeltaLTVCol As Long)

    Const NoiseThreshold As Double = 0.005
    Const StrongThreshold As Double = 0.02

    Dim DeltaLTV As Double
    Dim r As Long

    If LastRow < 2 Then Exit Sub

    '
    ' Basic display formatting.
    '
    With ws.Range( _
            ws.Cells(2, LTVCol), _
            ws.Cells(LastRow, LTVCol))

        .NumberFormat = "0.0%"
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic

    End With

    With ws.Range( _
            ws.Cells(2, DeltaLTVCol), _
            ws.Cells(LastRow, DeltaLTVCol))
    
        .NumberFormat = _
            "+0.0%;-0.0%;-"
    
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic
    
    End With

    '
    ' First available snapshot has no reliable delta.
    '
    For r = 3 To LastRow

        If Len(ws.Cells(r, DeltaLTVCol).Value2) > 0 _
           And IsNumeric( _
                ws.Cells(r, DeltaLTVCol).Value2) Then

            DeltaLTV = _
                CDbl( _
                    ws.Cells(r, DeltaLTVCol).Value2)

            If Abs(DeltaLTV) >= NoiseThreshold Then

                ApplyLTVTrendHighlight _
                    ws.Cells(r, LTVCol), _
                    DeltaLTV, _
                    StrongThreshold

                ApplyLTVTrendHighlight _
                    ws.Cells(r, DeltaLTVCol), _
                    DeltaLTV, _
                    StrongThreshold

            End If

        End If

    Next r

End Sub

Private Sub ApplyLTVTrendHighlight( _
    ByVal TargetCell As Range, _
    ByVal DeltaLTV As Double, _
    ByVal StrongThreshold As Double)

    With TargetCell

        .Interior.Pattern = xlSolid
        .Font.Bold = True

        If DeltaLTV > 0 Then

            '
            ' LTV increased: deterioration.
            '
            If DeltaLTV >= StrongThreshold Then

                .Interior.Color = _
                    RGB(255, 120, 120)

                .Font.Color = _
                    RGB(135, 0, 0)

            Else

                .Interior.Color = _
                    RGB(255, 199, 206)

                .Font.Color = _
                    RGB(156, 0, 6)

            End If

        Else

            '
            ' LTV decreased: improvement.
            '
            If Abs(DeltaLTV) >= StrongThreshold Then

                .Interior.Color = _
                    RGB(130, 210, 145)

                .Font.Color = _
                    RGB(0, 85, 25)

            Else

                .Interior.Color = _
                    RGB(198, 239, 206)

                .Font.Color = _
                    RGB(0, 97, 0)

            End If

        End If

        .HorizontalAlignment = xlRight
        .IndentLevel = 1

    End With

End Sub


