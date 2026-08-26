Attribute VB_Name = "WeeklyAnalysisLayout"
Option Explicit

' v38: formats report tables after their data is written. Header alignment is
' derived only from the final data or from an explicitly supplied source range.
' Paired risk tables allow for five columns plus one spacer column.

Public Type ReportLayout

    ReportRow As Long
    ReportCol As Long

    HeaderRow As Long
    HeaderCol As Long
    
    PortfolioRow As Long
    PortfolioCol As Long

    BreakdownRow As Long
    BreakdownCol As Long

    NewLoanRow As Long
    NewLoanCol As Long

    EndedLoanRow As Long
    EndedLoanCol As Long

    EnteredRow As Long
    EnteredCol As Long

    PieRow As Long
    PieCol As Long
    PieHeightRows As Long
    
    CommentRow As Long
    CommentCol As Long

    RiskRow As Long
    RiskCol As Long

    RiskExSegRow As Long
    RiskExSegCol As Long

    RiskSectionGapRows As Long

    CountryRiskRow As Long
    CountryRiskCol As Long

    CountryRiskExSegRow As Long
    CountryRiskExSegCol As Long

    SectorRiskRow As Long
    SectorRiskCol As Long

    SectorRiskExSegRow As Long
    SectorRiskExSegCol As Long

End Type

Public Layout As ReportLayout

Public Sub InitializeLayout()

    Layout.ReportRow = 2
    Layout.ReportCol = 2
    
    
    Layout.HeaderRow = Layout.ReportRow
    Layout.HeaderCol = Layout.ReportCol


    ' Portfolio (left)

    Layout.PortfolioRow = _
        Layout.ReportRow + 3

    Layout.PortfolioCol = _
        Layout.ReportCol
    

    ' Breakdown (middle top)

    Layout.BreakdownRow = _
        Layout.ReportRow + 3

    Layout.BreakdownCol = _
        Layout.ReportCol + 5

    ' Entered (middle bottom)

    Layout.EnteredRow = _
        Layout.ReportRow + 12

    Layout.EnteredCol = _
        Layout.ReportCol + 5

    ' New Lombards (right top)

    Layout.NewLoanRow = _
        Layout.ReportRow + 3

    Layout.NewLoanCol = _
        Layout.ReportCol + 15

    ' Ended Lombards (right bottom)

    Layout.EndedLoanRow = _
        Layout.ReportRow + 8

    Layout.EndedLoanCol = _
        Layout.ReportCol + 15

    ' Pie Chart

    Layout.PieRow = _
        Layout.ReportRow + 18

    Layout.PieCol = _
        Layout.ReportCol + 5
        
    Layout.PieHeightRows = 22
        
    ' Comments

    Layout.CommentRow = _
        Layout.ReportRow + 12

    Layout.CommentCol = _
        Layout.ReportCol

    ' Risk concentration (far right, aligned with the upper report tables)

    Layout.RiskRow = _
        Layout.ReportRow + 3

    Layout.RiskCol = _
        Layout.ReportCol + 20

    ' Risk concentration excluding DPM
    ' One empty spacer column is left between the two five-column tables.

    Layout.RiskExSegRow = _
        Layout.RiskRow

    Layout.RiskExSegCol = _
        Layout.RiskCol + 6

    ' The main module recalculates the rows below after each pair of risk
    ' tables has been written. This remains the one central spacing control.

    Layout.RiskSectionGapRows = 1

    ' Provisional values; BuildRiskGranularitySection replaces them with
    ' positions based on the actual table heights.

    Layout.CountryRiskRow = _
        Layout.RiskRow

    Layout.CountryRiskCol = _
        Layout.RiskCol

    Layout.CountryRiskExSegRow = _
        Layout.CountryRiskRow

    Layout.CountryRiskExSegCol = _
        Layout.RiskExSegCol

    Layout.SectorRiskRow = _
        Layout.RiskRow

    Layout.SectorRiskCol = _
        Layout.RiskCol

    Layout.SectorRiskExSegRow = _
        Layout.SectorRiskRow

    Layout.SectorRiskExSegCol = _
        Layout.RiskExSegCol

End Sub

Public Sub AddMediumTopBorder( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal FirstCol As Long, _
    ByVal LastCol As Long)

    With ws.Range( _
        ws.Cells(RowNo, FirstCol), _
        ws.Cells(RowNo, LastCol)).Borders(xlEdgeTop)

        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(128, 128, 128)

    End With

End Sub

Public Sub AddBottomBorder( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal FirstCol As Long, _
    ByVal LastCol As Long)

    With ws.Range( _
        ws.Cells(RowNo, FirstCol), _
        ws.Cells(RowNo, LastCol)).Borders(xlEdgeBottom)

        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(128, 128, 128)

    End With

End Sub

Public Sub AddRightBorder( _
    ByVal ws As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal LastRow As Long, _
    ByVal ColNo As Long)

    With ws.Range( _
        ws.Cells(FirstRow, ColNo), _
        ws.Cells(LastRow, ColNo)).Borders(xlEdgeRight)

        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(128, 128, 128)

    End With

End Sub

Public Sub FormatFirstColumn( _
    ByVal ws As Worksheet, _
    ByVal FirstDataRow As Long, _
    ByVal LastDataRow As Long, _
    ByVal ColNo As Long)

    With ws.Range( _
        ws.Cells(FirstDataRow, ColNo), _
        ws.Cells(LastDataRow, ColNo))

        .Font.Bold = True

        .HorizontalAlignment = xlLeft

'        .VerticalAlignment = xlCenter

    End With

    AddRightBorder _
        ws, _
        FirstDataRow, _
        LastDataRow, _
        ColNo

End Sub

Public Sub FormatTotalRow( _
    ByVal ws As Worksheet, _
    ByVal FirstDataCol As Long, _
    ByVal LastDataCol As Long, _
    ByVal RowNo As Long)

    With ws.Range( _
        ws.Cells(RowNo, FirstDataCol), _
        ws.Cells(RowNo, LastDataCol))

        .Font.Bold = True

'        .HorizontalAlignment = xlLeft

'        .VerticalAlignment = xlCenter

    End With

    AddMediumTopBorder _
        ws, _
        RowNo, _
        FirstDataCol, _
        LastDataCol

End Sub

Public Sub FormatReportTable( _
    ByVal rng As Range, _
    Optional ByVal HeaderRows As Long = 1, _
    Optional ByVal AlignmentSource As Range = Nothing)

    Dim DataRow As Long

    With rng

'        .Font.Name = "UniCredit"
        .Font.Size = 10

        .VerticalAlignment = xlCenter

    End With

    '
    ' Header
    '

    With rng.Rows(1).Resize(HeaderRows)

        .Font.Bold = True

        .VerticalAlignment = xlCenter

        .Interior.Color = RGB(255, 255, 255)

    End With

    AlignReportHeaderColumns _
        rng, _
        HeaderRows, _
        AlignmentSource

    '
    ' Header separator
    '

    With rng.Rows(HeaderRows).Borders(xlEdgeBottom)

        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(111, 38, 61)

    End With

    '
    ' Zebra stripes
    '

    For DataRow = HeaderRows + 1 To rng.Rows.Count

        If (DataRow - HeaderRows) Mod 2 = 1 Then

            rng.Rows(DataRow).Interior.Color = _
                RGB(242, 242, 242)

        Else

            rng.Rows(DataRow).Interior.Color = _
                RGB(255, 255, 255)

        End If

    Next DataRow

End Sub

Private Sub AlignReportHeaderColumns( _
    ByVal rng As Range, _
    ByVal HeaderRows As Long, _
    Optional ByVal AlignmentSource As Range = Nothing)

    Dim DataCol As Long
    Dim DataRow As Long

    Dim NumericCount As Long
    Dim TextCount As Long

    Dim CellValue As Variant
    Dim HeaderAlignment As XlHAlign

    Dim SourceRange As Range

    If rng Is Nothing Then Exit Sub
    If HeaderRows < 1 Then Exit Sub
    If HeaderRows > rng.Rows.Count Then Exit Sub

    If AlignmentSource Is Nothing Then

        If HeaderRows >= rng.Rows.Count Then Exit Sub

        Set SourceRange = _
            rng.Offset(HeaderRows, 0).Resize( _
                rng.Rows.Count - HeaderRows, _
                rng.Columns.Count)

    Else

        Set SourceRange = AlignmentSource

    End If

    If SourceRange Is Nothing Then Exit Sub
    If SourceRange.Columns.Count < rng.Columns.Count Then Exit Sub

    For DataCol = 1 To rng.Columns.Count

        NumericCount = 0
        TextCount = 0

        For DataRow = 1 To SourceRange.Rows.Count

            If Not SourceRange.Cells( _
                    DataRow, _
                    DataCol).MergeCells Then

                CellValue = _
                    SourceRange.Cells( _
                        DataRow, _
                        DataCol).Value2

                If Not IsError(CellValue) Then

                    If Len(Trim$(CStr(CellValue))) > 0 Then

                        If IsNumeric(CellValue) Or _
                           IsDate(CellValue) Then

                            NumericCount = _
                                NumericCount + 1

                        Else

                            TextCount = TextCount + 1

                        End If

                    End If

                End If

            End If

        Next DataRow

        If NumericCount > TextCount Then

            HeaderAlignment = xlRight

        ElseIf TextCount > NumericCount Then

            HeaderAlignment = xlLeft

        Else

            HeaderAlignment = xlRight

        End If

        rng.Cells( _
            HeaderRows, _
            DataCol).HorizontalAlignment = _
            HeaderAlignment

    Next DataCol

End Sub


