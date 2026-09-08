Attribute VB_Name = "WeeklyAnalysisLayout"
Option Explicit

' v38: formats report tables after their data is written. Header alignment is
' derived only from the final data or from an explicitly supplied source range.
' Paired risk tables allow for six columns plus one spacer column.

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
    ' The left column's tables are five wide now, so this starts one column
    ' further out to keep the single spacer column between them.

    Layout.BreakdownRow = _
        Layout.ReportRow + 3

    Layout.BreakdownCol = _
        Layout.ReportCol + 6

    ' Entered (middle bottom)
    ' The breakdown above it is ten rows: title, header, three dated blocks
    ' of two and the two change rows; one blank row, then this.

    Layout.EnteredRow = _
        Layout.ReportRow + 14

    Layout.EnteredCol = _
        Layout.ReportCol + 6

    ' New Lombards (left, under the overview)
    ' The overview runs to six data rows - year-end, three months, the
    ' compared date, the current date - so its last row is ReportRow + 10;
    ' one blank row, then this.  Same five columns as the overview, so the
    ' three tables read as one column of figures.

    Layout.NewLoanRow = _
        Layout.ReportRow + 12

    Layout.NewLoanCol = _
        Layout.ReportCol

    ' Ended Lombards (left, under the new loans)
    ' A movement table is five rows: title, header, the two dated rows and
    ' the change row; one blank row between.

    Layout.EndedLoanRow = _
        Layout.ReportRow + 18

    Layout.EndedLoanCol = _
        Layout.ReportCol

    ' Pie Chart
    ' Below the entered table, which is seven rows: title, header, two dated
    ' pairs of amounts and shares, and the change row.

    Layout.PieRow = _
        Layout.ReportRow + 22

    Layout.PieCol = _
        Layout.ReportCol + 6
        
    Layout.PieHeightRows = 22
        
    ' Comments
    ' Under the ended loans, as wide as the three tables above it (five
    ' columns); the box runs from the row after this to the bottom of the
    ' pie.

    Layout.CommentRow = _
        Layout.ReportRow + 24

    Layout.CommentCol = _
        Layout.ReportCol

    ' Risk concentration (right of the breakdown, aligned with the upper
    ' report tables).  The movement tables moved under the overview, so this
    ' follows the breakdown's last column with the single spacer column
    ' between them.

    Layout.RiskRow = _
        Layout.ReportRow + 3

    Layout.RiskCol = _
        Layout.ReportCol + 16

    ' Risk concentration excluding DPM
    ' One empty spacer column is left between the two six-column tables:
    ' rank, the move since the compared date, name, and three figures.

    Layout.RiskExSegRow = _
        Layout.RiskRow

    Layout.RiskExSegCol = _
        Layout.RiskCol + 7

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

Private Sub AddRightBorder( _
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
