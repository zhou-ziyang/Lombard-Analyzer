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
