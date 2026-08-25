Attribute VB_Name = "RevenueTools"
Option Explicit

Public Sub BuildRevenueSummary()

    '=========================================================================
    ' Worksheets
    '=========================================================================

    Dim wsCode As Worksheet
    Dim wsDelta As Worksheet
    Dim wsRevenueSummary As Worksheet

    '=========================================================================
    ' Row / Loop Variables
    '=========================================================================

    Dim LastRow As Long
    Dim OutputRow As Long

    Dim i As Long
    Dim c As Long

    Dim k As Variant

    '=========================================================================
    ' Columns
    '=========================================================================

    Dim AssetClassCol As Long
    Dim BookingDateCol As Long
    Dim PositionValueCol As Long

    '=========================================================================
    ' Revenue Calculation
    '=========================================================================

    Dim ForecastYear As Long

    Dim AssetClass As String

    Dim BookingDate As Date
    Dim PositionValue As Double

    Dim RevenueRate As Double

    Dim AnnualRevenue As Double
    Dim FYRevenue As Double

    Dim AnnualAMRevenue As Double
    Dim FYAMRevenue As Double

    Dim EligibleMonths As Long

    '=========================================================================
    ' Aggregations
    '=========================================================================

    Dim TotalEligibleAssets As Double

    Dim TotalLombardAnnualRevenue As Double
    Dim TotalLombardFYRevenue As Double

    Dim AssetManagementAnnualRevenue As Double
    Dim AssetManagementFYRevenue As Double

    '=========================================================================
    ' Dictionaries
    '=========================================================================

    Dim DictAssetVolume As Object
    Dim DictAnnualRevenue As Object
    Dim DictFYRevenue As Object
    Dim DictAnnualAMRevenue As Object
    Dim DictFYAMRevenue As Object

    Set DictAssetVolume = _
        CreateObject("Scripting.Dictionary")

    Set DictAnnualRevenue = _
        CreateObject("Scripting.Dictionary")

    Set DictFYRevenue = _
        CreateObject("Scripting.Dictionary")

    Set DictAnnualAMRevenue = _
        CreateObject("Scripting.Dictionary")

    Set DictFYAMRevenue = _
        CreateObject("Scripting.Dictionary")

    '=========================================================================
    ' Setup
    '=========================================================================

    Set wsCode = Worksheets("Home")

    ForecastYear = _
        Year(wsCode.Range("AnalysisEndDate").Value)

    Set wsDelta = Worksheets( _
        "Delta_" & _
        GetDateCode( _
            wsCode.Range("AnalysisEndDate").Value))

    '=========================================================================
    ' Locate Columns
    '=========================================================================

    AssetClassCol = _
        FindColumnByHeader( _
            wsDelta, _
            "Asset Class")

    BookingDateCol = _
        FindColumnByHeader( _
            wsDelta, _
            "Booking Date")

    PositionValueCol = _
        FindColumnByHeader( _
            wsDelta, _
            "Position Value")

    If AssetClassCol = 0 Then
        Fatal "Column not found: Asset Class"
    End If

    If BookingDateCol = 0 Then
        Fatal "Column not found: Booking Date"
    End If

    If PositionValueCol = 0 Then
        Fatal "Column not found: Position Value"
    End If

    '=========================================================================
    ' Aggregate Revenue Data
    '=========================================================================

    LastRow = _
        GetLastRow(wsDelta, "A")

    For i = 2 To LastRow

        AssetClass = _
            Trim$(CStr( _
                wsDelta.Cells(i, AssetClassCol).Value))

        If AssetClass <> "" _
            And AssetClass <> "Non Eligible Asset" Then

            BookingDate = _
                wsDelta.Cells(i, BookingDateCol).Value

            PositionValue = _
                ParseCsvDouble( _
                    wsDelta.Cells(i, PositionValueCol).Value)

            RevenueRate = _
                GetRevenueRate(AssetClass)

            EligibleMonths = _
                GetMonthsInForecastYear( _
                    BookingDate, _
                    ForecastYear)

            AnnualRevenue = _
                PositionValue * RevenueRate

            FYRevenue = _
                AnnualRevenue * _
                EligibleMonths / 12#

            AnnualAMRevenue = _
                PositionValue * 0.0014

            FYAMRevenue = _
                AnnualAMRevenue * _
                EligibleMonths / 12#

            If DictAssetVolume.Exists(AssetClass) Then

                DictAssetVolume(AssetClass) = _
                    DictAssetVolume(AssetClass) + _
                    PositionValue

                DictAnnualRevenue(AssetClass) = _
                    DictAnnualRevenue(AssetClass) + _
                    AnnualRevenue

                DictFYRevenue(AssetClass) = _
                    DictFYRevenue(AssetClass) + _
                    FYRevenue

                DictFYAMRevenue(AssetClass) = _
                    DictFYAMRevenue(AssetClass) + _
                    FYAMRevenue

            Else

                DictAssetVolume.Add _
                    AssetClass, _
                    PositionValue

                DictAnnualRevenue.Add _
                    AssetClass, _
                    AnnualRevenue

                DictFYRevenue.Add _
                    AssetClass, _
                    FYRevenue

                DictAnnualAMRevenue.Add _
                    AssetClass, _
                    AnnualAMRevenue

                DictFYAMRevenue.Add _
                    AssetClass, _
                    FYAMRevenue

            End If

            TotalEligibleAssets = _
                TotalEligibleAssets + _
                PositionValue

            TotalLombardAnnualRevenue = _
                TotalLombardAnnualRevenue + _
                AnnualRevenue

            TotalLombardFYRevenue = _
                TotalLombardFYRevenue + _
                FYRevenue

            AssetManagementAnnualRevenue = _
                AssetManagementAnnualRevenue + _
                AnnualAMRevenue

            AssetManagementFYRevenue = _
                AssetManagementFYRevenue + _
                FYAMRevenue

        End If

    Next i

    '=========================================================================
    ' Create Output Sheet
    '=========================================================================

    Set wsRevenueSummary = _
        CreateOrReplaceSheet("Revenue Summary")

    wsRevenueSummary.Cells(1, 1).Value = "Asset Class"
    wsRevenueSummary.Cells(1, 2).Value = "Eligible Assets (€)"
    wsRevenueSummary.Cells(1, 3).Value = "Revenue Rate (%)"
    wsRevenueSummary.Cells(1, 4).Value = "Annualized Revenue Estimate (€)1"
    wsRevenueSummary.Cells(1, 5).Value = _
        "FY " & ForecastYear & _
        " Revenue Impact (€)2"

    For c = 4 To 5

        With wsRevenueSummary.Cells(1, c)

            .Characters( _
                Len(.Value), _
                1).Font.Superscript = True

        End With

    Next c

    '=========================================================================
    ' Write Revenue Summary
    '=========================================================================

    OutputRow = 2

    For Each k In DictAssetVolume.Keys

        wsRevenueSummary.Cells(OutputRow, 1).Value = k

        wsRevenueSummary.Cells(OutputRow, 2).Value = _
            DictAssetVolume(k)

        wsRevenueSummary.Cells(OutputRow, 3).Value = _
            GetRevenueRate(k)

'        wsRevenueSummary.Cells(OutputRow, 4).Value = _
'            DictAnnualRevenue(k)
        
        wsRevenueSummary.Cells(OutputRow, 4).Formula = _
            "=B" & OutputRow & "*C" & OutputRow

        wsRevenueSummary.Cells(OutputRow, 5).Value = _
            DictFYRevenue(k)

        OutputRow = OutputRow + 1

    Next k

    wsRevenueSummary.Cells(OutputRow, 1).Value = _
        "Total Lombard"

'    wsRevenueSummary.Cells(OutputRow, 4).Value = _
'        TotalLombardAnnualRevenue

    wsRevenueSummary.Cells(OutputRow, 4).Formula = _
        "=SUM(D" & (OutputRow - UBound(DictAssetVolume.Keys) - 1) & _
        ":D" & (OutputRow - 1) & ")"

''    wsRevenueSummary.Cells(OutputRow, 5).Value = _
''        TotalLombardFYRevenue

    wsRevenueSummary.Cells(OutputRow, 5).Formula = _
        "=SUM(E" & (OutputRow - UBound(DictAssetVolume.Keys) - 1) & _
        ":E" & (OutputRow - 1) & ")"

'    AssetManagementAnnualRevenue = _
'        TotalEligibleAssets * 0.0014
'
'    AssetManagementFYRevenue = _
'        AssetManagementAnnualRevenue

    wsRevenueSummary.Cells(OutputRow + 1, 1).Value = _
        "Asset Management Uplift"
        
'    With wsRevenueSummary.Cells(OutputRow + 1, 1)
'
'        .Characters( _
'            Len(.Value), _
'            1).Font.Superscript = True
'
'    End With

'    wsRevenueSummary.Cells(OutputRow + 1, 2).Value = _
'        TotalEligibleAssets
        
    wsRevenueSummary.Cells(OutputRow + 1, 2).Formula = _
        "=SUM(B" & (OutputRow - UBound(DictAssetVolume.Keys) - 1) & _
        ":B" & (OutputRow - 1) & ")"

    wsRevenueSummary.Cells(OutputRow + 1, 3).Value = _
        0.0014

'    wsRevenueSummary.Cells(OutputRow + 1, 4).Value = _
'        AssetManagementAnnualRevenue
        
    wsRevenueSummary.Cells(OutputRow + 1, 4).Formula = _
        "=B" & (OutputRow + 1) & _
        "*C" & (OutputRow + 1)

    wsRevenueSummary.Cells(OutputRow + 1, 5).Value = _
        AssetManagementFYRevenue

    wsRevenueSummary.Cells(OutputRow + 2, 1).Value = _
        "Total"

'    wsRevenueSummary.Cells(OutputRow + 2, 4).Value = _
'        TotalLombardAnnualRevenue + _
'        AssetManagementAnnualRevenue

    wsRevenueSummary.Cells(OutputRow + 2, 4).Formula = _
        "=D" & (OutputRow) & _
        "+D" & (OutputRow + 1)

'    wsRevenueSummary.Cells(OutputRow + 2, 5).Value = _
'        TotalLombardFYRevenue + _
'        AssetManagementFYRevenue

    wsRevenueSummary.Cells(OutputRow + 2, 5).Formula = _
        "=E" & (OutputRow) & _
        "+E" & (OutputRow + 1)

    '=========================================================================
    ' Formatting
    '=========================================================================

    wsRevenueSummary.Columns("B:B").NumberFormat = _
        "#,##0.00"

    wsRevenueSummary.Columns("C:C").NumberFormat = _
        "0.00%"

    wsRevenueSummary.Columns("D:E").NumberFormat = _
        "#,##0.00"

    wsRevenueSummary.Columns.AutoFit

    '=========================================================================
    ' Notes
    '=========================================================================

    OutputRow = OutputRow + 4

'    wsRevenueSummary.Cells(OutputRow, 1).Value = "Notes"
'    wsRevenueSummary.Cells(OutputRow, 1).Font.Bold = True

'    With wsRevenueSummary.Range("A" & OutputRow & ":E" & OutputRow)
'
'        .Merge
'        .Value = _
'            "1 New eligible collateral volume identified within the selected analysis period."
'        .WrapText = True
'
'    End With

'    With wsRevenueSummary.Range("A" & OutputRow + 1 & ":E" & OutputRow + 1)
'
'        .Merge
'        .Value = _
'            "2 Predefined annual revenue rate assumption by collateral asset class."
'        .WrapText = True
'
'    End With

    With wsRevenueSummary.Range("A" & OutputRow & ":E" & OutputRow)

        .Merge
        .Value = _
            "1 Estimated annual revenue potential under current conditions, assuming the identified collateral portfolio remains unchanged."
        .WrapText = True

    End With

    With wsRevenueSummary.Range("A" & OutputRow + 1 & ":E" & OutputRow + 1)

        .Merge
        .Value = _
            "2 Estimated revenue contribution expected within the current fiscal year, considering the booking date of the collateral."
        .WrapText = True

    End With

'    With wsRevenueSummary.Range("A" & OutputRow + 4 & ":E" & OutputRow + 4)
'
'        .Merge
'        .Value = _
'            "5 An additional annual revenue rate of 0.14% is applied to eligible assets to estimate Asset Management revenues."
'        .WrapText = True
'
'    End With
    
    For c = 0 To 1

        With wsRevenueSummary.Cells(OutputRow + c, 1)

            .Characters( _
                0, _
                1).Font.Superscript = True

        End With

    Next c

'    wsRevenueSummary.Rows(OutputRow + 1).RowHeight = 15
'    wsRevenueSummary.Rows(OutputRow + 2).RowHeight = 15
'    wsRevenueSummary.Rows(OutputRow + 3).RowHeight = 15
'    wsRevenueSummary.Rows(OutputRow + 4).RowHeight = 30

FormatRevenueSummary _
    wsRevenueSummary, _
    OutputRow - 2

End Sub

Public Sub FormatRevenueSummary( _
    ByVal ws As Worksheet, _
    ByVal LastSummaryRow As Long)

    Dim NotesRow As Long

    NotesRow = LastSummaryRow + 2

    '=========================================================
    ' Main Table
    '=========================================================

    FormatReportTable _
        ws.Range( _
            ws.Cells(1, 1), _
            ws.Cells(LastSummaryRow, 5))

    '=========================================================
    ' First Column
    '=========================================================

'    FormatFirstColumn _
'        ws, _
'        2, _
'        LastSummaryRow, _
'        1
    ws.Cells(1, 1).HorizontalAlignment = xlLeft

    '=========================================================
    ' Totals Section
    '=========================================================

    FormatTotalRow ws, 1, 5, LastSummaryRow - 2
    FormatTotalRow ws, 1, 5, LastSummaryRow

'    AddTopBorder _
'        ws, _
'        LastSummaryRow - 2, _
'        1, _
'        5
'
'    AddBottomBorder _
'        ws, _
'        LastSummaryRow, _
'        1, _
'        5

    '=========================================================
    ' Number Formats
    '=========================================================

    ws.Columns("B:B").NumberFormat = _
        "#,##0.00"

    ws.Columns("C:C").NumberFormat = _
        "0.00%"

    ws.Columns("D:E").NumberFormat = _
        "#,##0.00"

    '=========================================================
    ' Notes Section
    '=========================================================

'    ws.Cells(NotesRow, 1).Font.Bold = True

    With ws.Range( _
        ws.Cells(NotesRow, 1), _
        ws.Cells(NotesRow + 1, 5))
        .Font.Size = 9

'        .Borders(xlEdgeTop).LineStyle = _
'            xlContinuous
'
'        .Borders(xlEdgeTop).Weight = _
'            xlThin

'        .Borders(xlEdgeTop).Color = _
'            RGB(128, 128, 128)

    End With

'    ws.Rows(NotesRow + 1).RowHeight = 18
'    ws.Rows(NotesRow + 2).RowHeight = 18
'    ws.Rows(NotesRow + 3).RowHeight = 30
'    ws.Rows(NotesRow + 4).RowHeight = 30
'    ws.Rows(NotesRow + 5).RowHeight = 18

    '=========================================================
    ' General
    '=========================================================

    ws.Columns.AutoFit

End Sub

Public Function GetRevenueRate( _
    ByVal AssetClass As String) As Double

    Select Case AssetClass

        Case "Bonds"
            GetRevenueRate = 0.003

        Case "Insurance"
            GetRevenueRate = 0.003

        Case "GP"
            GetRevenueRate = 0.005

        Case "Certificates"
            GetRevenueRate = 0.009

        Case "Equity"
            GetRevenueRate = 0.003

        Case "Funds"
            GetRevenueRate = 0.003

        Case "Cash"
            GetRevenueRate = 0#

        Case Else
            GetRevenueRate = 0#

    End Select

End Function

Public Function GetMonthsInForecastYear( _
    ByVal StartDate As Date, _
    ByVal ForecastYear As Long) As Long

    Dim StartYear As Long

    StartYear = Year(StartDate)

    Select Case True

        Case StartYear < ForecastYear

            GetMonthsInForecastYear = 12

        Case StartYear = ForecastYear

            GetMonthsInForecastYear = _
                12 - Month(StartDate) + 1

        Case Else

            GetMonthsInForecastYear = 0

    End Select

End Function

Public Function GetForecastYear( _
    ByVal EndDate As Date) As Long

    GetForecastYear = Year(EndDate)

End Function



