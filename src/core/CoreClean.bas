Attribute VB_Name = "CoreClean"
Option Explicit

Public Sub Clean()

    Dim KeepSheets As Variant
    Dim i As Long

    KeepSheets = Array( _
        "Home", "Code", "DateRange", "PEC List Old", "PEC List", _
        "Database", "MissingPEC", "Ended Lombards", "Report", _
        "Possible Upsize", "CLN", "CLN Report", _
        "Non-Eligible ISIN", "Companies", "Countries", _
        "Fund Parent Companies", "Bond Issuers", "Equity Names", "Name Variants" _
    )

    With ThisWorkbook

        .Worksheets("Home").Range("Q1:AC99999").Clear

        Application.DisplayAlerts = False

        For i = .Sheets.Count To 1 Step -1
            If IsError(Application.Match( _
                    .Sheets(i).name, KeepSheets, 0)) Then

                .Sheets(i).Delete

            End If
        Next i

        Application.DisplayAlerts = True

    End With

End Sub

