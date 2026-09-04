Attribute VB_Name = "CoreReportFormat"
Option Explicit

'
' Table formatting shared by every report this workbook builds:
' the weekly analysis, the revenue summary and the NDG journey.
' It knows about header rows, body rows and total rows, and
' nothing about what any of them contain.
'

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

Private Sub AddMediumTopBorder( _
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
