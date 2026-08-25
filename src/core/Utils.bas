Attribute VB_Name = "Utils"
Option Explicit

Public Function CreateOrReplaceSheet( _
    ByVal SheetName As String) As Worksheet

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim PreviousDisplayAlerts As Boolean
    Dim PreviousEnableEvents As Boolean

    Dim i As Long

    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    Set wb = ThisWorkbook

    PreviousDisplayAlerts = Application.DisplayAlerts
    PreviousEnableEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    If wb.ProtectStructure Then

        Err.Raise _
            vbObjectError + 9010, _
            "CreateOrReplaceSheet", _
            "The workbook structure is protected."

    End If

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '
    ' Try to get existing worksheet.
    '
    On Error Resume Next

    Set ws = wb.Worksheets(SheetName)

    On Error GoTo ErrorHandler

    If ws Is Nothing Then

        '
        ' Sheet does not exist:
        ' create a new one.
        '
        Set ws = wb.Worksheets.Add( _
            After:=wb.Sheets(wb.Sheets.Count))

        ws.name = SheetName

    Else

        '
        ' Sheet already exists:
        ' keep the same Worksheet object.
        '
        ' This avoids invalidating existing
        ' Worksheet references elsewhere.
        '

'
' Remove active filters first.
'
If ws.FilterMode Then
    ws.ShowAllData
End If

If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
End If

'
' Convert existing Excel tables back to ordinary ranges.
' The sheet will be cleared afterwards.
'
For i = ws.ListObjects.Count To 1 Step -1
    ws.ListObjects(i).Unlist
Next i

'
' Unmerge only the worksheet area that has actually been used.
'
ws.UsedRange.UnMerge

'
' Remove shapes, charts, buttons, etc.
'
For i = ws.Shapes.Count To 1 Step -1
    ws.Shapes(i).Delete
Next i

        '
        ' Clear cells including values,
        ' formulas and formatting.
        '
        ws.Cells.Clear

    End If

    Set CreateOrReplaceSheet = ws

CleanExit:

    Application.DisplayAlerts = PreviousDisplayAlerts
    Application.EnableEvents = PreviousEnableEvents

    Exit Function


ErrorHandler:

    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    Application.DisplayAlerts = PreviousDisplayAlerts
    Application.EnableEvents = PreviousEnableEvents

    On Error GoTo 0

    Err.Raise _
        ErrorNumber, _
        ErrorSource, _
        ErrorDescription

End Function

Public Function SheetExists(ByVal SheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function

Public Function GetLastRow( _
    ws As Worksheet, _
    Col As Variant) As Long

    GetLastRow = ws.Cells( _
        ws.Rows.Count, Col).End(xlUp).Row

End Function

Public Function GetYTDDate( _
    ByVal ReferenceDate As Date) As Date

    GetYTDDate = DateSerial( _
        Year(ReferenceDate) - 1, _
        12, _
        31)

End Function

Public Sub Note( _
    ByVal Message As String)

    If NoteHandler <> "" Then

        Application.Run NoteHandler, Message

    Else

        MsgBox Message

    End If

End Sub

Public Sub Fatal( _
    ByVal Message As String)

    Err.Raise _
        vbObjectError + 9000, _
        "Fatal Error", _
        Message

End Sub

Public Sub ResetExcel()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False

End Sub

Public Function ParseCsvDouble( _
    ByVal ValueIn As Variant) As Double

    ParseCsvDouble = _
        VBA.Val(CStr(Nz(ValueIn, 0)))

End Function

Public Function WorksheetDouble( _
    ByVal ValueIn As Variant) As Double

    If IsError(ValueIn) _
       Or IsEmpty(ValueIn) _
       Or IsNull(ValueIn) _
       Or Trim$(CStr(ValueIn)) = "" Then

        WorksheetDouble = 0

    ElseIf IsNumeric(ValueIn) Then

        WorksheetDouble = _
            CDbl(ValueIn)

    Else

        WorksheetDouble = 0

    End If

End Function

Public Function FindColumnByHeader( _
    ByVal ws As Worksheet, _
    ByVal HeaderName As String, _
    Optional ByVal HeaderRow As Long = 1) As Long

    Dim LastCol As Long
    Dim c As Long

    LastCol = _
        ws.Cells(HeaderRow, ws.Columns.Count) _
        .End(xlToLeft).Column

    For c = 1 To LastCol

        If Trim$(CStr(ws.Cells(HeaderRow, c).Value)) = _
           HeaderName Then

            FindColumnByHeader = c
            Exit Function

        End If

    Next c

    FindColumnByHeader = 0

End Function

Public Sub RenameHeader( _
    ByVal ws As Worksheet, _
    ByVal OldName As String, _
    ByVal NewName As String)

    Dim Col As Long

    Col = _
        FindColumnByHeader( _
            ws, _
            OldName)

    If Col > 0 Then

        ws.Cells(1, Col).Value = _
            NewName

    End If

End Sub

Public Function GetNumericValue( _
    ByVal v As Variant) As Double

    If IsError(v) Then
        GetNumericValue = 0
        Exit Function
    End If

    If IsEmpty(v) Then
        GetNumericValue = 0
        Exit Function
    End If

    If Len(Trim$(CStr(v))) = 0 Then
        GetNumericValue = 0
        Exit Function
    End If

    If Not IsNumeric(v) Then
        GetNumericValue = 0
        Exit Function
    End If

    GetNumericValue = CDbl(v)

End Function

