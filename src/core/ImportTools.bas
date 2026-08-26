Attribute VB_Name = "ImportTools"
Option Explicit

Public MissingFiles As String
Public OverwriteExistingSheets As Boolean
Public SheetOverwriteDecisionMade As Boolean

Public Function PathSelection() As String

    Dim Path As String

    Path = Trim( _
        ThisWorkbook.Worksheets("Home") _
        .Range("path").Value)

    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

    If Len(Dir(Path, vbDirectory)) = 0 Then
        Fatal _
            "Source folder does not exist:" & _
            vbCrLf & vbCrLf & _
            Path
    End If

    PathSelection = Path

End Function

Public Function GetComparisonDate( _
    ByVal ReferenceDate As Date, _
    Optional ByVal MonthsBack As Long = 1) As Date

    Dim TargetYear As Long
    Dim TargetMonth As Long
    Dim TargetDay As Long
    Dim LastDayOfTargetMonth As Long

    TargetYear = Year(ReferenceDate)
    TargetMonth = Month(ReferenceDate) - MonthsBack
    TargetDay = Day(ReferenceDate)

    '
    ' DateSerial carries a month number outside 1-12 into the neighbouring
    ' year by itself, which is what we want. It also carries a day number
    ' the month does not have into the following month, which is not: one
    ' month before 31 March would come back as 3 March. Day zero of the
    ' next month is the last day of this one, so clamp to that.
    '
    LastDayOfTargetMonth = _
        Day(DateSerial(TargetYear, TargetMonth + 1, 0))

    If TargetDay > LastDayOfTargetMonth Then
        TargetDay = LastDayOfTargetMonth
    End If

    GetComparisonDate = _
        DateSerial(TargetYear, TargetMonth, TargetDay)

End Function

Public Sub ResetSheetOverwriteDecision()

    OverwriteExistingSheets = False
    SheetOverwriteDecisionMade = False

End Sub

Public Function ShouldOverwriteExistingSheets( _
    Optional ByVal ExistingItemDescription As String = _
        "generated worksheets") As Boolean

    Dim Answer As VbMsgBoxResult

    If Not SheetOverwriteDecisionMade Then

        Answer = MsgBox( _
            "Existing " & ExistingItemDescription & _
            " were found." & vbCrLf & vbCrLf & _
            "YES = Reload/rebuild from current source data" & vbCrLf & _
            "NO = Reuse them unchanged", _
            vbYesNo + vbQuestion, _
            "Reuse Existing Data")

        OverwriteExistingSheets = _
            (Answer = vbYes)

        SheetOverwriteDecisionMade = True

    End If

    ShouldOverwriteExistingSheets = _
        OverwriteExistingSheets

End Function

Public Function ImportAccountsByDate( _
    ByVal DateCode As String) As Worksheet

    Set ImportAccountsByDate = _
        ImportCsvByDate( _
            DateCode, _
            "Accounts")

End Function

Public Function ImportPositionsByDate( _
    ByVal DateCode As String) As Worksheet

    Set ImportPositionsByDate = _
        ImportCsvByDate( _
            DateCode, _
            "Positions")

End Function

Private Function ImportCsvByDate( _
    ByVal DateCode As String, _
    ByVal FileType As String) As Worksheet

    Dim ws As Worksheet

    Dim Path As String
    Dim csvFilePath As String

    Dim SheetName As String
    Dim FileSuffix As String

    Dim csvData As Variant
    Dim Headers As String
    Dim Line As String

    Dim i As Long
    Dim LastRow As Long

    Select Case UCase(FileType)

        Case "POSITIONS"

            FileSuffix = _
                "_Lombard_Loans_ITA_Positions.csv"

        Case "ACCOUNTS"

            FileSuffix = _
                "_Lombard_Loans_ITA_Accounts.csv"

        Case Else

            Err.Raise vbObjectError + 1000, _
                      "ImportCsvByDate", _
                      "Unknown file type"

    End Select

    Path = PathSelection()

    csvFilePath = _
        Path & _
        DateCode & _
        FileSuffix

    If Dir(csvFilePath) = "" Then

        MissingFiles = _
            MissingFiles & vbCrLf & _
            csvFilePath

        Exit Function

    End If

    If FileLen(csvFilePath) = 0 Then

        MissingFiles = _
            MissingFiles & vbCrLf & _
            csvFilePath

        Exit Function

    End If

    SheetName = _
        FileType & _
        " Output " & _
        DateCode

    If SheetExists(SheetName) Then

        If Not ShouldOverwriteExistingSheets( _
                "source worksheets") Then

            Set ImportCsvByDate = _
                Worksheets(SheetName)

            Exit Function

        End If

    End If

    Set ws = CreateOrReplaceSheet(SheetName)

    If ws Is Nothing Then Exit Function

    csvData = ReadAllLines(csvFilePath)

    Headers = csvData(0)

    ws.Range("B2").Value = Headers

    For i = LBound(csvData) To UBound(csvData)

        Line = csvData(i)

        '
        ' LastRow was 1 on every iteration, whatever the sheet held: the
        ' search started at row 2, so A1 is the only cell End(xlUp) could
        ' reach. That made the target B1.Offset(1 + i) = B(2 + i), which is
        ' what the line below writes directly. The original is kept until
        ' this path has a caller again and the change can be exercised:
        '
        '        LastRow = _
        '            ws.Range("A2").End(xlUp).Row
        '
        '        ws.Range("B1").Offset( _
        '            LastRow + i, _
        '            0).Value = Line
        '
        ws.Range("B2").Offset(i, 0).Value = Line

    Next i

    LastRow = _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ws.Range("B2:B" & LastRow). _
        TextToColumns _
        Destination:=ws.Range("B2"), _
        DataType:=xlDelimited, _
        Other:=True, _
        OtherChar:=Chr(10), _
        ThousandsSeparator:=",", _
        TrailingMinusNumbers:=True

    Set ImportCsvByDate = ws

End Function

Public Function SourceFileExists( _
    ByVal CheckDate As Date, _
    ByVal FileType As String) As Boolean

    Dim FileSuffix As String
    Dim FilePath As String

    Select Case UCase(FileType)

        Case "POSITIONS"

            FileSuffix = _
                "_Lombard_Loans_ITA_Positions.csv"

        Case "ACCOUNTS"

            FileSuffix = _
                "_Lombard_Loans_ITA_Accounts.csv"

        Case Else

            Err.Raise vbObjectError + 1000, _
                      "SourceFileExists", _
                      "Unknown file type"

    End Select

    FilePath = _
        PathSelection() & _
        GetDateCode(CheckDate) & _
        FileSuffix

    SourceFileExists = _
        (Dir(FilePath) <> "")

End Function

Public Function ResolveAvailableDate( _
    ByVal RequestedDate As Date, _
    ByVal FileType As String, _
    Optional ByVal MaxDaysForward As Long = 10) As Date

    Dim Candidate As Date
    Dim i As Long

    Candidate = RequestedDate

    For i = 0 To MaxDaysForward

        If SourceFileExists(Candidate, FileType) Then

            If Candidate <> RequestedDate Then

                Note _
                    FileType & " file missing" & vbLf & _
                    Format(RequestedDate, "yyyymmdd") & _
                    " replaced by " & _
                    Format(Candidate, "yyyymmdd")

            End If

            ResolveAvailableDate = Candidate
            Exit Function

        End If

        Candidate = Candidate + 1

    Next i

    MissingFiles = _
        MissingFiles & vbCrLf & _
        FileType & ": " & _
        Format(RequestedDate, "yyyymmdd")

    Err.Raise _
        vbObjectError + 2000, _
        "ResolveAvailableDate", _
        "No " & FileType & _
        " file found from " & _
        Format(RequestedDate, "dd/mm/yyyy") & _
        " within +" & MaxDaysForward & " days."

End Function

Public Function ResolveComparisonDate( _
    ByVal RequestedDate As Date, _
    Optional ByVal MaxDaysForward As Long = 10) As Date

    Dim Candidate As Date
    Dim i As Long

    Candidate = RequestedDate

    For i = 0 To MaxDaysForward

        If SourceFileExists(Candidate, "POSITIONS") _
           And SourceFileExists(Candidate, "ACCOUNTS") Then

            If Candidate <> RequestedDate Then

                Note _
                    "Comparison source files missing" & vbLf & _
                    "Requested date: " & _
                    Format(RequestedDate, "dd/mm/yyyy") & vbLf & _
                    "Using next available date: " & _
                    Format(Candidate, "dd/mm/yyyy")

            End If

            ResolveComparisonDate = Candidate
            Exit Function

        End If

        Candidate = Candidate + 1

    Next i

    Fatal _
        "No comparison data found after " & _
        Format(RequestedDate, "dd/mm/yyyy")

End Function


