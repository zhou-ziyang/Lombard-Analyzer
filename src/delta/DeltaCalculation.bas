Attribute VB_Name = "DeltaCalculation"
Option Explicit

Public Sub BuildPositionMovements()

    '=========================================================================
    ' Settings
    '=========================================================================

    Dim BasePath As String
    Dim DateCode As String

    Dim AnalysisStartDate As Date
    Dim AnalysisEndDate As Date

    Dim StartSnapshotDate As Date

    '=========================================================================
    ' Snapshot files
    '=========================================================================

    Dim StartSnapshotFile As String
    Dim EndSnapshotFile As String

    Dim StartSnapshotLines As Variant
    Dim EndSnapshotLines As Variant

    Dim HeaderFields As Variant

    '=========================================================================
    ' Column indexes
    '=========================================================================

    Dim idxNDG As Long
    Dim idxCO As Long
    Dim idxISIN As Long
    Dim idxSec As Long
    Dim idxAsset As Long
    Dim idxValue As Long

    '=========================================================================
    ' Position caches
    '=========================================================================

    Dim StartCache As PositionCache
    Dim EndCache As PositionCache

    '=========================================================================
    ' Dictionaries
    '=========================================================================

    Dim DictStartPositions As Object
    Dim DictEndPositions As Object
    Dim DictStartNDGs As Object

    Dim DictFirstSeenDates As Object
    Dim DictLastSeenDates As Object

    Dim DictAssetAggregation As Object

    '=========================================================================
    ' Output
    '=========================================================================

    Dim DeltaData() As Variant
    Dim ClosedData() As Variant

    Dim DeltaCount As Long
    Dim ClosedCount As Long

    Dim DeltaCapacity As Long
    Dim ClosedCapacity As Long

    Dim ColCount As Long

    Dim wsDelta As Worksheet
    Dim wsClosed As Worksheet
    Dim wsAggregate As Worksheet

    '=========================================================================
    ' Row variables
    '=========================================================================

    Dim Arr As Variant

    Dim Key As String
    Dim NDG As String

    Dim ChangeType As String

    Dim Asset As String
    Dim AssetClass As String

    Dim BookingDate As Date
    Dim LastSeenDate As Date

    Dim PositionValue As Double

    '=========================================================================
    ' Loop variables
    '=========================================================================

    Dim i As Long
    Dim j As Long

    Dim k As Variant
    Dim OutputRow As Long

    Application.ScreenUpdating = False

'   On Error GoTo ErrHandler

    '=========================================================================
    ' Read settings
    '=========================================================================

    AnalysisStartDate = _
        Worksheets("Home").Range("AnalysisStartDate").Value

    AnalysisEndDate = _
        Worksheets("Home").Range("AnalysisEndDate").Value

    DateCode = _
        GetDateCode(AnalysisEndDate)

    BasePath = _
        PathSelection()

    '=========================================================================
    ' Resolve snapshot files
    '=========================================================================

    EndSnapshotFile = _
        BasePath & _
        DateCode & _
        "_Lombard_Loans_ITA_Positions.csv"

    If Dir(EndSnapshotFile) = "" Then

        Fatal _
            "Positions file not found:" & _
            vbCrLf & EndSnapshotFile

    End If

    StartSnapshotDate = _
        FindNearestExistingDate( _
            BasePath, _
            AnalysisStartDate, _
            "_Lombard_Loans_ITA_Positions.csv", _
            10)

    If StartSnapshotDate = 0 Then

        Fatal _
            "Start date positions file not found."

    End If

    StartSnapshotFile = _
        BasePath & _
        GetDateCode(StartSnapshotDate) & _
        "_Lombard_Loans_ITA_Positions.csv"

    '=========================================================================
    ' Load snapshot data
    '=========================================================================

    EndSnapshotLines = _
        ReadAllLines(EndSnapshotFile)

    StartSnapshotLines = _
        ReadAllLines(StartSnapshotFile)

    If UBound(EndSnapshotLines) < 0 Then
        Fatal "Positions file is empty:" & vbCrLf & EndSnapshotFile
    End If

    If UBound(StartSnapshotLines) < 0 Then
        Fatal "Positions file is empty:" & vbCrLf & StartSnapshotFile
    End If

    HeaderFields = _
        Split(EndSnapshotLines(0), ";")

    '=========================================================================
    ' Read header indexes
    '=========================================================================

    idxNDG = RequiredHeaderIndex(HeaderFields, "NDG", EndSnapshotFile)
    idxCO = RequiredHeaderIndex(HeaderFields, "CO_FT_GAR", EndSnapshotFile)
    idxISIN = RequiredHeaderIndex(HeaderFields, "ISIN", EndSnapshotFile)
    idxSec = RequiredHeaderIndex(HeaderFields, "Security Name", EndSnapshotFile)

    idxAsset = _
        RequiredHeaderIndex( _
            HeaderFields, _
            "Asset Type / Classification", _
            EndSnapshotFile)

    idxValue = _
        RequiredHeaderIndex( _
            HeaderFields, _
            "Position Value", _
            EndSnapshotFile)

    '=========================================================================
    ' Build caches
    '=========================================================================

    StartCache = _
        LoadPositionCache( _
            StartSnapshotLines, _
            idxNDG, _
            idxCO, _
            idxISIN, _
            idxSec, _
            idxAsset)

    EndCache = _
        LoadPositionCache( _
            EndSnapshotLines, _
            idxNDG, _
            idxCO, _
            idxISIN, _
            idxSec, _
            idxAsset)

    '=========================================================================
    ' Build support dictionaries
    '=========================================================================

'    Set DictFirstSeenDates = _
'        BuildFirstSeenDictionary( _
'            BasePath, _
'            AnalysisStartDate, _
'            AnalysisEndDate)
'
'    Set DictLastSeenDates = _
'        BuildLastSeenDictionary( _
'            BasePath, _
'            AnalysisStartDate, _
'            AnalysisEndDate)

    Set DictFirstSeenDates = _
        CreateObject("Scripting.Dictionary")
    
    Set DictLastSeenDates = _
        CreateObject("Scripting.Dictionary")

    BuildPositionDateDictionaries _
        BasePath, _
        AnalysisStartDate, _
        AnalysisEndDate, _
        DictFirstSeenDates, _
        DictLastSeenDates

    Set DictStartPositions = _
        CreateObject("Scripting.Dictionary")

    Set DictEndPositions = _
        CreateObject("Scripting.Dictionary")

    Set DictStartNDGs = _
        CreateObject("Scripting.Dictionary")

    Set DictAssetAggregation = _
        CreateObject("Scripting.Dictionary")

    '=========================================================================
    ' Initialize output arrays
    '=========================================================================

    ColCount = _
        UBound(HeaderFields) - _
        LBound(HeaderFields) + 5

    '
    ' A file holding only a header row has no data lines.  Keep at least
    ' one slot so the ReDim cannot become (1 To 0); DeltaCount and
    ' ClosedCount decide how much is written back.
    '
    DeltaCapacity = UBound(EndSnapshotLines)
    If DeltaCapacity < 1 Then DeltaCapacity = 1

    ClosedCapacity = UBound(StartSnapshotLines)
    If ClosedCapacity < 1 Then ClosedCapacity = 1

    ReDim DeltaData( _
        1 To DeltaCapacity, _
        1 To ColCount)

    ReDim ClosedData( _
        1 To ClosedCapacity, _
        1 To ColCount)

    '=========================================================================
    ' Build start position lookups
    '=========================================================================

    For i = 1 To StartCache.LineCount

        If Not IsEmpty(StartCache.Data(i)) Then

            Arr = StartCache.Data(i)

            NDG = SafeField(Arr, idxNDG)

            If NDG <> "" Then
                DictStartNDGs(NDG) = True
            End If

            Key = StartCache.Keys(i)

            If Key <> "" Then
                DictStartPositions(Key) = True
            End If

        End If

    Next i

    '=========================================================================
    ' Detect new positions
    '=========================================================================

    DeltaCount = 0

    For i = 1 To EndCache.LineCount

        If Not IsEmpty(EndCache.Data(i)) Then

            Arr = EndCache.Data(i)

            Key = EndCache.Keys(i)

            If Key <> "" Then

                DictEndPositions(Key) = True

                If Not DictStartPositions.Exists(Key) Then

                    NDG = SafeField(Arr, idxNDG)

                    If DictStartNDGs.Exists(NDG) Then
                        ChangeType = "New Position"
                    Else
                        ChangeType = "New NDG"
                    End If

                    Asset = SafeField(Arr, idxAsset)

                    AssetClass = _
                        GetAssetClass(Asset)

                    BookingDate = AnalysisEndDate

                    If DictFirstSeenDates.Exists(Key) Then
                        BookingDate = DictFirstSeenDates(Key)
                    End If

                    PositionValue = _
                        ParseCsvDouble( _
                            SafeField(Arr, idxValue))

                    DeltaCount = DeltaCount + 1

                    DeltaData(DeltaCount, 1) = ChangeType
                    DeltaData(DeltaCount, 2) = AnalysisEndDate
                    DeltaData(DeltaCount, 3) = AssetClass
                    DeltaData(DeltaCount, 4) = BookingDate

                    For j = LBound(Arr) To UBound(Arr)

                        DeltaData(DeltaCount, j + 5) = _
                            Arr(j)

                    Next j

                    If Asset <> "" Then

                        If DictAssetAggregation.Exists(Asset) Then
                            DictAssetAggregation(Asset) = _
                                DictAssetAggregation(Asset) + PositionValue
                        Else
                            DictAssetAggregation.Add Asset, PositionValue
                        End If

                    End If

                End If

            End If

        End If

    Next i

    '=========================================================================
    ' Detect closed positions
    '=========================================================================

    ClosedCount = 0

    For i = 1 To StartCache.LineCount

        If Not IsEmpty(StartCache.Data(i)) Then

            Arr = StartCache.Data(i)

            Key = StartCache.Keys(i)

            If Key <> "" Then

                If Not DictEndPositions.Exists(Key) Then

                    AssetClass = _
                        GetAssetClass( _
                            SafeField(Arr, idxAsset))

                    LastSeenDate = AnalysisEndDate

                    If DictLastSeenDates.Exists(Key) Then
                        LastSeenDate = DictLastSeenDates(Key)
                    End If

                    ClosedCount = ClosedCount + 1

                    ClosedData(ClosedCount, 1) = _
                        "Closed Position"

                    ClosedData(ClosedCount, 2) = _
                        AnalysisEndDate

                    ClosedData(ClosedCount, 3) = _
                        AssetClass

                    ClosedData(ClosedCount, 4) = _
                        LastSeenDate

                    For j = LBound(Arr) To UBound(Arr)

                        ClosedData(ClosedCount, j + 5) = _
                            Arr(j)

                    Next j

                End If

            End If

        End If

    Next i

    '=========================================================================
    ' Write Delta sheet
    '=========================================================================

    Set wsDelta = _
        CreateOrReplaceSheet("Delta_" & DateCode)

    wsDelta.Cells(1, 1).Value = "ChangeType"
    wsDelta.Cells(1, 2).Value = "Snapshot Date"
    wsDelta.Cells(1, 3).Value = "Asset Class"
    wsDelta.Cells(1, 4).Value = "Booking Date"

    For j = LBound(HeaderFields) To UBound(HeaderFields)

        wsDelta.Cells(1, j + 5).Value = _
            HeaderFields(j)

    Next j

    If DeltaCount > 0 Then

        wsDelta.Range("A2") _
            .Resize(DeltaCount, ColCount) _
            .Value = DeltaData

    End If

    '=========================================================================
    ' Write Closed sheet
    '=========================================================================

    Set wsClosed = _
        CreateOrReplaceSheet("Closed_" & DateCode)

    wsClosed.Cells(1, 1).Value = "ChangeType"
    wsClosed.Cells(1, 2).Value = "Snapshot Date"
    wsClosed.Cells(1, 3).Value = "Asset Class"
    wsClosed.Cells(1, 4).Value = "End Date"

    For j = LBound(HeaderFields) To UBound(HeaderFields)

        wsClosed.Cells(1, j + 5).Value = _
            HeaderFields(j)

    Next j

    If ClosedCount > 0 Then

        wsClosed.Range("A2") _
            .Resize(ClosedCount, ColCount) _
            .Value = ClosedData

    End If

    '=========================================================================
    ' Write Aggregate sheet
    '=========================================================================

'    Set wsAggregate = _
'        CreateOrReplaceSheet("Aggregate")
'
'    wsAggregate.Cells(1, 1).Value = _
'        "Asset Type"
'
'    wsAggregate.Cells(1, 2).Value = _
'        "Total Value"
'
'    OutputRow = 2
'
'    For Each k In DictAssetAggregation.Keys
'
'        wsAggregate.Cells(OutputRow, 1).Value = k
'        wsAggregate.Cells(OutputRow, 2).Value = DictAssetAggregation(k)
'
'        OutputRow = OutputRow + 1
'
'    Next k
'
'    wsDelta.Columns.AutoFit
'    wsClosed.Columns.AutoFit
'    wsAggregate.Columns.AutoFit
'
'    Note _
'        "DONE!" & vbCrLf & _
'        "New: " & DeltaCount & vbCrLf & _
'        "Closed: " & ClosedCount

CleanExit:

    ResetExcel

    Exit Sub

ErrHandler:

    ResetExcel

    MsgBox Err.Description, vbCritical

End Sub

Private Function FindNearestExistingDate(BasePath As String, TargetDate As Date, Suffix As String, maxBack As Long) As Date
    Dim i As Long, d As Date
    d = TargetDate
    
    For i = 0 To maxBack
        If Dir(BasePath & Format(d, "yyyymmdd") & Suffix) <> "" Then
            FindNearestExistingDate = d
            Exit Function
        End If
        d = d - 1
    Next i
    
    FindNearestExistingDate = 0
End Function

Public Function ReadAllLines(FilePath As String) As Variant

    Dim FileNumber As Integer
    Dim FileSize As Long
    Dim FileText As String

    FileNumber = FreeFile

    Open FilePath For Binary Access Read As #FileNumber

    FileSize = LOF(FileNumber)

    If FileSize > 0 Then

        FileText = Space$(FileSize)
        Get #FileNumber, , FileText

    End If

    Close #FileNumber

    '
    ' Accept CRLF, LF and CR line endings.  An LF-terminated extract
    ' used to parse as a single line and yield no data at all.
    '
    FileText = Replace(FileText, vbCrLf, vbLf)
    FileText = Replace(FileText, vbCr, vbLf)

    ReadAllLines = Split(FileText, vbLf)

End Function

Public Function FindHeaderIndex(hdr As Variant, name As String) As Long
    Dim i As Long
    For i = LBound(hdr) To UBound(hdr)
        If Trim(hdr(i)) = name Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
    FindHeaderIndex = -1
End Function

Public Function RequiredHeaderIndex( _
    ByRef HeaderFields As Variant, _
    ByVal HeaderName As String, _
    Optional ByVal FileName As String = "") As Long

    RequiredHeaderIndex = _
        FindHeaderIndex(HeaderFields, HeaderName)

    If RequiredHeaderIndex = -1 Then

        If FileName = "" Then

            Fatal "Header not found: " & HeaderName

        Else

            Fatal _
                "Header not found: " & HeaderName & _
                vbCrLf & FileName

        End If

    End If

End Function

Public Sub BuildPositionDateDictionaries( _
    ByVal BasePath As String, _
    ByVal AnalysisStartDate As Date, _
    ByVal AnalysisEndDate As Date, _
    ByRef DictFirstSeenDates As Object, _
    ByRef DictLastSeenDates As Object)

    Dim Files As Variant
    Dim FileName As String
    Dim FileDate As Date

    Dim Lines As Variant
    Dim HeaderFields As Variant
    Dim Arr As Variant

    Dim idxNDG As Long
    Dim idxCO As Long
    Dim idxISIN As Long
    Dim idxSec As Long
    Dim idxAsset As Long

    Dim Key As String

    Dim i As Long
    Dim F As Long

    Files = _
        GetSortedPositionFiles( _
            BasePath, _
            AnalysisStartDate, _
            AnalysisEndDate)

    If IsEmpty(Files) Then Exit Sub

    For F = LBound(Files) To UBound(Files)

        FileName = Files(F)

        FileDate = DateSerial( _
            CLng(Left$(FileName, 4)), _
            CLng(Mid$(FileName, 5, 2)), _
            CLng(Mid$(FileName, 7, 2)))

        Lines = _
            ReadAllLines( _
                BasePath & FileName)

        If UBound(Lines) < 1 Then GoTo NextPositionFile

        HeaderFields = _
            Split(Lines(0), ";")

        idxNDG = _
            RequiredHeaderIndex(HeaderFields, "NDG", FileName)

        idxCO = _
            RequiredHeaderIndex(HeaderFields, "CO_FT_GAR", FileName)

        idxISIN = _
            RequiredHeaderIndex(HeaderFields, "ISIN", FileName)

        idxSec = _
            RequiredHeaderIndex(HeaderFields, "Security Name", FileName)

        idxAsset = _
            RequiredHeaderIndex( _
                HeaderFields, _
                "Asset Type / Classification", _
                FileName)
        For i = 1 To UBound(Lines)

            If Trim$(Lines(i)) <> "" Then

                Key = BuildPositionKeyFromLine( _
                          Lines(i), _
                          idxNDG, _
                          idxCO, _
                          idxISIN, _
                          idxSec, _
                          idxAsset)

                If Key <> "" Then
                
                    If Not DictFirstSeenDates.Exists(Key) Then

                        DictFirstSeenDates.Add _
                            Key, _
                            FileDate

                    End If

                    DictLastSeenDates(Key) = _
                        FileDate
                        
                End If

            End If

        Next i

NextPositionFile:
    Next F

End Sub

Public Function GetSortedPositionFiles( _
    ByVal BasePath As String, _
    ByVal AnalysisStartDate As Date, _
    ByVal AnalysisAnalysisEndDate As Date) As Variant

    Dim Files As Collection
    Dim FileArray() As String

    Dim FileName As String
    Dim FileDate As Date

    Dim i As Long
    Dim j As Long

    Dim Temp As String

    Set Files = New Collection

    If Right$(BasePath, 1) <> "\" Then
        BasePath = BasePath & "\"
    End If

    FileName = Dir( _
        BasePath & _
        "*_Lombard_Loans_ITA_Positions.csv")

    Do While FileName <> ""

        If IsNumeric(Left$(FileName, 8)) Then

            FileDate = DateSerial( _
                CLng(Left$(FileName, 4)), _
                CLng(Mid$(FileName, 5, 2)), _
                CLng(Mid$(FileName, 7, 2)))

            If DateValue(FileDate) >= DateValue(AnalysisStartDate) _
            And DateValue(FileDate) <= DateValue(AnalysisAnalysisEndDate) Then

                Files.Add FileName

            End If

        End If

        FileName = Dir

    Loop

    If Files.Count = 0 Then Exit Function

    ReDim FileArray(1 To Files.Count)

    For i = 1 To Files.Count

        FileArray(i) = Files(i)

    Next i

    'Sort ascending

    For i = 1 To UBound(FileArray) - 1

        For j = i + 1 To UBound(FileArray)

            If FileArray(i) > FileArray(j) Then

                Temp = FileArray(i)
                FileArray(i) = FileArray(j)
                FileArray(j) = Temp

            End If

        Next j

    Next i

    GetSortedPositionFiles = FileArray

End Function

'Public Function GetEffectiveMonths( _
'    ByVal FirstSeenDate As Date, _
'    ByVal ReportDate As Date) As Long
'
'    GetEffectiveMonths = _
'        Month(ReportDate) - _
'        Month(FirstSeenDate) + 1
'
'End Function

Private Function BuildPositionKeyFromLine( _
    ByVal Line As String, _
    ByVal idxNDG As Long, _
    ByVal idxCO As Long, _
    ByVal idxISIN As Long, _
    ByVal idxSec As Long, _
    ByVal idxAsset As Long) As String

    Dim CurrentField As Long

    Dim StartPos As Long
    Dim EndPos As Long

    Dim FieldValue As String

    Dim NDG As String
    Dim CO As String
    Dim ISIN As String
    Dim SecurityName As String
    Dim Asset As String

    Dim InstrumentId As String

    StartPos = 1
    CurrentField = 0

    Do

        EndPos = InStr(StartPos, Line, ";")

        If EndPos > 0 Then
            FieldValue = Trim$(Mid$(Line, _
                                    StartPos, _
                                    EndPos - StartPos))
        Else
            FieldValue = Trim$(Mid$(Line, StartPos))
        End If

        Select Case CurrentField

            Case idxNDG
                NDG = FieldValue

            Case idxCO
                CO = FieldValue

            Case idxISIN
                ISIN = FieldValue

            Case idxSec
                SecurityName = FieldValue

            Case idxAsset
                Asset = FieldValue

        End Select

        If CurrentField >= idxNDG _
        And CurrentField >= idxCO _
        And CurrentField >= idxISIN _
        And CurrentField >= idxSec _
        And CurrentField >= idxAsset Then

            Exit Do

        End If

        If EndPos = 0 Then Exit Do

        StartPos = EndPos + 1
        CurrentField = CurrentField + 1

    Loop

    If NDG = "" Then Exit Function

    If ISIN <> "" Then
        InstrumentId = ISIN
    Else
        InstrumentId = SecurityName
    End If

    BuildPositionKeyFromLine = _
        NDG & "|" & _
        CO & "|" & _
        InstrumentId & "|" & _
        Asset

End Function
