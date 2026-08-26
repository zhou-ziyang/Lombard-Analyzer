Attribute VB_Name = "ToolsCountryProbe"
Option Explicit

'
' Builds the Country of Risk concentration, Full scope, twice over the same
' RiskExposure staging table: once with worksheet formulas, once with a VBA
' pass written from the rules in AggregateUnifiedRiskStageData, and puts the
' difference between them in a column.
'
' The question it answers is whether the concentration arithmetic can leave
' VBA and live in the sheet.  It reads the staging table and writes one new
' sheet; it changes nothing the weekly report depends on, and it does not
' need WeeklyAnalysisGenerate to be present.
'
' The rules it reproduces, all from that module:
'
'   Country of Risk shows Equity, Corporate Bonds and Certificates.  The
'   other two classes and the Overall subsection are commented out in
'   BuildRiskSubtableVisibility, so they are absent here too.
'
'   A row reaches the geography dimension only if it is an entity exposure:
'   Exposure Type is not "Certificate non-entity component" and Exposure Name
'   does not carry the __CERTIFICATE_NON_ENTITY__| prefix.  Non-entity rows
'   are excluded outright, so the share denominator uses the same filter as
'   the ranked list - unlike the Issuer dimension, where such rows count
'   towards the denominator but never appear in the list.
'
'   Blank geography becomes "Others".
'
'   Ranking is by value descending, ties broken by name ascending, top ten.
'
'   The #NDG on a country row counts that country''s distinct NDGs.  The one
'   on the total row counts the distinct NDGs of the ten countries taken
'   together - a union, not the sum of the ten counts above it.
'
' Full scope is every row, DPM and Non-DPM alike, which is what the report''s
' left-hand table shows.
'
' The ranked table is one formula per asset class.  Nothing is defensive
' about it: a wrong formula shows up as a non-zero difference against the
' VBA column beside it, which is what the sheet is for.
'

Private Const PROBE_SHEET As String = "Country Probe"
Private Const STAGE_TABLE As String = "RiskExposure"

Private Const NON_ENTITY_PREFIX As String = "__CERTIFICATE_NON_ENTITY__|"
Private Const NON_ENTITY_TYPE As String = "Certificate non-entity component"
Private Const OTHER_DIMENSION As String = "Others"
Private Const TOP_COUNT As Long = 10

Private Const BLOCK_HEIGHT As Long = 15
Private Const FIRST_BLOCK_ROW As Long = 8

Public Sub BuildCountryConcentrationProbe()

    Dim StageTable As ListObject
    Dim StageData As Variant
    Dim ws As Worksheet
    Dim Col As Object

    Dim Classes As Variant
    Dim ClassIndex As Long
    Dim BlockRow As Long

    Dim PreviousUpdating As Boolean

    On Error GoTo Failed

    Set StageTable = FindStageTable()

    If StageTable Is Nothing Then

        MsgBox _
            "No """ & STAGE_TABLE & """ table in this workbook." & _
            vbCrLf & vbCrLf & _
            "Run Weekly Analysis once so the staging table exists.", _
            vbExclamation, "Country probe"

        Exit Sub

    End If

    If StageTable.ListRows.Count = 0 Then

        MsgBox STAGE_TABLE & " is empty.", vbExclamation, "Country probe"
        Exit Sub

    End If

    Set Col = ColumnIndexes(StageTable)
    StageData = StageTable.DataBodyRange.Value

    PreviousUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    RemoveLegacyProbeNames

    Set ws = ReplaceSheet(PROBE_SHEET)

    WriteProbeHeader ws, StageTable

    Classes = Array("Equity", "Corporate Bonds", "Certificates")
    BlockRow = FIRST_BLOCK_ROW

    For ClassIndex = LBound(Classes) To UBound(Classes)

        WriteClassBlock ws, BlockRow, CStr(Classes(ClassIndex)), StageData, Col

        BlockRow = BlockRow + BLOCK_HEIGHT

    Next ClassIndex

    WriteSummaryFormulas ws, BlockRow

    ws.Calculate
    ws.Columns("A:K").AutoFit

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = PreviousUpdating

    Exit Sub

Failed:

    Application.ScreenUpdating = True

    MsgBox _
        "Country probe failed:" & vbCrLf & _
        Err.Number & "  " & Err.Description, _
        vbCritical, "Country probe"

End Sub

'
' An earlier version of this probe defined workbook names for the staging
' columns.  Names.Add was what Excel had been rejecting, and the bindings
' live in each formula's LET now, so any left behind are removed.
'
Private Sub RemoveLegacyProbeNames()

    Dim Legacy As Variant
    Dim NameText As Variant

    Legacy = Array("pxGeography", "pxEntity", "pxValue", "pxNDG", "pxClass")

    On Error Resume Next

    For Each NameText In Legacy
        ThisWorkbook.Names(CStr(NameText)).Delete
    Next NameText

End Sub

Private Sub WriteProbeHeader( _
    ByVal ws As Worksheet, _
    ByVal StageTable As ListObject)

    ws.Range("A1").Value = "Country of Risk concentration, Full scope"

    ws.Range("A2").Value = _
        "formulas over " & STAGE_TABLE & " (" & _
        Format$(StageTable.ListRows.Count, "#,##0") & _
        " rows) against a VBA pass over the same table"

    ws.Range("A4").Value = "largest value difference"
    ws.Range("A5").Value = "largest #NDG difference"
    ws.Range("A6").Value = "country lists that differ"

    ws.Range("A1").Font.Bold = True
    ws.Range("A4:A6").Font.Bold = True
    ws.Range("B4").NumberFormat = "#,##0.0000"
    ws.Range("B5:B6").NumberFormat = "0"

End Sub

Private Sub WriteSummaryFormulas( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long)

    Dim Span As String

    Span = CStr(FIRST_BLOCK_ROW) & ":I" & CStr(LastRow)

    ws.Range("B4").Formula2 = _
        "=MAX(IF(ISNUMBER(I" & Span & "),ABS(I" & Span & "),0))"

    Span = CStr(FIRST_BLOCK_ROW) & ":J" & CStr(LastRow)

    ws.Range("B5").Formula2 = _
        "=MAX(IF(ISNUMBER(J" & Span & "),ABS(J" & Span & "),0))"

    ws.Range("B6").Formula2 = _
        "=COUNTIF(K" & CStr(FIRST_BLOCK_ROW) & ":K" & CStr(LastRow) & _
        ",""different"")"

End Sub

Private Sub WriteClassBlock( _
    ByVal ws As Worksheet, _
    ByVal TopRow As Long, _
    ByVal AssetClass As String, _
    ByRef StageData As Variant, _
    ByVal Col As Object)

    Dim RefNames() As String
    Dim RefValues() As Double
    Dim RefNDGs() As Long
    Dim RefCount As Long
    Dim RefTotal As Double
    Dim RefUnionNDG As Long

    Dim FirstRow As Long
    Dim TotalRow As Long
    Dim TotalCell As String
    Dim r As Long
    Dim i As Long

    FirstRow = TopRow + 2
    TotalRow = FirstRow + TOP_COUNT
    TotalCell = "$C$" & CStr(TotalRow + 1)

    ws.Cells(TopRow, 1).Value = AssetClass
    ws.Cells(TopRow, 1).Font.Bold = True

    WriteBlockHeaders ws, TopRow + 1

    WriteRankedTable ws, FirstRow, AssetClass

    For i = 1 To TOP_COUNT

        r = FirstRow + i - 1

        ws.Cells(r, 1).Value = i

        ws.Cells(r, 5).Formula2 = _
            "=IF($B" & CStr(r) & "="""",""""," & _
            "C" & CStr(r) & "/" & TotalCell & ")"

        ws.Cells(r, 9).Formula2 = _
            "=IF(OR($B" & CStr(r) & "="""",$F" & CStr(r) & "=""""),""""," & _
            "C" & CStr(r) & "-G" & CStr(r) & ")"

        ws.Cells(r, 10).Formula2 = _
            "=IF(OR($B" & CStr(r) & "="""",$F" & CStr(r) & "=""""),""""," & _
            "D" & CStr(r) & "-H" & CStr(r) & ")"

        ws.Cells(r, 11).Formula2 = _
            "=IF($B" & CStr(r) & "=$F" & CStr(r) & ",""same"",""different"")"

    Next i

    '
    ' Total row.  The #NDG here is the distinct NDGs of the ten countries
    ' together, which is why it is not a sum of the column above it.
    '
    ws.Cells(TotalRow, 2).Value = "Top " & CStr(TOP_COUNT)

    ws.Cells(TotalRow, 3).Formula2 = _
        "=SUM(C" & CStr(FirstRow) & ":C" & CStr(TotalRow - 1) & ")"

    ws.Cells(TotalRow, 4).Formula2 = _
        UnionNdgFormula(AssetClass, FirstRow, TotalRow)

    ws.Cells(TotalRow, 5).Formula2 = "=C" & CStr(TotalRow) & "/" & TotalCell

    ws.Cells(TotalRow, 7).Formula2 = _
        "=SUM(G" & CStr(FirstRow) & ":G" & CStr(TotalRow - 1) & ")"

    ws.Cells(TotalRow, 9).Formula2 = _
        "=C" & CStr(TotalRow) & "-G" & CStr(TotalRow)

    ws.Cells(TotalRow, 10).Formula2 = _
        "=D" & CStr(TotalRow) & "-H" & CStr(TotalRow)

    '
    ' Category total, the share denominator: every row of the class that
    ' reached the dimension, ranked or not.
    '
    ws.Cells(TotalRow + 1, 2).Value = "Category total"
    ws.Cells(TotalRow + 1, 3).Formula2 = CategoryTotalFormula(AssetClass)

    ws.Cells(TotalRow + 1, 9).Formula2 = _
        "=C" & CStr(TotalRow + 1) & "-G" & CStr(TotalRow + 1)

    ReferenceTopTen _
        StageData, AssetClass, Col, _
        RefNames, RefValues, RefNDGs, RefCount, RefTotal, RefUnionNDG

    For i = 1 To RefCount

        r = FirstRow + i - 1

        ws.Cells(r, 6).Value = RefNames(i)
        ws.Cells(r, 7).Value = RefValues(i)
        ws.Cells(r, 8).Value = RefNDGs(i)

    Next i

    ws.Cells(TotalRow, 8).Value = RefUnionNDG
    ws.Cells(TotalRow + 1, 7).Value = RefTotal

    FormatBlock ws, FirstRow, TotalRow

End Sub

'
' The whole ranked table in one formula: country, total value and distinct
' NDG count, ordered by value descending with ties broken by name ascending,
' first ten.  GROUPBY takes a LAMBDA where an aggregate is wanted, so the
' distinct count is written where SUM sits rather than assembled from UNIQUE
' and COUNTA a row at a time.
'
' GROUPBY names its value columns in a row of its own - "SUM" and "CUSTOM"
' for these two - and that row is text, so sorting by value descending
' carries it to the top and pushes the tenth country out of the table.  DROP
' takes it off before the sort.
'
Private Sub WriteRankedTable( _
    ByVal ws As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal AssetClass As String)

    ws.Cells(FirstRow, 2).Formula2 = _
        "=LET(" & _
        StageBindings(WantGeography:=True, WantNDG:=True, WantValue:=True) & _
        "k," & KeepExpression(AssetClass) & "," & _
        "g,FILTER(geo,k),v,FILTER(val,k),n,FILTER(ndg,k)," & _
        "a,DROP(GROUPBY(g,HSTACK(v,n)," & _
        "HSTACK(SUM,LAMBDA(x,COUNTA(UNIQUE(x)))),0,0),1)," & _
        "IFERROR(TAKE(SORTBY(a,CHOOSECOLS(a,2),-1,CHOOSECOLS(a,1),1)," & _
        CStr(TOP_COUNT) & "),""""))"

End Sub

'
' The staging table''s columns, bound for a LET.  Each formula asks for the
' measures it goes on to read; cls, typ and nme are what KeepExpression
' needs, so they are always bound.
'
Private Function StageBindings( _
    ByVal WantGeography As Boolean, _
    ByVal WantNDG As Boolean, _
    ByVal WantValue As Boolean) As String

    StageBindings = _
        "cls," & STAGE_TABLE & "[Risk Asset Class]," & _
        "typ," & STAGE_TABLE & "[Exposure Type]," & _
        "nme," & STAGE_TABLE & "[Exposure Name],"

    If WantGeography Then

        StageBindings = StageBindings & _
            "geo,IF(TRIM(" & STAGE_TABLE & "[Geography])=""""," & _
            """" & OTHER_DIMENSION & """,TRIM(" & _
            STAGE_TABLE & "[Geography])),"

    End If

    If WantNDG Then
        StageBindings = StageBindings & "ndg," & STAGE_TABLE & "[NDG],"
    End If

    If WantValue Then

        StageBindings = StageBindings & _
            "val," & STAGE_TABLE & "[Allocated Collateral Value],"

    End If

End Function

'
' The rows of one asset class that reach the geography dimension: entity
' exposures only, which is why the share denominator below uses this same
' expression as the ranked list.
'
Private Function KeepExpression( _
    ByVal AssetClass As String) As String

    KeepExpression = _
        "(cls=""" & AssetClass & """)*" & _
        "(typ<>""" & NON_ENTITY_TYPE & """)*" & _
        "(LEFT(nme," & CStr(Len(NON_ENTITY_PREFIX)) & ")<>""" & _
        NON_ENTITY_PREFIX & """)"

End Function

Private Function UnionNdgFormula( _
    ByVal AssetClass As String, _
    ByVal FirstRow As Long, _
    ByVal TotalRow As Long) As String

    UnionNdgFormula = _
        "=LET(" & _
        StageBindings(WantGeography:=True, WantNDG:=True, WantValue:=False) & _
        "k," & KeepExpression(AssetClass) & "," & _
        "COUNTA(UNIQUE(FILTER(ndg,k*ISNUMBER(MATCH(geo,$B$" & _
        CStr(FirstRow) & ":$B$" & CStr(TotalRow - 1) & ",0))))))"

End Function

Private Function CategoryTotalFormula( _
    ByVal AssetClass As String) As String

    CategoryTotalFormula = _
        "=LET(" & _
        StageBindings(WantGeography:=False, WantNDG:=False, WantValue:=True) & _
        "k," & KeepExpression(AssetClass) & ",SUM(FILTER(val,k,0)))"

End Function

Private Sub WriteBlockHeaders( _
    ByVal ws As Worksheet, _
    ByVal HeaderRow As Long)

    Dim Headers As Variant
    Dim i As Long

    Headers = Array( _
        "#", "Country", "Value (formula)", "#NDG (formula)", _
        "Share (formula)", "Country (VBA)", "Value (VBA)", "#NDG (VBA)", _
        "d Value", "d #NDG", "name")

    For i = LBound(Headers) To UBound(Headers)
        ws.Cells(HeaderRow, i + 1).Value = Headers(i)
    Next i

    ws.Range( _
        ws.Cells(HeaderRow, 1), _
        ws.Cells(HeaderRow, UBound(Headers) + 1)).Font.Bold = True

End Sub

Private Sub FormatBlock( _
    ByVal ws As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal TotalRow As Long)

    ws.Range(ws.Cells(FirstRow, 3), ws.Cells(TotalRow + 1, 3)) _
        .NumberFormat = "#,##0"

    ws.Range(ws.Cells(FirstRow, 7), ws.Cells(TotalRow + 1, 7)) _
        .NumberFormat = "#,##0"

    ws.Range(ws.Cells(FirstRow, 5), ws.Cells(TotalRow, 5)) _
        .NumberFormat = "0.00%"

    ws.Range(ws.Cells(FirstRow, 9), ws.Cells(TotalRow + 1, 9)) _
        .NumberFormat = "#,##0.0000"

End Sub

'
' The oracle: the same rules again, in the language the report is written
' in.  Deliberately independent of the formulas above - if both were derived
' from one shared helper they could agree while both being wrong.
'
Private Sub ReferenceTopTen( _
    ByRef StageData As Variant, _
    ByVal AssetClass As String, _
    ByVal Col As Object, _
    ByRef OutNames() As String, _
    ByRef OutValues() As Double, _
    ByRef OutNDGs() As Long, _
    ByRef OutCount As Long, _
    ByRef OutTotal As Double, _
    ByRef OutUnionNDG As Long)

    Dim Values As Object
    Dim NDGSets As Object
    Dim UnionSet As Object
    Dim NDGSet As Object

    Dim Names() As String
    Dim Totals() As Double

    Dim Country As String
    Dim ExposureName As String
    Dim ExposureType As String
    Dim RowValue As Double
    Dim Item As Variant

    Dim RowNo As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long

    Dim SwapText As String
    Dim SwapValue As Double

    Set Values = CreateObject("Scripting.Dictionary")
    Values.CompareMode = vbTextCompare
    Set NDGSets = CreateObject("Scripting.Dictionary")
    NDGSets.CompareMode = vbTextCompare

    For RowNo = 1 To UBound(StageData, 1)

        If StrComp(Trim$(CStr(StageData(RowNo, Col("Risk Asset Class")))), _
                   AssetClass, vbTextCompare) = 0 Then

            ExposureType = Trim$(CStr(StageData(RowNo, Col("Exposure Type"))))
            ExposureName = CStr(StageData(RowNo, Col("Exposure Name")))

            If StrComp(ExposureType, NON_ENTITY_TYPE, vbTextCompare) <> 0 _
               And Left$(ExposureName, Len(NON_ENTITY_PREFIX)) <> _
                   NON_ENTITY_PREFIX Then

                Country = Trim$(CStr(StageData(RowNo, Col("Geography"))))
                If Country = "" Then Country = OTHER_DIMENSION

                RowValue = _
                    CDbl(StageData(RowNo, Col("Allocated Collateral Value")))

                Values(Country) = Values(Country) + RowValue
                OutTotal = OutTotal + RowValue

                If Not NDGSets.Exists(Country) Then

                    Set NDGSet = CreateObject("Scripting.Dictionary")
                    NDGSet.CompareMode = vbTextCompare
                    NDGSets.Add Country, NDGSet

                End If

                NDGSets(Country)(CStr(StageData(RowNo, Col("NDG")))) = True

            End If

        End If

    Next RowNo

    n = Values.Count
    If n = 0 Then Exit Sub

    ReDim Names(1 To n)
    ReDim Totals(1 To n)

    i = 1
    For Each Item In Values.Keys
        Names(i) = CStr(Item)
        Totals(i) = CDbl(Values(Item))
        i = i + 1
    Next Item

    For i = 1 To n - 1
        For j = i + 1 To n

            If Totals(j) > Totals(i) Or _
               (Totals(j) = Totals(i) And _
                StrComp(Names(j), Names(i), vbTextCompare) < 0) Then

                SwapValue = Totals(i)
                Totals(i) = Totals(j)
                Totals(j) = SwapValue

                SwapText = Names(i)
                Names(i) = Names(j)
                Names(j) = SwapText

            End If

        Next j
    Next i

    OutCount = n
    If OutCount > TOP_COUNT Then OutCount = TOP_COUNT

    ReDim OutNames(1 To OutCount)
    ReDim OutValues(1 To OutCount)
    ReDim OutNDGs(1 To OutCount)

    Set UnionSet = CreateObject("Scripting.Dictionary")
    UnionSet.CompareMode = vbTextCompare

    For i = 1 To OutCount

        OutNames(i) = Names(i)
        OutValues(i) = Totals(i)
        OutNDGs(i) = NDGSets(Names(i)).Count

        For Each Item In NDGSets(Names(i)).Keys
            UnionSet(CStr(Item)) = True
        Next Item

    Next i

    OutUnionNDG = UnionSet.Count

End Sub

'
' Header name to column number, so a reordered staging table does not
' silently read the wrong column.
'
Private Function ColumnIndexes( _
    ByVal StageTable As ListObject) As Object

    Dim Indexes As Object
    Dim Wanted As Variant
    Dim HeaderName As Variant

    Set Indexes = CreateObject("Scripting.Dictionary")
    Indexes.CompareMode = vbTextCompare

    Wanted = Array( _
        "Geography", "Allocated Collateral Value", "NDG", _
        "Risk Asset Class", "Exposure Name", "Exposure Type")

    For Each HeaderName In Wanted
        Indexes(CStr(HeaderName)) = _
            StageTable.ListColumns(CStr(HeaderName)).Index
    Next HeaderName

    Set ColumnIndexes = Indexes

End Function

Private Function FindStageTable() As ListObject

    Dim ws As Worksheet
    Dim Candidate As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each Candidate In ws.ListObjects

            If StrComp(Candidate.Name, STAGE_TABLE, vbTextCompare) = 0 Then
                Set FindStageTable = Candidate
                Exit Function
            End If

        Next Candidate
    Next ws

End Function

Private Function ReplaceSheet( _
    ByVal SheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim PreviousAlerts As Boolean

    PreviousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets(SheetName).Delete
    On Error GoTo 0

    Application.DisplayAlerts = PreviousAlerts

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    ws.Name = SheetName

    Set ReplaceSheet = ws

End Function
