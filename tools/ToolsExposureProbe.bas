Attribute VB_Name = "ToolsExposureProbe"
Option Explicit

'
' Rebuilds the whole exposure concentration section from the RiskExposure
' staging table with worksheet formulas, alongside a VBA pass written from
' the rules in AggregateUnifiedRiskStageData, and puts the difference between
' them in a column.
'
' Twenty-two subtables: three dimensions, the asset classes
' BuildRiskSubtableVisibility leaves visible in each, and both account
' scopes.  It reads the staging table and writes one new sheet; it changes
' nothing the weekly report depends on, and it does not need
' WeeklyAnalysisGenerate to be present.
'
' The rules it reproduces, all from that module:
'
'   Issuer shows Equity, Corporate Bonds, Sovereign Bonds, Funds and
'   Certificates.  Country of Risk and Sector show Equity, Corporate Bonds
'   and Certificates.  The Overall subsection is commented out in all three.
'
'   Issuer aggregates every row of a class.  A row that is not an entity
'   exposure - Exposure Type is "Certificate non-entity component", or
'   Exposure Name already carries the __CERTIFICATE_NON_ENTITY__| prefix -
'   has that prefix on its aggregation name, which keeps it out of the ranked
'   list but not out of the share denominator.  So for Issuer the denominator
'   is larger than the sum of what could ever be ranked.
'
'   Country of Risk and Sector aggregate entity rows only, so there the
'   denominator uses the same filter as the ranked list.
'
'   Blank geography or sector becomes "Others".
'
'   Ranking is by value descending, ties broken by name ascending, top ten.
'
'   The #NDG on a row counts that name''s distinct NDGs.  The one on the
'   total row counts the distinct NDGs of the ten taken together - a union,
'   not the sum of the ten counts above it.
'
'   Full is every row.  Excl. DPM is the rows whose Account Scope is
'   Non-DPM; the report builds Full by merging the two, which comes to the
'   same thing.
'
' The ranked table is one formula per subtable.  Nothing is defensive about
' it: a wrong formula shows up as a non-zero difference against the VBA
' column beside it, which is what the sheet is for.
'

Private Const PROBE_SHEET As String = "Exposure Probe"
Private Const STAGE_TABLE As String = "RiskExposure"

Private Const NON_ENTITY_PREFIX As String = "__CERTIFICATE_NON_ENTITY__|"
Private Const NON_ENTITY_TYPE As String = "Certificate non-entity component"
Private Const OTHER_DIMENSION As String = "Others"
Private Const NON_DPM_SCOPE As String = "Non-DPM"
Private Const TOP_COUNT As Long = 10

Private Const BLOCK_HEIGHT As Long = 15
Private Const FIRST_BLOCK_ROW As Long = 8

Public Sub BuildExposureProbe()

    Dim StageTable As ListObject
    Dim StageData As Variant
    Dim ws As Worksheet
    Dim Col As Object

    Dim DimensionSpec As Variant
    Dim ClassLabel As Variant
    Dim ScopeSpec As Variant

    Dim BlockRow As Long
    Dim BlockCount As Long

    Dim PreviousUpdating As Boolean
    Dim PreviousCalculation As XlCalculation

    On Error GoTo Failed

    Set StageTable = FindStageTable()

    If StageTable Is Nothing Then

        MsgBox _
            "No """ & STAGE_TABLE & """ table in this workbook." & _
            vbCrLf & vbCrLf & _
            "Run Weekly Analysis once so the staging table exists.", _
            vbExclamation, "Exposure probe"

        Exit Sub

    End If

    If StageTable.ListRows.Count = 0 Then

        MsgBox STAGE_TABLE & " is empty.", vbExclamation, "Exposure probe"
        Exit Sub

    End If

    Set Col = ColumnIndexes(StageTable)
    StageData = StageTable.DataBodyRange.Value

    PreviousUpdating = Application.ScreenUpdating
    PreviousCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ReplaceSheet(PROBE_SHEET)

    BlockRow = FIRST_BLOCK_ROW

    For Each DimensionSpec In Dimensions()

        For Each ScopeSpec In Scopes()

            For Each ClassLabel In VisibleClasses(CStr(DimensionSpec(0)))

                WriteBlock _
                    ws, _
                    BlockRow, _
                    DimensionSpec, _
                    ClassSpec(CStr(ClassLabel)), _
                    ScopeSpec, _
                    StageData, _
                    Col

                BlockRow = BlockRow + BLOCK_HEIGHT
                BlockCount = BlockCount + 1

            Next ClassLabel

        Next ScopeSpec

    Next DimensionSpec

    WriteProbeHeader ws, StageTable, BlockCount, BlockRow

    ws.Calculate
    Application.Calculation = PreviousCalculation

    ws.Columns("A:K").AutoFit
    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = PreviousUpdating

    Exit Sub

Failed:

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox _
        "Exposure probe failed:" & vbCrLf & _
        Err.Number & "  " & Err.Description, _
        vbCritical, "Exposure probe"

End Sub

'
' Key, the staging column it groups by, the LET name that column is bound to,
' and whether its share denominator is restricted to entity rows.  Issuer is
' the one that is not: a certificate component with no resolved entity counts
' towards the total the ranked names are measured against.
'
Private Function Dimensions() As Variant

    Dimensions = Array( _
        Array("Issuer", "Exposure Name", "nme", False), _
        Array("Country of Risk", "Geography", "geo", True), _
        Array("Sector", "Sector", "sec", True))

End Function

Private Function Scopes() As Variant

    Scopes = Array( _
        Array("Full", False), _
        Array("Excl. DPM", True))

End Function

'
' Mirrors BuildRiskSubtableVisibility.  Change that and this has to follow.
'
Private Function VisibleClasses( _
    ByVal DimensionKey As String) As Variant

    If DimensionKey = "Issuer" Then

        VisibleClasses = Array( _
            "Equity", "Corporate Bonds", "Sovereign Bonds", _
            "Funds", "Certificates")

    Else

        VisibleClasses = Array("Equity", "Corporate Bonds", "Certificates")

    End If

End Function

'
' Element 0 labels the block; the rest are the Risk Asset Class values that
' belong to it.  Certificates is written into the staging table as
' "Certificates (Excl. Protected)", and RiskAssetIndexFromClass accepts the
' bare name too, so both are matched.
'
Private Function ClassSpec( _
    ByVal Label As String) As Variant

    If Label = "Certificates" Then

        ClassSpec = Array( _
            "Certificates", "Certificates", "Certificates (Excl. Protected)")

    Else

        ClassSpec = Array(Label, Label)

    End If

End Function

Private Sub WriteProbeHeader( _
    ByVal ws As Worksheet, _
    ByVal StageTable As ListObject, _
    ByVal BlockCount As Long, _
    ByVal LastRow As Long)

    Dim Span As String

    ws.Range("A1").Value = "Exposure concentration, rebuilt from " & STAGE_TABLE

    ws.Range("A2").Value = _
        CStr(BlockCount) & " subtables over " & _
        Format$(StageTable.ListRows.Count, "#,##0") & _
        " staging rows, formulas against a VBA pass over the same table"

    ws.Range("A4").Value = "largest value difference"
    ws.Range("A5").Value = "largest #NDG difference"
    ws.Range("A6").Value = "rows where the names differ"

    Span = CStr(FIRST_BLOCK_ROW) & ":I" & CStr(LastRow)

    ws.Range("B4").Formula2 = _
        "=MAX(IF(ISNUMBER(I" & Span & "),ABS(I" & Span & "),0))"

    Span = CStr(FIRST_BLOCK_ROW) & ":J" & CStr(LastRow)

    ws.Range("B5").Formula2 = _
        "=MAX(IF(ISNUMBER(J" & Span & "),ABS(J" & Span & "),0))"

    ws.Range("B6").Formula2 = _
        "=COUNTIF(K" & CStr(FIRST_BLOCK_ROW) & ":K" & CStr(LastRow) & _
        ",""different"")"

    ws.Range("A1").Font.Bold = True
    ws.Range("A4:A6").Font.Bold = True
    ws.Range("B4").NumberFormat = "#,##0.0000"
    ws.Range("B5:B6").NumberFormat = "0"

End Sub

Private Sub WriteBlock( _
    ByVal ws As Worksheet, _
    ByVal TopRow As Long, _
    ByVal DimensionSpec As Variant, _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant, _
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

    ws.Cells(TopRow, 1).Value = _
        CStr(DimensionSpec(0)) & "  -  " & CStr(ClassSpecValue(0)) & _
        "  -  " & CStr(ScopeSpec(0))

    ws.Cells(TopRow, 1).Font.Bold = True

    WriteBlockHeaders ws, TopRow + 1, CStr(DimensionSpec(0))

    ws.Cells(FirstRow, 2).Formula2 = _
        RankedTableFormula(DimensionSpec, ClassSpecValue, ScopeSpec)

    For i = 1 To TOP_COUNT

        r = FirstRow + i - 1

        ws.Cells(r, 1).Value = i

        ws.Cells(r, 9).Formula2 = _
            "=IF(OR($B" & CStr(r) & "="""",$F" & CStr(r) & "=""""),""""," & _
            "C" & CStr(r) & "-G" & CStr(r) & ")"

        ws.Cells(r, 10).Formula2 = _
            "=IF(OR($B" & CStr(r) & "="""",$F" & CStr(r) & "=""""),""""," & _
            "E" & CStr(r) & "-H" & CStr(r) & ")"

        ws.Cells(r, 11).Formula2 = _
            "=IF($B" & CStr(r) & "=$F" & CStr(r) & ",""same"",""different"")"

    Next i

    '
    ' Total row.  The #NDG here is the distinct NDGs of the ten names
    ' together, which is why it is not a sum of the column above it.
    '
    ws.Cells(TotalRow, 2).Value = "Top " & CStr(TOP_COUNT)

    ws.Cells(TotalRow, 3).Formula2 = _
        "=SUM(C" & CStr(FirstRow) & ":C" & CStr(TotalRow - 1) & ")"

    ws.Cells(TotalRow, 4).Formula2 = "=C" & CStr(TotalRow) & "/" & TotalCell

    ws.Cells(TotalRow, 5).Formula2 = _
        UnionNdgFormula( _
            DimensionSpec, ClassSpecValue, ScopeSpec, FirstRow, TotalRow)

    ws.Cells(TotalRow, 7).Formula2 = _
        "=SUM(G" & CStr(FirstRow) & ":G" & CStr(TotalRow - 1) & ")"

    ws.Cells(TotalRow, 9).Formula2 = _
        "=C" & CStr(TotalRow) & "-G" & CStr(TotalRow)

    ws.Cells(TotalRow, 10).Formula2 = _
        "=E" & CStr(TotalRow) & "-H" & CStr(TotalRow)

    '
    ' Category total, the share denominator.  For Issuer this includes the
    ' rows that can never be ranked.
    '
    ws.Cells(TotalRow + 1, 2).Value = "Category total"

    ws.Cells(TotalRow + 1, 3).Formula2 = _
        CategoryTotalFormula(DimensionSpec, ClassSpecValue, ScopeSpec)

    ws.Cells(TotalRow + 1, 9).Formula2 = _
        "=C" & CStr(TotalRow + 1) & "-G" & CStr(TotalRow + 1)

    ReferenceTopTen _
        StageData, DimensionSpec, ClassSpecValue, ScopeSpec, Col, _
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
' The whole ranked table in one formula: name, total value and distinct NDG
' count, ordered by value descending with ties broken by name ascending,
' first ten.  GROUPBY takes a LAMBDA where an aggregate is wanted, so the
' distinct count is written where SUM sits rather than assembled from UNIQUE
' and COUNTA a row at a time.
'
' GROUPBY names its value columns in a row of its own - "SUM" and "CUSTOM"
' for these two - and that row is text, so sorting by value descending
' carries it to the top and pushes the tenth name out of the table.  DROP
' takes it off before the sort.
'
Private Function RankedTableFormula( _
    ByVal DimensionSpec As Variant, _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant) As String

    RankedTableFormula = _
        "=LET(" & _
        Bindings( _
            DimensionSpec, _
            "cls,typ,nme,dim,ndg,val" & ScopeBinding(ScopeSpec)) & _
        "k," & KeepExpression(ClassSpecValue, ScopeSpec, True) & "," & _
        "g,FILTER(" & CStr(DimensionSpec(2)) & ",k)," & _
        "v,FILTER(val,k),n,FILTER(ndg,k)," & _
        "a,DROP(GROUPBY(g,HSTACK(v,n)," & _
        "HSTACK(SUM,LAMBDA(x,COUNTA(UNIQUE(x)))),0,0),1)," & _
        "s,TAKE(SORTBY(a,CHOOSECOLS(a,2),-1,CHOOSECOLS(a,1),1)," & _
        CStr(TOP_COUNT) & ")," & _
        "d," & DenominatorExpression(ClassSpecValue, ScopeSpec, DimensionSpec) & _
        ",IFERROR(HSTACK(CHOOSECOLS(s,1),CHOOSECOLS(s,2)," & _
        "CHOOSECOLS(s,2)/d,CHOOSECOLS(s,3)),""""))"

End Function

'
' The share denominator, inside the same formula as the ranked table so the
' percentage does not depend on a cell elsewhere.
'
' For Country of Risk and Sector it is the total of the rows that were
' ranked, which is SUM of the values already filtered.  For Issuer it is
' larger: rows that are not entity exposures were aggregated under a
' prefixed name, so they count towards the total but can never appear in
' the list.
'
Private Function DenominatorExpression( _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant, _
    ByVal DimensionSpec As Variant) As String

    If CBool(DimensionSpec(3)) Then

        DenominatorExpression = "SUM(v)"

    Else

        DenominatorExpression = _
            "SUM(FILTER(val," & _
            KeepExpression(ClassSpecValue, ScopeSpec, False) & ",0))"

    End If

End Function

Private Function UnionNdgFormula( _
    ByVal DimensionSpec As Variant, _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant, _
    ByVal FirstRow As Long, _
    ByVal TotalRow As Long) As String

    UnionNdgFormula = _
        "=LET(" & _
        Bindings(DimensionSpec, "cls,typ,nme,dim,ndg" & ScopeBinding(ScopeSpec)) & _
        "k," & KeepExpression(ClassSpecValue, ScopeSpec, True) & "," & _
        "COUNTA(UNIQUE(FILTER(ndg,k*ISNUMBER(MATCH(" & _
        CStr(DimensionSpec(2)) & ",$B$" & CStr(FirstRow) & ":$B$" & _
        CStr(TotalRow - 1) & ",0))))))"

End Function

Private Function CategoryTotalFormula( _
    ByVal DimensionSpec As Variant, _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant) As String

    Dim EntityOnly As Boolean

    EntityOnly = CBool(DimensionSpec(3))

    CategoryTotalFormula = _
        "=LET(" & _
        Bindings( _
            DimensionSpec, _
            IIf(EntityOnly, "cls,typ,nme,val", "cls,val") & _
            ScopeBinding(ScopeSpec)) & _
        "k," & KeepExpression(ClassSpecValue, ScopeSpec, EntityOnly) & "," & _
        "SUM(FILTER(val,k,0)))"

End Function

'
' The LET preamble, emitting exactly the bindings the caller says its
' calculation reads.  A LET that declares a name nothing reads is not worth
' finding out about the hard way, and the denominators read fewer of these
' than the ranked table does: no dimension column, and for Issuer no entity
' test either.
'
' "dim" stands for the dimension's own column.  For Issuer that is nme,
' which is already in the list, so it is emitted once.
'
Private Function Bindings( _
    ByVal DimensionSpec As Variant, _
    ByVal Wanted As String) As String

    Dim DimensionName As String
    Dim WantsDimension As Boolean

    DimensionName = CStr(DimensionSpec(2))
    WantsDimension = _
        (InStr(1, Wanted, "dim") > 0) And (DimensionName <> "nme")

    If InStr(1, Wanted, "cls") > 0 Then
        Bindings = Bindings & "cls," & StageColumn("Risk Asset Class") & ","
    End If

    If InStr(1, Wanted, "typ") > 0 Then
        Bindings = Bindings & "typ," & StageColumn("Exposure Type") & ","
    End If

    If InStr(1, Wanted, "nme") > 0 Then
        Bindings = Bindings & "nme," & StageColumn("Exposure Name") & ","
    End If

    If WantsDimension Then

        Bindings = Bindings & _
            DimensionName & ",IF(TRIM(" & _
            StageColumn(CStr(DimensionSpec(1))) & ")=""""," & _
            """" & OTHER_DIMENSION & """,TRIM(" & _
            StageColumn(CStr(DimensionSpec(1))) & ")),"

    End If

    If InStr(1, Wanted, "scp") > 0 Then
        Bindings = Bindings & "scp," & StageColumn("Account Scope") & ","
    End If

    If InStr(1, Wanted, "ndg") > 0 Then
        Bindings = Bindings & "ndg," & StageColumn("NDG") & ","
    End If

    If InStr(1, Wanted, "val") > 0 Then

        Bindings = Bindings & _
            "val," & StageColumn("Allocated Collateral Value") & ","

    End If

End Function

Private Function ScopeBinding( _
    ByVal ScopeSpec As Variant) As String

    If CBool(ScopeSpec(1)) Then ScopeBinding = ",scp"

End Function

Private Function StageColumn( _
    ByVal HeaderName As String) As String

    StageColumn = STAGE_TABLE & "[" & HeaderName & "]"

End Function

'
' The rows one subtable aggregates.  EntityOnly is False only for the Issuer
' share denominator, where a certificate component with no resolved entity
' still counts towards the total.
'
Private Function KeepExpression( _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant, _
    ByVal EntityOnly As Boolean) As String

    If UBound(ClassSpecValue) = 1 Then

        KeepExpression = "(cls=""" & CStr(ClassSpecValue(1)) & """)"

    Else

        KeepExpression = _
            "ISNUMBER(MATCH(cls," & ClassNameArray(ClassSpecValue) & ",0))"

    End If

    If EntityOnly Then

        KeepExpression = KeepExpression & _
            "*(typ<>""" & NON_ENTITY_TYPE & """)" & _
            "*(LEFT(nme," & CStr(Len(NON_ENTITY_PREFIX)) & ")<>""" & _
            NON_ENTITY_PREFIX & """)"

    End If

    If CBool(ScopeSpec(1)) Then

        KeepExpression = KeepExpression & _
            "*(TRIM(scp)=""" & NON_DPM_SCOPE & """)"

    End If

End Function

'
' The class values as a formula array constant, so one block can match the
' several strings the staging table uses for it.
'
Private Function ClassNameArray( _
    ByVal ClassSpecValue As Variant) As String

    Dim i As Long

    For i = 1 To UBound(ClassSpecValue)

        If i > 1 Then ClassNameArray = ClassNameArray & ","

        ClassNameArray = _
            ClassNameArray & """" & CStr(ClassSpecValue(i)) & """"

    Next i

    ClassNameArray = "{" & ClassNameArray & "}"

End Function

Private Sub WriteBlockHeaders( _
    ByVal ws As Worksheet, _
    ByVal HeaderRow As Long, _
    ByVal DimensionLabel As String)

    Dim Headers As Variant
    Dim i As Long

    Headers = Array( _
        "#", DimensionLabel, "Value (formula)", "Share (formula)", _
        "#NDG (formula)", DimensionLabel & " (VBA)", "Value (VBA)", _
        "#NDG (VBA)", "d Value", "d #NDG", "name")

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

    ws.Range(ws.Cells(FirstRow, 4), ws.Cells(TotalRow, 4)) _
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
    ByVal DimensionSpec As Variant, _
    ByVal ClassSpecValue As Variant, _
    ByVal ScopeSpec As Variant, _
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

    Dim GroupName As String
    Dim ExposureName As String
    Dim ExposureType As String
    Dim RowValue As Double
    Dim IsEntity As Boolean
    Dim EntityOnlyTotal As Boolean
    Dim Item As Variant

    Dim RowNo As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim Best As Long

    Dim SwapText As String
    Dim SwapValue As Double

    EntityOnlyTotal = CBool(DimensionSpec(3))

    Set Values = CreateObject("Scripting.Dictionary")
    Values.CompareMode = vbTextCompare
    Set NDGSets = CreateObject("Scripting.Dictionary")
    NDGSets.CompareMode = vbTextCompare

    For RowNo = 1 To UBound(StageData, 1)

        If ClassMatches( _
               ClassSpecValue, _
               CStr(StageData(RowNo, Col("Risk Asset Class")))) _
           And ScopeMatches( _
               ScopeSpec, _
               CStr(StageData(RowNo, Col("Account Scope")))) Then

            ExposureType = Trim$(CStr(StageData(RowNo, Col("Exposure Type"))))
            ExposureName = CStr(StageData(RowNo, Col("Exposure Name")))

            IsEntity = _
                StrComp(ExposureType, NON_ENTITY_TYPE, vbTextCompare) <> 0 _
                And Left$(ExposureName, Len(NON_ENTITY_PREFIX)) <> _
                    NON_ENTITY_PREFIX

            RowValue = _
                CDbl(StageData(RowNo, Col("Allocated Collateral Value")))

            If IsEntity Or Not EntityOnlyTotal Then
                OutTotal = OutTotal + RowValue
            End If

            If IsEntity Then

                '
                ' Issuer groups by the Exposure Name as it stands; the
                ' formula binds that column directly.  Geography and Sector
                ' are trimmed and blanks become "Others", which is what
                ' their binding does.  A row with no Exposure Name never
                ' reaches the staging table - AppendRiskStageRow drops it -
                ' so the difference only matters for keeping the two sides
                ' honest with each other.
                '
                GroupName = _
                    CStr(StageData(RowNo, Col(CStr(DimensionSpec(1)))))

                If CStr(DimensionSpec(2)) <> "nme" Then

                    GroupName = Trim$(GroupName)
                    If GroupName = "" Then GroupName = OTHER_DIMENSION

                End If

                Values(GroupName) = Values(GroupName) + RowValue

                If Not NDGSets.Exists(GroupName) Then

                    Set NDGSet = CreateObject("Scripting.Dictionary")
                    NDGSet.CompareMode = vbTextCompare
                    NDGSets.Add GroupName, NDGSet

                End If

                NDGSets(GroupName)(CStr(StageData(RowNo, Col("NDG")))) = True

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

    OutCount = n
    If OutCount > TOP_COUNT Then OutCount = TOP_COUNT

    '
    ' Only the first ten are ever read, so this selects them instead of
    ' ordering the whole set: one scan of what is left per place, ten scans
    ' in total.  Sorting all of it costs n squared, and n here is the number
    ' of distinct names - a few dozen countries or sectors, but thousands of
    ' issuers.
    '
    ' The comparison is the one WriteTopExposureGroup makes: value
    ' descending, name ascending on a tie.  The report sorts in full there,
    ' for the same top ten.
    '
    For i = 1 To OutCount

        Best = i

        For j = i + 1 To n

            If Totals(j) > Totals(Best) Or _
               (Totals(j) = Totals(Best) And _
                StrComp(Names(j), Names(Best), vbTextCompare) < 0) Then

                Best = j

            End If

        Next j

        If Best <> i Then

            SwapValue = Totals(i)
            Totals(i) = Totals(Best)
            Totals(Best) = SwapValue

            SwapText = Names(i)
            Names(i) = Names(Best)
            Names(Best) = SwapText

        End If

    Next i

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

Private Function ClassMatches( _
    ByVal ClassSpecValue As Variant, _
    ByVal RiskAssetClass As String) As Boolean

    Dim i As Long

    For i = 1 To UBound(ClassSpecValue)

        If StrComp(Trim$(RiskAssetClass), _
                   CStr(ClassSpecValue(i)), vbTextCompare) = 0 Then

            ClassMatches = True
            Exit Function

        End If

    Next i

End Function

Private Function ScopeMatches( _
    ByVal ScopeSpec As Variant, _
    ByVal AccountScope As String) As Boolean

    If Not CBool(ScopeSpec(1)) Then

        ScopeMatches = True

    Else

        ScopeMatches = _
            StrComp(Trim$(AccountScope), NON_DPM_SCOPE, vbTextCompare) = 0

    End If

End Function

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
        "Geography", "Sector", "Allocated Collateral Value", "NDG", _
        "Risk Asset Class", "Exposure Name", "Exposure Type", "Account Scope")

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
