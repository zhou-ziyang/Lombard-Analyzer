Attribute VB_Name = "WeeklyAnalysisEmail"
Option Explicit

' v49: refers to aggregated accounts in the email body.

Private HtmlFragmentCounter As Long
Private HtmlStyleBlocks As String

' =====================================================================
' COMPARISON FEATURE - declarations
'
' What CreateWeeklyComparisonEmail, at the end of this module, reads.
' VBA wants them above the first procedure; delete them with it.
' =====================================================================

Private Const COMPARISON_SHEET As String = "Weekly Comparison"
Private Const COMPARISON_DATE_NAME As String = "WeeklyCompareDate"

' =====================================================================
' COMPARISON FEATURE - end of declarations
' =====================================================================

Public Sub CreateWeeklyEmail()

    HtmlFragmentCounter = 0
    HtmlStyleBlocks = ""
    
    If Layout.PortfolioRow = 0 Then
        InitializeLayout
    End If

    Dim ws As Worksheet

    Dim OutApp As Object
    Dim OutMail As Object

    Dim HTMLBody As String
    Dim RiskHTML As String

    Dim ReportDateValue As Variant
    Dim ReportDate As Date

    Dim ToAddresses As String
    Dim CcAddresses As String

    Dim WordEditor As Object
    Dim WordRange As Object
    Dim shp As Object

    Set ws = ThisWorkbook.Worksheets("Weekly Analysis")
'   Set ws = ActiveSheet

    If InStr(1, ws.name, _
        "Weekly Analysis", _
        vbTextCompare) = 0 Then

        MsgBox _
            "Current worksheet is not a Weekly Analysis report.", _
            vbExclamation

        Exit Sub

    End If

    ReportDateValue = _
        ThisWorkbook.Worksheets("Home") _
        .Range("WeeklyEndDate").Value

    If Not IsDate(ReportDateValue) Then

        MsgBox _
            "WeeklyEndDate does not contain a valid report date.", _
            vbExclamation

        Exit Sub

    End If

    ReportDate = CDate(ReportDateValue)

    ToAddresses = HomeSetting("EmailTo")
    CcAddresses = HomeSetting("EmailCc")

    If ToAddresses = "" Then

        MsgBox _
            "No recipient is configured." & vbCrLf & vbCrLf & _
            "Add a cell on Home named EmailTo — and EmailCc if you " & _
            "want one — holding the addresses separated by " & _
            "semicolons. The draft opens without a recipient until " & _
            "then.", _
            vbInformation, _
            "Weekly Lombard Analysis"

    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    '
    ' Outer page
    '

    HTMLBody = _
        "<html>" & _
        "<head><style>" & _
        ".email-container,.email-container td,.email-container th{" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif " & _
        "!important;}" & _
".rth-table td,.rth-table th{" & _
"padding:4px 5pt !important;" & _
"white-space:nowrap !important;}" & _
        ".rth-title-row td,.rth-title-row th{" & _
        "font-size:14pt !important;" & _
        "padding-left:2px !important;" & _
        "margin-left:0 !important;" & _
        "text-indent:0 !important;" & _
        "text-align:left !important;}" & _
        ".rth-table{border-collapse:collapse;}" & _
        "</style>[[TABLESTYLES]]</head>" & _
        "<body style='margin:0;" & _
        "padding:30px;" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif;'>"

    '
    ' White container
    '

    HTMLBody = HTMLBody & _
        "<table class='email-container' align='left' " & _
        "width='1200' " & _
        "cellpadding='30' " & _
        "cellspacing='0' " & _
        "style='width:1200px; table-layout:fixed;" & _
        "background-color:#ffffff;" & _
        "border:1px solid #ccc;" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif;'>" & _
        "<tr><td>"

    '
    ' Header
    '

    HTMLBody = HTMLBody & _
        "<div style='" & _
        "font-size:28pt;" & _
        "font-weight:bold;" & _
        "margin-top:5px;" & _
        "margin-bottom:30px;'>" & _
        "Weekly Lombard Analysis" & _
        "</div>"

    '
    ' Intro
    '

    HTMLBody = HTMLBody & _
        "Ciao Rossella,<br><br>"

    HTMLBody = HTMLBody & _
        "Please find below the weekly Lombard loan portfolio analysis " & _
        "as of " & _
        Format(ReportDate, "dd.mm.yyyy") & _
        "." & _
        "<br><br>" & _
        "The report covers the portfolio overview, collateral breakdown, " & _
        "and monthly loan activity, followed by the exposure concentration " & _
        "by name, geography, and sector. Concentration results " & _
        "are shown both for the full portfolio and for the portfolio " & _
        "excl. aggregated accounts." & _
        "<br><br><br>"

    HTMLBody = HTMLBody & _
        "<div style='font-size:22pt;font-weight:bold;" & _
        "margin-top:0;margin-bottom:14px;'>" & _
        "Portfolio Overview &amp; Activity" & _
        "</div>"

    '
    ' Report blocks, in reading order: the active loans with the collateral
    ' breakdown straight under them, then the two movement tables, then
    ' what entered.  Each one is a Layout anchor plus the height and width
    ' of the block that starts there.
    '

    HTMLBody = HTMLBody & _
        BlockHtml(ws, Layout.PortfolioRow, Layout.PortfolioCol, 9, 4) & _
        BlockHtml(ws, Layout.BreakdownRow, Layout.BreakdownCol, 9, 8) & _
        BlockHtml(ws, Layout.NewLoanRow, Layout.NewLoanCol, 3, 4) & _
        BlockHtml(ws, Layout.EndedLoanRow, Layout.EndedLoanCol, 3, 4) & _
        BlockHtml(ws, Layout.EnteredRow, Layout.EnteredCol, 5, 8)

    '
    ' Pie Chart Placeholder
    '

    HTMLBody = HTMLBody & _
        "<br>[[PIECHART]]<br>"

    '
    ' Exposure concentration
    '

    RiskHTML = BuildRiskAnalysisHTML(ws)

    If RiskHTML <> "" Then

        HTMLBody = HTMLBody & _
            "<br><br><br>" & _
            "<div style='font-size:22pt;font-weight:bold;" & _
            "margin-top:0;margin-bottom:14px;'>" & _
            "Exposure Concentration" & _
            "</div>" & _
            RiskHTML

    End If

    HTMLBody = HTMLBody & _
        "<p>Grazie<br>SECF Trading</p>"

    '
    ' Close container
    '

    HTMLBody = HTMLBody & _
        "</td></tr></table>" & _
        "</body></html>"

    HTMLBody = _
        Replace( _
            HTMLBody, _
            "[[TABLESTYLES]]", _
            HtmlStyleBlocks, _
            1, _
            -1, _
            vbBinaryCompare)

    With OutMail

        .To = ToAddresses
        .CC = CcAddresses

        .Subject = _
            "Weekly Lombard Analysis " & _
            Format(Date, "dd/mm/yyyy")

        .HTMLBody = HTMLBody

        .Display

    End With

    '
    ' Insert Pie Chart at Placeholder
    '

    Set WordEditor = _
        OutMail.GetInspector.WordEditor

    Set WordRange = WordEditor.Content

    With WordRange.Find

        .ClearFormatting
        .Text = "[[PIECHART]]"

        If .Execute Then

            WordRange.Text = ""

            ws.ChartObjects("CollateralPie").Chart.CopyPicture _
                Appearance:=xlScreen, _
                Format:=xlPicture

            WordRange.Paste

            Set shp = _
                WordEditor.InlineShapes( _
                    WordEditor.InlineShapes.Count)

            shp.Width = _
                ws.ChartObjects("CollateralPie").Width

'            shp.Range.ParagraphFormat.Alignment = 1

        End If

    End With
    
    

End Sub

'
' Reads a configuration cell from Home by defined name. A name that does
' not exist yet returns an empty string rather than raising, so the report
' still opens while the workbook is being set up.
'
Private Function HomeSetting( _
    ByVal RangeName As String) As String

    Dim RawValue As Variant

    On Error GoTo NotConfigured

    RawValue = _
        ThisWorkbook.Worksheets("Home").Range(RangeName).Value

    If IsError(RawValue) Then GoTo NotConfigured

    HomeSetting = Trim$(CStr(RawValue))

    Exit Function

NotConfigured:

End Function

Private Function BlockHtml( _
    ByVal ws As Worksheet, _
    ByVal TopRow As Long, _
    ByVal LeftCol As Long, _
    ByVal HeightRows As Long, _
    ByVal WidthCols As Long) As String

    BlockHtml = _
        RangeToHTMLFragment( _
            ws.Range( _
                ws.Cells(TopRow, LeftCol), _
                ws.Cells(TopRow + HeightRows, LeftCol + WidthCols)))

End Function

Private Function BuildRiskAnalysisHTML( _
    ByVal ws As Worksheet) As String

    Dim SectionTitles As Variant
    Dim SectionTitle As Variant

    Dim SectionHTML As String
    Dim ResultHTML As String

    SectionTitles = _
        Array( _
            "Name Concentration - Top 10", _
            "Geographic Concentration - Top 10", _
            "Sector Concentration - Top 10")

    For Each SectionTitle In SectionTitles

        SectionHTML = _
            RiskTablePairToHTML( _
                ws, _
                CStr(SectionTitle))

        If SectionHTML <> "" Then

            If ResultHTML <> "" Then

                ResultHTML = _
                    ResultHTML & "<br><br>"

            End If

            ResultHTML = _
                ResultHTML & SectionHTML

        End If

    Next SectionTitle

    BuildRiskAnalysisHTML = ResultHTML

End Function

Private Function RiskTablePairToHTML( _
    ByVal ws As Worksheet, _
    ByVal SectionTitle As String) As String

    Dim FirstRow As Long
    Dim LastRow As Long

    FirstRow = _
        FindReportSectionRow( _
            ws, _
            Layout.RiskCol, _
            SectionTitle)

    If FirstRow = 0 Then Exit Function

    LastRow = _
        LastContiguousRiskRow( _
            ws, _
            FirstRow, _
            Layout.RiskCol, _
            Layout.RiskExSegCol + 4)

    If LastRow <= FirstRow Then Exit Function

    RiskTablePairToHTML = _
        RangeToHTMLFragment( _
            ws.Range( _
                ws.Cells( _
                    FirstRow, _
                    Layout.RiskCol), _
                ws.Cells( _
                    LastRow, _
                    Layout.RiskExSegCol + 4)), _
            0)

End Function

Private Function FindReportSectionRow( _
    ByVal ws As Worksheet, _
    ByVal SearchCol As Long, _
    ByVal SectionTitle As String) As Long

    Dim FoundCell As Range

    Set FoundCell = _
        ws.Columns(SearchCol).Find( _
            What:=SectionTitle, _
            After:=ws.Cells(1, SearchCol), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)

    If Not FoundCell Is Nothing Then

        FindReportSectionRow = FoundCell.Row

    End If

End Function

Private Function LastContiguousRiskRow( _
    ByVal ws As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal FirstCol As Long, _
    ByVal LastCol As Long) As Long

    Dim CurrentRow As Long

    CurrentRow = FirstRow

    Do While CurrentRow <= ws.Rows.Count

        If Application.CountA( _
                ws.Range( _
                    ws.Cells(CurrentRow, FirstCol), _
                    ws.Cells(CurrentRow, LastCol))) = 0 Then

            Exit Do

        End If

        LastContiguousRiskRow = CurrentRow
        CurrentRow = CurrentRow + 1

    Loop

End Function

Private Function RangeToHTMLFragment( _
    ByVal rng As Range, _
    Optional ByVal ExtraWidthPt As Long = 10) As String

    Dim Html As String

    Dim StyleStart As Long
    Dim StyleEnd As Long

    Dim TableStart As Long
    Dim TableEnd As Long

    Dim StyleBlock As String
    Dim TableBlock As String
    Dim CssPrefix As String

    Html = RangeToHTML(rng)

    HtmlFragmentCounter = HtmlFragmentCounter + 1

    CssPrefix = _
        "rth" & CStr(HtmlFragmentCounter) & "-"

    '
    ' Extract style section
    '

    StyleStart = InStr(1, Html, "<style", vbTextCompare)

    If StyleStart > 0 Then

        StyleEnd = InStr( _
            StyleStart, _
            Html, _
            "</style>", _
            vbTextCompare)

        If StyleEnd > 0 Then

            StyleBlock = Mid$( _
                Html, _
                StyleStart, _
                StyleEnd - StyleStart + 8)

        End If

    End If

    '
    ' Extract table section
    '

    TableStart = InStr(1, Html, "<table", vbTextCompare)

    TableEnd = InStrRev( _
        Html, _
        "</table>", _
        , _
        vbTextCompare)

    If TableStart > 0 _
       And TableEnd > 0 Then

        TableBlock = Mid$( _
            Html, _
            TableStart, _
            TableEnd - TableStart + Len("</table>"))

        TableBlock = RemoveWidths(TableBlock)

        TableBlock = _
            FormatTitleRow( _
                TableBlock, _
                4)
            
        If ExtraWidthPt <> 0 Then

            TableBlock = ExpandColWidths( _
                TableBlock, _
                ExtraWidthPt)

        End If

        TableBlock = Replace( _
            TableBlock, _
            "<table ", _
            "<table class='rth-table' ", _
            1, 1, vbTextCompare)

    End If

    '
    ' Excel restarts CSS class names at xl65 for every exported range.
    ' Give each fragment its own namespace before combining them in one email.
    '

    StyleBlock = _
        PrefixExcelCssClasses( _
            StyleBlock, _
            CssPrefix)

    TableBlock = _
        PrefixExcelHtmlClasses( _
            TableBlock, _
            CssPrefix)

    If StyleBlock <> "" Then

        HtmlStyleBlocks = _
            HtmlStyleBlocks & vbCrLf & StyleBlock

    End If

    RangeToHTMLFragment = TableBlock
        
End Function

Private Function FormatTitleRow( _
    ByVal Html As String, _
    ByVal GapPx As Long) As String

    Dim FirstRowStart As Long
    Dim FirstRowEnd As Long
    Dim LastCellEnd As Long
    Dim GapHTML As String

    If Html = "" Or GapPx <= 0 Then

        FormatTitleRow = Html
        Exit Function

    End If

    FirstRowStart = _
        InStr(1, Html, "<tr", vbTextCompare)

    If FirstRowStart = 0 Then

        FormatTitleRow = Html
        Exit Function

    End If

    Html = _
        Left$(Html, FirstRowStart - 1) & _
        "<tr class='rth-title-row'" & _
        Mid$(Html, FirstRowStart + 3)

    FirstRowEnd = _
        InStr( _
            FirstRowStart, _
            Html, _
            "</tr>", _
            vbTextCompare)

    If FirstRowEnd = 0 Then

        FormatTitleRow = Html
        Exit Function

    End If

    LastCellEnd = _
        InStrRev( _
            Html, _
            "</td>", _
            FirstRowEnd, _
            vbTextCompare)

    If LastCellEnd = 0 Then

        LastCellEnd = _
            InStrRev( _
                Html, _
                "</th>", _
                FirstRowEnd, _
                vbTextCompare)

    End If

    If LastCellEnd = 0 Then

        FormatTitleRow = Html
        Exit Function

    End If

    GapHTML = _
        "<div style='height:" & CStr(GapPx) & "px;" & _
        "line-height:" & CStr(GapPx) & "px;" & _
        "font-size:1px;margin:0;padding:0;" & _
        "mso-line-height-rule:exactly;'>&nbsp;</div>"

    FormatTitleRow = _
        Left$(Html, LastCellEnd - 1) & _
        GapHTML & _
        Mid$(Html, LastCellEnd)

End Function

Private Function PrefixExcelCssClasses( _
    ByVal Html As String, _
    ByVal CssPrefix As String) As String

    Dim RE As Object

    If Html = "" Then Exit Function

    Set RE = CreateObject("VBScript.RegExp")

    RE.Global = True
    RE.IgnoreCase = True
    RE.Pattern = "\.xl([0-9]+)\b"

    PrefixExcelCssClasses = _
        RE.Replace( _
            Html, _
            "." & CssPrefix & "xl$1")

End Function

Private Function PrefixExcelHtmlClasses( _
    ByVal Html As String, _
    ByVal CssPrefix As String) As String

    Dim RE As Object

    If Html = "" Then Exit Function

    Set RE = CreateObject("VBScript.RegExp")

    RE.Global = True
    RE.IgnoreCase = True
    RE.Pattern = _
        "(class\s*=\s*[""']?)xl([0-9]+)\b"

    PrefixExcelHtmlClasses = _
        RE.Replace( _
            Html, _
            "$1" & CssPrefix & "xl$2")

End Function

Private Function ExpandColWidths( _
    ByVal Html As String, _
    ByVal ExtraPt As Long) As String

    Dim RE As Object
    Dim Matches As Object
    Dim M As Object

    Dim OldWidth As Long
    Dim NewWidth As Long

    Set RE = CreateObject("VBScript.RegExp")

    RE.Global = True
    RE.IgnoreCase = True

    RE.Pattern = "width:(\d+)pt"

    Set Matches = RE.Execute(Html)

    Dim i As Long

    For i = Matches.Count - 1 To 0 Step -1

        Set M = Matches(i)

        OldWidth = CLng(M.SubMatches(0))

        NewWidth = OldWidth + ExtraPt

        Html = _
            Left$(Html, M.FirstIndex) & _
            "width:" & NewWidth & "pt" & _
            Mid$(Html, M.FirstIndex + M.Length + 1)

    Next i

    ExpandColWidths = Html

End Function

Private Function RemoveWidths( _
    ByVal Html As String) As String

    Dim RE As Object

    Set RE = CreateObject("VBScript.RegExp")

    RE.Global = True
    RE.IgnoreCase = True

    '
    ' width=123
    '
    RE.Pattern = "\swidth=\d+"
    Html = RE.Replace(Html, "")

    '
    ' width:123pt
    '
'    RE.Pattern = "width:\s*[\d\.]+pt;?"
'    Html = RE.Replace(Html, "")

    RemoveWidths = Html

End Function

Private Function RangeToHTML( _
    ByVal rng As Range) As String
' By Ron de Bruin...
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangeToHTML = ts.ReadAll
    ts.Close
    RangeToHTML = Replace(RangeToHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function


' =====================================================================
' COMPARISON FEATURE - start
'
' The email for the Weekly Comparison sheet, bound to the button
' GenerateWeeklyAnalysisComparison draws there: CreateWeeklyEmail's page
' and blocks read from that sheet, the intro naming the date compared to,
' and each table taken at the height it was built - the comparison rows
' make the tables taller, and by how much depends on the dates.  Nothing
' outside the marked blocks depends on this; delete it with them.
' =====================================================================

Public Sub CreateWeeklyComparisonEmail()

    HtmlFragmentCounter = 0
    HtmlStyleBlocks = ""

    InitializeLayout

    Dim ws As Worksheet

    Dim OutApp As Object
    Dim OutMail As Object

    Dim HTMLBody As String
    Dim RiskHTML As String

    Dim ReportDateValue As Variant
    Dim ReportDate As Date

    Dim CompareDateValue As Variant
    Dim CompareDate As Date

    Dim ToAddresses As String
    Dim CcAddresses As String

    Dim WordEditor As Object
    Dim WordRange As Object
    Dim shp As Object

    Dim LeftRow As Long
    Dim MiddleRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COMPARISON_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "There is no " & COMPARISON_SHEET & " sheet yet. " & _
            "Run the Weekly Comparison first.", _
            vbExclamation

        Exit Sub

    End If

    ReportDateValue = _
        ThisWorkbook.Worksheets("Home") _
        .Range("WeeklyEndDate").Value

    If Not IsDate(ReportDateValue) Then

        MsgBox _
            "WeeklyEndDate does not contain a valid report date.", _
            vbExclamation

        Exit Sub

    End If

    ReportDate = CDate(ReportDateValue)

    On Error Resume Next
    CompareDateValue = _
        ThisWorkbook.Worksheets("Home") _
        .Range(COMPARISON_DATE_NAME).Value
    On Error GoTo 0

    If Not IsDate(CompareDateValue) Then

        MsgBox _
            COMPARISON_DATE_NAME & " does not contain a valid date.", _
            vbExclamation

        Exit Sub

    End If

    CompareDate = CDate(CompareDateValue)

    ToAddresses = HomeSetting("EmailTo")
    CcAddresses = HomeSetting("EmailCc")

    If ToAddresses = "" Then

        MsgBox _
            "No recipient is configured." & vbCrLf & vbCrLf & _
            "Add a cell on Home named EmailTo - and EmailCc if you " & _
            "want one - holding the addresses separated by " & _
            "semicolons. The draft opens without a recipient until " & _
            "then.", _
            vbInformation, _
            "Weekly Lombard Analysis"

    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    '
    ' Outer page
    '

    HTMLBody = _
        "<html>" & _
        "<head><style>" & _
        ".email-container,.email-container td,.email-container th{" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif " & _
        "!important;}" & _
        ".rth-table td,.rth-table th{" & _
        "padding:4px 5pt !important;" & _
        "white-space:nowrap !important;}" & _
        ".rth-title-row td,.rth-title-row th{" & _
        "font-size:14pt !important;" & _
        "padding-left:2px !important;" & _
        "margin-left:0 !important;" & _
        "text-indent:0 !important;" & _
        "text-align:left !important;}" & _
        ".rth-table{border-collapse:collapse;}" & _
        "</style>[[TABLESTYLES]]</head>" & _
        "<body style='margin:0;" & _
        "padding:30px;" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif;'>"

    '
    ' White container
    '

    HTMLBody = HTMLBody & _
        "<table class='email-container' align='left' " & _
        "width='1200' " & _
        "cellpadding='30' " & _
        "cellspacing='0' " & _
        "style='width:1200px; table-layout:fixed;" & _
        "background-color:#ffffff;" & _
        "border:1px solid #ccc;" & _
        "font-family:Aptos Display,Aptos,UniCredit,Calibri,sans-serif;'>" & _
        "<tr><td>"

    '
    ' Header
    '

    HTMLBody = HTMLBody & _
        "<div style='" & _
        "font-size:28pt;" & _
        "font-weight:bold;" & _
        "margin-top:5px;" & _
        "margin-bottom:30px;'>" & _
        "Weekly Lombard Analysis" & _
        "</div>"

    '
    ' Intro, naming the report compared to
    '

    HTMLBody = HTMLBody & _
        "Ciao Rossella,<br><br>"

    HTMLBody = HTMLBody & _
        "Please find below the weekly Lombard loan portfolio analysis " & _
        "as of " & _
        Format(ReportDate, "dd.mm.yyyy") & _
        ", compared to the report as of " & _
        Format(CompareDate, "dd.mm.yyyy") & _
        "." & _
        "<br><br>" & _
        "The report covers the portfolio overview, collateral breakdown, " & _
        "and monthly loan activity, followed by the exposure concentration " & _
        "by name, geography, and sector. Concentration results " & _
        "are shown both for the full portfolio and for the portfolio " & _
        "excl. aggregated accounts." & _
        "<br><br><br>"

    HTMLBody = HTMLBody & _
        "<div style='font-size:22pt;font-weight:bold;" & _
        "margin-top:0;margin-bottom:14px;'>" & _
        "Portfolio Overview &amp; Activity" & _
        "</div>"

    '
    ' Report blocks in the weekly email's order - the active loans, the
    ' breakdown, the two movement tables, what entered - each at the
    ' height it was built, read down the sheet's left column and its
    ' middle column in turn.
    '

    LeftRow = Layout.PortfolioRow
    MiddleRow = Layout.BreakdownRow

    HTMLBody = HTMLBody & _
        ComparisonBlockHtml(ws, LeftRow, Layout.PortfolioCol, 4)
    HTMLBody = HTMLBody & _
        ComparisonBlockHtml(ws, MiddleRow, Layout.BreakdownCol, 8)
    HTMLBody = HTMLBody & _
        ComparisonBlockHtml(ws, LeftRow, Layout.PortfolioCol, 4)
    HTMLBody = HTMLBody & _
        ComparisonBlockHtml(ws, LeftRow, Layout.PortfolioCol, 4)
    HTMLBody = HTMLBody & _
        ComparisonBlockHtml(ws, MiddleRow, Layout.BreakdownCol, 8)

    '
    ' Pie Chart Placeholder
    '

    HTMLBody = HTMLBody & _
        "<br>[[PIECHART]]<br>"

    '
    ' Exposure concentration
    '

    RiskHTML = BuildRiskAnalysisHTML(ws)

    If RiskHTML <> "" Then

        HTMLBody = HTMLBody & _
            "<br><br><br>" & _
            "<div style='font-size:22pt;font-weight:bold;" & _
            "margin-top:0;margin-bottom:14px;'>" & _
            "Exposure Concentration" & _
            "</div>" & _
            RiskHTML

    End If

    HTMLBody = HTMLBody & _
        "<p>Grazie<br>SECF Trading</p>"

    '
    ' Close container
    '

    HTMLBody = HTMLBody & _
        "</td></tr></table>" & _
        "</body></html>"

    HTMLBody = _
        Replace( _
            HTMLBody, _
            "[[TABLESTYLES]]", _
            HtmlStyleBlocks, _
            1, _
            -1, _
            vbBinaryCompare)

    With OutMail

        .To = ToAddresses
        .CC = CcAddresses

        .Subject = _
            "Weekly Lombard Analysis " & _
            Format(Date, "dd/mm/yyyy")

        .HTMLBody = HTMLBody

        .Display

    End With

    '
    ' Insert Pie Chart at Placeholder
    '

    Set WordEditor = _
        OutMail.GetInspector.WordEditor

    Set WordRange = WordEditor.Content

    With WordRange.Find

        .ClearFormatting
        .Text = "[[PIECHART]]"

        If .Execute Then

            WordRange.Text = ""

            ws.ChartObjects("CollateralPie").Chart.CopyPicture _
                Appearance:=xlScreen, _
                Format:=xlPicture

            WordRange.Paste

            Set shp = _
                WordEditor.InlineShapes( _
                    WordEditor.InlineShapes.Count)

            shp.Width = _
                ws.ChartObjects("CollateralPie").Width

        End If

    End With

End Sub

'
' One table's HTML: from the first row at or under TopRow with anything
' in it, down to the blank row that ends the table.  Leaves TopRow on the
' row where the next table can start, so the tables of one column are
' read one after the other whatever their heights.
'
Private Function ComparisonBlockHtml( _
    ByVal ws As Worksheet, _
    ByRef TopRow As Long, _
    ByVal LeftCol As Long, _
    ByVal WidthCols As Long) As String

    Dim FirstRow As Long
    Dim LastRow As Long

    FirstRow = TopRow

    Do While ComparisonRowIsBlank(ws, FirstRow, LeftCol, WidthCols)

        FirstRow = FirstRow + 1

        If FirstRow > TopRow + 60 Then Exit Function

    Loop

    LastRow = FirstRow

    Do While Not ComparisonRowIsBlank(ws, LastRow + 1, LeftCol, WidthCols)

        LastRow = LastRow + 1

    Loop

    ComparisonBlockHtml = _
        BlockHtml(ws, FirstRow, LeftCol, LastRow - FirstRow, WidthCols)

    TopRow = LastRow + 2

End Function

Private Function ComparisonRowIsBlank( _
    ByVal ws As Worksheet, _
    ByVal RowNo As Long, _
    ByVal LeftCol As Long, _
    ByVal WidthCols As Long) As Boolean

    ComparisonRowIsBlank = _
        (Application.CountA( _
            ws.Range( _
                ws.Cells(RowNo, LeftCol), _
                ws.Cells(RowNo, LeftCol + WidthCols))) = 0)

End Function

' =====================================================================
' COMPARISON FEATURE - end
' =====================================================================
