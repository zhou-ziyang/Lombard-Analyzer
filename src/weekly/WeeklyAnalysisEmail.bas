Attribute VB_Name = "WeeklyAnalysisEmail"
Option Explicit

' v49: refers to aggregated accounts in the email body.

Private HtmlFragmentCounter As Long
Private HtmlStyleBlocks As String

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
    ' Report blocks, in reading order. Each one is a Layout anchor plus
    ' the height and width of the block that starts there.
    '

    HTMLBody = HTMLBody & _
        BlockHtml(ws, Layout.PortfolioRow, Layout.PortfolioCol, 7, 3) & _
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



