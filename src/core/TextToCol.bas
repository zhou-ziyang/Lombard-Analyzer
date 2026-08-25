Attribute VB_Name = "TextToCol"
Sub text_to_col1()

    ThisWorkbook.Worksheets("Home").Range("AnalysisEndDate").TextToColumns _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=True, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

End Sub

