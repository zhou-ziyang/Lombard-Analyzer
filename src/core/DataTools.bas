Attribute VB_Name = "DataTools"
Option Explicit

Public Function GetDateCode( _
    ByVal InputDate As Date) As String

    GetDateCode = Format(InputDate, "yyyymmdd")

End Function

Public Function Nz(ValueIn As Variant, Optional DefaultValue As Variant = 0) As Variant

    If IsError(ValueIn) Then
        Nz = DefaultValue
    ElseIf IsEmpty(ValueIn) Then
        Nz = DefaultValue
    ElseIf Trim(CStr(ValueIn)) = "" Then
        Nz = DefaultValue
    Else
        Nz = ValueIn
    End If

End Function

