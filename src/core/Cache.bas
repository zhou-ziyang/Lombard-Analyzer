Attribute VB_Name = "Cache"
Option Explicit

Public Type PositionCache

    Data() As Variant
    Keys() As String

    LineCount As Long

End Type

Public Function LoadPositionCache( _
    ByVal Lines As Variant, _
    ByVal idxNDG As Long, _
    ByVal idxCO As Long, _
    ByVal idxISIN As Long, _
    ByVal idxSec As Long, _
    ByVal idxAsset As Long) As PositionCache

    Dim Cache As PositionCache

    Dim i As Long
    Dim Arr As Variant
    Dim Capacity As Long

    Cache.LineCount = UBound(Lines)
    If Cache.LineCount < 0 Then Cache.LineCount = 0

    '
    ' A file holding only a header row has no data lines.  The arrays are
    ' still allocated so callers can index them safely, while LineCount
    ' keeps their loops empty.
    '
    Capacity = Cache.LineCount
    If Capacity < 1 Then Capacity = 1

    ReDim Cache.Data(1 To Capacity)
    ReDim Cache.Keys(1 To Capacity)

    For i = 1 To Cache.LineCount

        If Trim(Lines(i)) <> "" Then

            Arr = Split(Lines(i), ";")

            Cache.Data(i) = Arr

            Cache.Keys(i) = _
                BuildPositionKey( _
                    Arr, _
                    idxNDG, _
                    idxCO, _
                    idxISIN, _
                    idxSec, _
                    idxAsset)

        End If

    Next i

    LoadPositionCache = Cache

End Function

Public Function BuildPositionKey( _
    ByVal PositionFields As Variant, _
    ByVal idxNDG As Long, _
    ByVal idxCO As Long, _
    ByVal idxISIN As Long, _
    ByVal idxSec As Long, _
    ByVal idxAsset As Long) As String

    Dim NDG As String
    Dim InstrumentId As String

    NDG = SafeField(PositionFields, idxNDG)

    If NDG = "" Then Exit Function

    If SafeField(PositionFields, idxISIN) <> "" Then
        InstrumentId = SafeField(PositionFields, idxISIN)
    Else
        InstrumentId = SafeField(PositionFields, idxSec)
    End If

    BuildPositionKey = _
        NDG & "|" & _
        SafeField(PositionFields, idxCO) & "|" & _
        InstrumentId & "|" & _
        SafeField(PositionFields, idxAsset)

End Function

Public Function SafeField(Arr As Variant, idx As Long) As String
    If idx >= 0 And idx <= UBound(Arr) Then
        SafeField = Trim(Arr(idx))
    Else
        SafeField = ""
    End If
End Function
