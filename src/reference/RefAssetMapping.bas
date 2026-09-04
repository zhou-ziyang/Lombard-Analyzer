Attribute VB_Name = "RefAssetMapping"
Option Explicit

'====================================================
' Asset Classification
'====================================================

'
' Maps the raw Sophis Asset Type / Classification string onto the reporting
' asset classes. Matching is done on an upper-cased copy, and the patterns are
' upper case to match: Like is case-sensitive under this module's default
' Option Compare Binary, so a re-cased extract would otherwise send a whole
' asset class to UNKNOWN.
'
' UNKNOWN means the string did not match any pattern, not that the position is
' ineligible. Only an explicit Non Eligible marking makes a position
' ineligible, so callers count UNKNOWN collateral as eligible and surface the
' unmatched string for the mapping rules to be extended.
'
Public Function GetAssetClass( _
    ByVal AssetType As String) As String

    AssetType = UCase$(Trim$(AssetType))

    If AssetType Like "NON ELIGIBLE*" Then
        GetAssetClass = "Non Eligible Asset"

    ElseIf AssetType Like "CERTIFICATES*" Then
        GetAssetClass = "Certificates"

    ElseIf AssetType Like "CURRENCY*" Then
        GetAssetClass = "Cash"

    ElseIf AssetType Like "SEGREGATED ACCOUNT*" Then
        GetAssetClass = "GP"

    ElseIf AssetType Like "FUNDS*" Then
        GetAssetClass = "Funds"

    ElseIf AssetType Like "INSURANCE*" Then
        GetAssetClass = "Insurance"

    ElseIf AssetType Like "SENIOR CORPORATE*" _
        Or AssetType Like "SOVEREIGN*" _
        Or AssetType Like "SUBORDINATED CORPORATE*" Then
        GetAssetClass = "Bonds"

    ElseIf AssetType Like "STOCKS*" Then
        GetAssetClass = "Equity"

    Else
        GetAssetClass = "UNKNOWN"

    End If

End Function


'====================================================
' Unknown Asset Registration
'====================================================

'
' Records an asset type GetAssetClass could not place, once per run, and
' routes it through Note so the weekly report picks it up. Do not add a
' NoteHandler parameter here: Note reads the global of that name, and a
' parameter would shadow it inside this procedure.
'
Public Sub RegisterUnknownAsset( _
    ByVal AssetType As String, _
    ByRef UnknownAssets As Object)

    If AssetType = "" Then Exit Sub

    If Not UnknownAssets.Exists(AssetType) Then

        UnknownAssets.Add AssetType, True

        Note _
            "Unmapped Asset Type" & vbLf & AssetType

    End If

End Sub
