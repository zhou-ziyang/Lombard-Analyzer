Attribute VB_Name = "AssetMapping"
Option Explicit

'====================================================
' Asset Classification
'====================================================

Public Function GetAssetClass( _
    ByVal AssetType As String) As String

    AssetType = Trim(AssetType)

    If AssetType Like "Non Eligible*" Then

        GetAssetClass = "Non Eligible Asset"

    ElseIf AssetType Like "Certificates*" Then

        GetAssetClass = "Certificates"

    ElseIf AssetType Like "Currency*" Then

        GetAssetClass = "Cash"

    ElseIf AssetType Like "Segregated Account*" Then

        GetAssetClass = "GP"

    ElseIf AssetType Like "Funds*" Then

        GetAssetClass = "Funds"

    ElseIf AssetType Like "Insurance*" Then

        GetAssetClass = "Insurance"

    ElseIf AssetType Like "Senior Corporate*" _
        Or AssetType Like "Sovereign*" _
        Or AssetType Like "Subordinated Corporate*" Then

        GetAssetClass = "Bonds"

    ElseIf AssetType Like "Stocks*" Then

        GetAssetClass = "Equity"

    Else

        GetAssetClass = "UNKNOWN"

    End If

End Function


'====================================================
' Unknown Asset Registration
'====================================================

Public Sub RegisterUnknownAsset( _
    ByVal AssetType As String, _
    ByRef UnknownAssets As Object, _
    Optional ByVal NoteHandler As String = "")

    If AssetType = "" Then Exit Sub

    If Not UnknownAssets.Exists(AssetType) Then

        UnknownAssets.Add AssetType, True

        Note _
            "Unmapped Asset Type" & vbLf & AssetType

    End If

End Sub
