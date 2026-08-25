Attribute VB_Name = "PortfolioDataTools"
Option Explicit

Public Type PortfolioStats

    LoanCount As Long
    ApprovedAmount As Double
    DrawnAmount As Double
    CollateralValue As Double

End Type

'Public Function GetPortfolioStats( _
'    ByVal RefDate As Date) _
'    As PortfolioStats
'
'    Dim Stats As PortfolioStats
'
'    Dim wsAccounts As Worksheet
'    Dim wsPositions As Worksheet
'
'    Dim LastRow As Long
'    Dim r As Long
'
'    Dim DictNDG As Object
'    Dim NDG As String
'
'    Set DictNDG = CreateObject("Scripting.Dictionary")
'
'    Set wsAccounts = _
'        ImportAccountsByDate( _
'            GetDateCode(RefDate))
'
'    Set wsPositions = _
'        ImportPositionsByDate( _
'            GetDateCode(RefDate))
'
'    '
'    ' Accounts
'    '
'
'    If Not wsAccounts Is Nothing Then
'
'        LastRow = GetLastRow(wsAccounts, "B")
'
'        For r = 3 To LastRow
'
'            NDG = Trim(wsAccounts.Cells(r, 2).Value)
'
'            If NDG <> "" Then
'
'                DictNDG(NDG) = True
'
'            End If
'
'            Stats.ApprovedAmount = _
'                Stats.ApprovedAmount + _
'                CDbl(Nz(wsAccounts.Cells(r, 5)))
'
'            Stats.DrawnAmount = _
'                Stats.DrawnAmount + _
'                CDbl(Nz(wsAccounts.Cells(r, 6)))
'
'        Next r
'
'        Stats.LoanCount = DictNDG.Count
'
'    End If
'
'    '
'    ' Positions
'    '
'
'    If Not wsPositions Is Nothing Then
'
'        LastRow = GetLastRow(wsPositions, "B")
'
'        For r = 3 To LastRow
'
'            Stats.CollateralValue = _
'                Stats.CollateralValue + _
'                CDbl(Nz(wsPositions.Cells(r, 14)))
'
'        Next r
'
'    End If
'
'    GetPortfolioStats = Stats
'
'End Function

Public Function GetPortfolioStats( _
    ByVal RefDate As Date) _
    As PortfolioStats

    Dim Stats As PortfolioStats

    Dim wsAccounts As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim DictNDG As Object
    Dim NDG As String

    Set DictNDG = CreateObject("Scripting.Dictionary")

    Set wsAccounts = _
        ImportAccountsByDate( _
            GetDateCode(RefDate))

    If Not wsAccounts Is Nothing Then

        LastRow = GetLastRow(wsAccounts, "B")

        For r = 3 To LastRow

            NDG = Trim(wsAccounts.Cells(r, 2).Value)

            If NDG <> "" Then

                DictNDG(NDG) = True

            End If

            Stats.ApprovedAmount = _
                Stats.ApprovedAmount + _
                CDbl(Nz(wsAccounts.Cells(r, 5)))

            Stats.DrawnAmount = _
                Stats.DrawnAmount + _
                CDbl(Nz(wsAccounts.Cells(r, 6)))

        Next r

        Stats.LoanCount = DictNDG.Count

    End If

    GetPortfolioStats = Stats

End Function

Public Function GetNDGDictionary( _
    ByVal wsAccounts As Worksheet) _
    As Object

    Dim Dict As Object

    Dim LastRow As Long
    Dim r As Long
    Dim NDG As String

    Set Dict = CreateObject("Scripting.Dictionary")

    LastRow = GetLastRow(wsAccounts, "B")

    For r = 3 To LastRow

        NDG = Trim(wsAccounts.Cells(r, 2).Value)

        If NDG <> "" Then

            Dict(NDG) = True

        End If

    Next r

    Set GetNDGDictionary = Dict

End Function

