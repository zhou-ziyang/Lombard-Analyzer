Attribute VB_Name = "ToolsInstall"
Option Explicit

'
' Replaces every standard module in this workbook with the .bas files in a
' folder.  Doing it by hand means removing each old module before importing
' its replacement, because Import does not overwrite: the VBE keeps the old
' module and names the new one DeltaCalculation1, and two modules holding the
' same procedures fail to compile with "Ambiguous name detected".
'
' Needs File > Options > Trust Center > Trust Center Settings > Macro
' Settings > Trust access to the VBA project object model.  Without it
' ThisWorkbook.VBProject raises 1004 and this reports that instead.
'
' The folder is expected to hold every module the workbook should end up
' with.  Anything present here and absent there is removed and not replaced,
' and the confirmation lists those separately before anything is touched.
'
' This module removes every standard module except itself - it cannot delete
' the code it is running.  Remove it by hand when the workbook is finished.
'

Private Const STANDARD_MODULE As Long = 1

Public Sub InstallModules()

    Dim Project As Object
    Dim Component As Object

    Dim FolderPath As String
    Dim BackupPath As String

    Dim FileNames As Collection
    Dim ModuleNames As Collection
    Dim Removing As Collection
    Dim Orphans As Collection

    Dim FileName As Variant
    Dim ModuleName As Variant

    Dim Message As String
    Dim Imported As Long

    On Error GoTo Failed

    Set Project = VbaProject()
    If Project Is Nothing Then Exit Sub

    FolderPath = ChosenFolder()
    If FolderPath = "" Then Exit Sub

    Set FileNames = BasFilesIn(FolderPath)

    If FileNames.Count = 0 Then
        MsgBox "No .bas files in:" & vbCrLf & FolderPath, vbExclamation
        Exit Sub
    End If

    Set ModuleNames = New Collection
    For Each FileName In FileNames
        ModuleNames.Add BaseName(CStr(FileName))
    Next FileName

    '
    ' This module is never removed, so it is never an orphan either.
    '
    ModuleNames.Add ThisModuleName()

    '
    ' Work out what goes before anything goes, so the confirmation can be
    ' honest about the modules that have no replacement in the folder.
    '
    Set Removing = New Collection
    Set Orphans = New Collection

    For Each Component In Project.VBComponents
        If Component.Type = STANDARD_MODULE _
           And Component.Name <> ThisModuleName() Then

            Removing.Add Component.Name

            If Not Contains(ModuleNames, Component.Name) Then
                Orphans.Add Component.Name
            End If

        End If
    Next Component

    Message = _
        "Remove " & Removing.Count & " module(s) and import " & _
        FileNames.Count & " file(s) from:" & vbCrLf & FolderPath

    If Orphans.Count > 0 Then
        Message = Message & vbCrLf & vbCrLf & _
            "Removed with NO replacement in that folder:" & vbCrLf & _
            "    " & Joined(Orphans, ", ")
    End If

    Message = Message & vbCrLf & vbCrLf & _
        "Every removed module is exported to a backup folder first."

    If MsgBox(Message, vbOKCancel + vbQuestion, "Install modules") <> vbOK Then
        Exit Sub
    End If

    BackupPath = MakeBackupFolder(FolderPath)

    For Each ModuleName In Removing
        Project.VBComponents(CStr(ModuleName)).Export _
            BackupPath & CStr(ModuleName) & ".bas"
    Next ModuleName

    For Each ModuleName In Removing
        Project.VBComponents.Remove Project.VBComponents(CStr(ModuleName))
    Next ModuleName

    For Each FileName In FileNames

        '
        ' Importing this module again would land it beside itself as
        ' ToolsInstall1, so the folder is allowed to contain it.
        '
        If StrComp(BaseName(CStr(FileName)), _
                   ThisModuleName(), vbTextCompare) <> 0 Then

            Project.VBComponents.Import CStr(FileName)
            Imported = Imported + 1

        End If

    Next FileName

    MsgBox _
        "Removed " & Removing.Count & ", imported " & Imported & "." & _
        vbCrLf & vbCrLf & _
        "Backup of the removed modules:" & vbCrLf & BackupPath & _
        vbCrLf & vbCrLf & _
        "Now run Debug > Compile VBAProject.", _
        vbInformation, "Install modules"

    Exit Sub

Failed:

    MsgBox _
        "Install failed at " & Err.Source & ":" & vbCrLf & _
        Err.Number & "  " & Err.Description & vbCrLf & vbCrLf & _
        "Modules already removed are in the backup folder.", _
        vbCritical, "Install modules"

End Sub

'
' ThisWorkbook.VBProject is the one thing here that needs the trust setting,
' so it is asked for once and the failure explains itself.
'
Private Function VbaProject() As Object

    On Error Resume Next

    Set VbaProject = ThisWorkbook.VBProject

    If Err.Number <> 0 Then

        Err.Clear

        MsgBox _
            "No programmatic access to the VBA project." & vbCrLf & vbCrLf & _
            "File > Options > Trust Center > Trust Center Settings >" & _
            vbCrLf & _
            "Macro Settings > Trust access to the VBA project object model", _
            vbExclamation, "Install modules"

        Set VbaProject = Nothing

    End If

End Function

Private Function ChosenFolder() As String

    Dim Dialog As FileDialog

    Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)

    Dialog.Title = "Folder holding the .bas files"

    If Dialog.Show <> -1 Then Exit Function

    ChosenFolder = Dialog.SelectedItems(1)

    If Right$(ChosenFolder, 1) <> Application.PathSeparator Then
        ChosenFolder = ChosenFolder & Application.PathSeparator
    End If

End Function

'
' Dir keeps one search at a time, so the whole listing is collected before
' any other file call runs.
'
Private Function BasFilesIn( _
    ByVal FolderPath As String) As Collection

    Dim FileName As String

    Set BasFilesIn = New Collection

    FileName = Dir$(FolderPath & "*.bas")

    Do While FileName <> ""
        BasFilesIn.Add FolderPath & FileName
        FileName = Dir$()
    Loop

End Function

Private Function MakeBackupFolder( _
    ByVal FolderPath As String) As String

    MakeBackupFolder = _
        FolderPath & "backup_" & _
        Format$(Now, "yyyymmdd_hhnnss") & Application.PathSeparator

    MkDir MakeBackupFolder

End Function

Private Function BaseName( _
    ByVal FilePath As String) As String

    BaseName = Mid$(FilePath, InStrRev(FilePath, Application.PathSeparator) + 1)
    BaseName = Left$(BaseName, Len(BaseName) - Len(".bas"))

End Function

Private Function Contains( _
    ByVal Names As Collection, _
    ByVal Wanted As String) As Boolean

    Dim Candidate As Variant

    For Each Candidate In Names
        If StrComp(CStr(Candidate), Wanted, vbTextCompare) = 0 Then
            Contains = True
            Exit Function
        End If
    Next Candidate

End Function

Private Function Joined( _
    ByVal Names As Collection, _
    ByVal Separator As String) As String

    Dim Candidate As Variant

    For Each Candidate In Names
        If Joined <> "" Then Joined = Joined & Separator
        Joined = Joined & CStr(Candidate)
    Next Candidate

End Function

'
' A module cannot delete the code it is running, and VBA gives a procedure no
' way to ask its own module's name, so it is written down here.  Rename the
' module in the VBE and this has to change with it.
'
Private Function ThisModuleName() As String

    ThisModuleName = "ToolsInstall"

End Function
