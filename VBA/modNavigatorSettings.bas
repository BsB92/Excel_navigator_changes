Attribute VB_Name = "modNavigatorSettings"
Option Explicit

Private Const REG_APP As String = "ExcelNavigator_v4.7"
Private Const REG_SEC_SETTINGS As String = "Settings"
Private Const KEY_DEFAULT_FOLDER As String = "DefaultFolder"
Private Const KEY_OPEN_FILES_FOLDER As String = "OpenFilesFolder"
Private Const HELP_FILE_NAME As String = "ExcelNavigator_Help.txt"

Public Function GetDefaultWorkingFolder() As String
    GetDefaultWorkingFolder = GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, ThisWorkbook.Path)
End Function

Public Sub SaveDefaultWorkingFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, folderPath
End Sub


Public Function GetOpenFilesFolder() As String
    GetOpenFilesFolder = GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_OPEN_FILES_FOLDER, ThisWorkbook.Path)
End Function

Public Sub SaveOpenFilesFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_OPEN_FILES_FOLDER, folderPath
End Sub

Public Function ResolveOpenFilesInitialFolder(ByVal fallbackFolder As String) As String
    ResolveOpenFilesInitialFolder = ResolveExistingFolder(Trim$(GetOpenFilesFolder()), fallbackFolder)
End Function

Public Function ResolveInitialFolder(ByVal fallbackFolder As String) As String
    ResolveInitialFolder = ResolveExistingFolder(Trim$(GetDefaultWorkingFolder()), fallbackFolder)
End Function

Private Function ResolveExistingFolder(ByVal configuredFolder As String, ByVal fallbackFolder As String) As String
    If Len(configuredFolder) > 0 Then
        If Len(Dir$(configuredFolder, vbDirectory)) > 0 Then
            ResolveExistingFolder = configuredFolder
            Exit Function
        End If
    End If

    ResolveExistingFolder = fallbackFolder
End Function

Public Sub OpenHelpInstructions()
    Dim helpPath As String
    helpPath = ThisWorkbook.Path & Application.PathSeparator & HELP_FILE_NAME
    If Len(Dir$(helpPath, vbNormal)) = 0 Then
        Err.Raise vbObjectError + 4701, "OpenHelpInstructions", "Brak pliku pomocy: " & helpPath
    End If
    ThisWorkbook.FollowHyperlink helpPath
End Sub
