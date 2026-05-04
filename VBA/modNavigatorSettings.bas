Attribute VB_Name = "modNavigatorSettings"
Option Explicit

Private Const REG_APP As String = "ExcelNavigator46_alpha"
Private Const REG_SEC_SETTINGS As String = "Settings"
Private Const KEY_DEFAULT_FOLDER As String = "DefaultFolder"
Private Const HELP_FILE_NAME As String = "ExcelNavigator_Help.txt"

Public Function GetDefaultWorkingFolder() As String
    GetDefaultWorkingFolder = GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, ThisWorkbook.Path)
End Function

Public Sub SaveDefaultWorkingFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, folderPath
End Sub

Public Function ResolveInitialFolder(ByVal fallbackFolder As String) As String
    Dim configuredFolder As String
    configuredFolder = Trim$(GetDefaultWorkingFolder())

    If Len(configuredFolder) > 0 Then
        If Len(Dir$(configuredFolder, vbDirectory)) > 0 Then
            ResolveInitialFolder = configuredFolder
            Exit Function
        End If
    End If

    ResolveInitialFolder = fallbackFolder
End Function

Public Sub OpenHelpInstructions()
    Dim helpPath As String
    helpPath = ThisWorkbook.Path & Application.PathSeparator & HELP_FILE_NAME
    If Len(Dir$(helpPath, vbNormal)) = 0 Then
        Err.Raise vbObjectError + 4701, "OpenHelpInstructions", "Brak pliku pomocy: " & helpPath
    End If
    ThisWorkbook.FollowHyperlink helpPath
End Sub
