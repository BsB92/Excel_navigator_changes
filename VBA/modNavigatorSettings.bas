Attribute VB_Name = "modNavigatorSettings"
Option Explicit

Private Const REG_APP As String = "ExcelNavigator_v5.2"
Private Const REG_SEC_SETTINGS As String = "Settings"
Private Const KEY_DEFAULT_FOLDER As String = "DefaultFolder"
Private Const KEY_OPEN_FILES_FOLDER As String = "OpenFilesFolder"
Private Const HELP_FILE_NAME As String = "ExcelNavigator_Help.txt"
Private Const KEY_SNAPSHOT_COMPARE_MODE As String = "SnapshotCompareMode"
Private Const KEY_OPEN_COPIED_FILES As String = "OpenCopiedFiles"
Private Const KEY_OPEN_TARGET_FOLDER As String = "OpenTargetFolder"

Public Const SNAP_COMPARE_MODE_STRICT As String = "STRICT"
Public Const SNAP_COMPARE_MODE_VALUE_ONLY As String = "VALUE_ONLY"
Public Const SNAP_COMPARE_MODE_HYBRID As String = "HYBRID"

Public Function GetDefaultWorkingFolder() As String
    GetDefaultWorkingFolder = NormalizeFolderPath(GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, ThisWorkbook.Path), ThisWorkbook.Path)
End Function

Public Sub SaveDefaultWorkingFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_DEFAULT_FOLDER, NormalizeFolderPath(folderPath, ThisWorkbook.Path)
End Sub


Public Function GetOpenFilesFolder() As String
    GetOpenFilesFolder = NormalizeFolderPath(GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_OPEN_FILES_FOLDER, ThisWorkbook.Path), ThisWorkbook.Path)
End Function

Public Sub SaveOpenFilesFolder(ByVal folderPath As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_OPEN_FILES_FOLDER, NormalizeFolderPath(folderPath, ThisWorkbook.Path)
End Sub

Public Function ResolveOpenFilesInitialFolder(ByVal fallbackFolder As String) As String
    ResolveOpenFilesInitialFolder = ResolveExistingFolder(Trim$(GetOpenFilesFolder()), fallbackFolder)
End Function

Public Function ResolveInitialFolder(ByVal fallbackFolder As String) As String
    ResolveInitialFolder = ResolveExistingFolder(Trim$(GetDefaultWorkingFolder()), fallbackFolder)
End Function

Public Function NormalizeUserFolderPath(ByVal folderPath As String, ByVal fallbackBase As String) As String
    NormalizeUserFolderPath = NormalizeFolderPath(folderPath, fallbackBase)
End Function

Private Function ResolveExistingFolder(ByVal configuredFolder As String, ByVal fallbackFolder As String) As String
    If Len(configuredFolder) > 0 Then
        If IsWebPath(configuredFolder) Then
            ResolveExistingFolder = configuredFolder
            Exit Function
        End If

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

Public Function GetSnapshotCompareMode() As String
    GetSnapshotCompareMode = UCase$(GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_SNAPSHOT_COMPARE_MODE, SNAP_COMPARE_MODE_VALUE_ONLY))
End Function

Public Sub SaveSnapshotCompareMode(ByVal modeValue As String)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_SNAPSHOT_COMPARE_MODE, UCase$(Trim$(modeValue))
End Sub

Public Function GetOpenCopiedFilesEnabled() As Boolean
    GetOpenCopiedFilesEnabled = CBool(Val(GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_OPEN_COPIED_FILES, "0")))
End Function

Public Sub SaveOpenCopiedFilesEnabled(ByVal enabled As Boolean)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_OPEN_COPIED_FILES, IIf(enabled, "1", "0")
End Sub

Public Function GetOpenTargetFolderEnabled() As Boolean
    GetOpenTargetFolderEnabled = CBool(Val(GetSetting(REG_APP, REG_SEC_SETTINGS, KEY_OPEN_TARGET_FOLDER, "0")))
End Function

Public Sub SaveOpenTargetFolderEnabled(ByVal enabled As Boolean)
    SaveSetting REG_APP, REG_SEC_SETTINGS, KEY_OPEN_TARGET_FOLDER, IIf(enabled, "1", "0")
End Sub

Private Function IsWebPath(ByVal folderPath As String) As Boolean
    Dim v As String
    v = LCase$(Trim$(folderPath))
    IsWebPath = (Left$(v, 7) = "http://" Or Left$(v, 8) = "https://")
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String, ByVal fallbackBase As String) As String
    Dim p As String
    Dim fso As Object

    p = Trim$(folderPath)
    If Len(p) = 0 Then
        NormalizeFolderPath = p
        Exit Function
    End If

    If IsWebPath(p) Then
        NormalizeFolderPath = p
        Exit Function
    End If

    If InStr(1, p, ":", vbTextCompare) = 0 And Left$(p, 2) <> "\" Then
        If Len(Trim$(fallbackBase)) > 0 Then p = Trim$(fallbackBase) & Application.PathSeparator & p
    End If

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then p = fso.GetAbsolutePathName(p)
    On Error GoTo 0

    NormalizeFolderPath = p
End Function
