Attribute VB_Name = "modSnapshotMain"
Option Explicit

Private Const SNAP_META_SHEET As String = "_snapshot_meta"
Private Const BLOCK_SIZE As Long = 20

Public Sub SnapshotCreateActiveWorkbook()
    On Error GoTo EH

    Dim wb As Workbook
    Dim snapshotPath As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6101, "SnapshotCreateActiveWorkbook", "No active workbook."
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 6103, "SnapshotCreateActiveWorkbook", "Workbook must be saved before creating snapshot."

    snapshotPath = CreateWorkbookSnapshotFile(wb)

    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo EH
    ShowSnapshotInfo "Snapshot created:" & vbCrLf & snapshotPath
    Exit Sub
EH:
    ShowSnapshotError "Snapshot creation failed", Err.Description
End Sub

Public Sub SnapshotCompareActiveWorkbook()
    On Error GoTo EH
    Dim wb As Workbook
    Dim fp As Variant
    Dim reportPath As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6102, "SnapshotCompareActiveWorkbook", "No active workbook."
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 6104, "SnapshotCompareActiveWorkbook", "Workbook must be saved before compare."

    fp = PickSnapshotFile("Select snapshot XLSX")
    If VarType(fp) = vbBoolean Then
        ShowSnapshotInfo "Compare canceled (no snapshot selected)."
        Exit Sub
    End If

    reportPath = CompareCurrentWorkbookToSnapshot(wb, CStr(fp))
    ShowCompareReportCreated reportPath
    Exit Sub
EH:
    ShowSnapshotError "Compare failed", Err.Description
End Sub

Public Sub SnapshotCompareTwoSnapshots()
    On Error GoTo EH
    Dim fpA As Variant, fpB As Variant, reportPath As String
    fpA = PickSnapshotFile("Select Snapshot A")
    If VarType(fpA) = vbBoolean Then
        ShowSnapshotInfo "Compare 2 canceled (Snapshot A not selected)."
        Exit Sub
    End If
    ShowSnapshotInfo "Select Snapshot B (second file for comparison)."
    fpB = PickSnapshotFile("Select Snapshot B")
    If VarType(fpB) = vbBoolean Then
        ShowSnapshotInfo "Compare 2 canceled (Snapshot B not selected)."
        Exit Sub
    End If

    reportPath = CompareTwoSnapshots(CStr(fpA), CStr(fpB))
    ShowCompareReportCreated reportPath
    Exit Sub
EH:
    ShowSnapshotError "Compare 2 snapshots failed", Err.Description
End Sub

Private Sub ShowCompareReportCreated(ByVal reportPath As String)
    Dim resp As VbMsgBoxResult

    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    AppActivate Application.Caption
    resp = MsgBox("Compare report created:" & vbCrLf & reportPath & vbCrLf & vbCrLf & "Open report now?", vbQuestion Or vbYesNo Or vbDefaultButton1 Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot")

    If resp = vbYes Then
        Workbooks.Open Filename:=reportPath, UpdateLinks:=0, ReadOnly:=True
    End If

    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
End Sub

Private Sub ShowSnapshotInfo(ByVal msg As String)
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    AppActivate Application.Caption
    MsgBox msg, vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
End Sub


Public Sub SnapshotCompareAnyTwoFiles()
    On Error GoTo EH
    Dim fpA As Variant, fpB As Variant, reportPath As String

    fpA = PickAnyExcelFile("Select File A")
    If VarType(fpA) = vbBoolean Then
        ShowSnapshotInfo "Compare Files canceled (File A not selected)."
        Exit Sub
    End If

    fpB = PickAnyExcelFile("Select File B")
    If VarType(fpB) = vbBoolean Then
        ShowSnapshotInfo "Compare Files canceled (File B not selected)."
        Exit Sub
    End If

    reportPath = CompareTwoSnapshots(CStr(fpA), CStr(fpB))
    ShowCompareReportCreated reportPath
    Exit Sub
EH:
    ShowSnapshotError "Compare Files failed", Err.Description
End Sub

Public Sub SnapshotOpenActiveWorkbookFolder()
    On Error GoTo EH
    Dim wb As Workbook
    Dim snapFolder As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6110, "SnapshotOpenActiveWorkbookFolder", "No active workbook."
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 6111, "SnapshotOpenActiveWorkbookFolder", "Workbook must be saved before opening snapshot folder."

    snapFolder = BuildSnapshotFolder(wb)
    If Not modSnapshotStorage.EnsureFolderPath(snapFolder) Then Err.Raise vbObjectError + 6112, "SnapshotOpenActiveWorkbookFolder", "Cannot create snapshot folder."

    OpenFolderInExplorer snapFolder
    Exit Sub
EH:
    ShowSnapshotError "Open snapshot folder failed", Err.Description
End Sub

Public Sub SnapshotCompareLatestTwo()
    On Error GoTo EH
    Dim wb As Workbook
    Dim folderPath As String
    Dim latestA As String
    Dim latestB As String
    Dim reportPath As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6113, "SnapshotCompareLatestTwo", "No active workbook."
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 6114, "SnapshotCompareLatestTwo", "Workbook must be saved before compare."

    folderPath = BuildSnapshotFolder(wb)
    GetTwoLatestSnapshotFiles folderPath, latestA, latestB

    If Len(latestA) = 0 Or Len(latestB) = 0 Then
        ShowSnapshotInfo "Need at least 2 snapshots in:" & vbCrLf & folderPath
        Exit Sub
    End If

    reportPath = CompareTwoSnapshots(latestA, latestB)
    ShowCompareReportCreated reportPath
    Exit Sub
EH:
    ShowSnapshotError "Compare latest snapshots failed", Err.Description
End Sub

Private Sub GetTwoLatestSnapshotFiles(ByVal folderPath As String, ByRef latestA As String, ByRef latestB As String)
    Dim fn As String
    Dim fullPath As String
    Dim dt As Date
    Dim best1 As Date, best2 As Date

    latestA = ""
    latestB = ""
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    fn = Dir$(folderPath & "\*_snapshot_*.xlsx")
    Do While Len(fn) > 0
        fullPath = folderPath & "\" & fn
        On Error Resume Next
        dt = FileDateTime(fullPath)
        On Error GoTo 0

        If dt >= best1 Then
            best2 = best1
            latestB = latestA
            best1 = dt
            latestA = fullPath
        ElseIf dt > best2 Then
            best2 = dt
            latestB = fullPath
        End If

        fn = Dir$
    Loop
End Sub

Private Sub OpenFolderInExplorer(ByVal folderPath As String)
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo 0
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Sub

Private Function PickAnyExcelFile(ByVal titleText As String) As Variant
    PickAnyExcelFile = PickExcelFileWithInitialFolder( _
        titleText, _
        "Excel Files (*.xls;*.xlsx;*.xlsm;*.xlsb),*.xls;*.xlsx;*.xlsm;*.xlsb", _
        modNavigatorSettings.ResolveOpenFilesInitialFolder(ThisWorkbook.Path) _
    )
End Function

Private Function PickSnapshotFile(ByVal titleText As String) As Variant
    PickSnapshotFile = PickExcelFileWithInitialFolder( _
        titleText, _
        "Excel Files (*.xlsx),*.xlsx", _
        modNavigatorSettings.ResolveOpenFilesInitialFolder(ThisWorkbook.Path) _
    )
End Function

Private Function PickExcelFileWithInitialFolder(ByVal titleText As String, ByVal fileFilter As String, ByVal initialFolder As String) As Variant
    Dim prevPath As String
    Dim prevDrive As String

    prevPath = CurDir$
    prevDrive = Left$(prevPath, 2)

    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    AppActivate Application.Caption
    If Len(initialFolder) > 0 Then
        If Mid$(initialFolder, 2, 1) = ":" Then ChDrive Left$(initialFolder, 1)
        ChDir initialFolder
    End If
    On Error GoTo 0

    PickExcelFileWithInitialFolder = Application.GetOpenFilename(fileFilter, , titleText)

    On Error Resume Next
    If Len(prevDrive) = 2 And Right$(prevDrive, 1) = ":" Then ChDrive Left$(prevDrive, 1)
    If Len(prevPath) > 0 Then ChDir prevPath
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Function

Private Sub ShowSnapshotError(ByVal prefix As String, ByVal errMsg As String)
    If Len(errMsg) = 0 Then errMsg = "Unknown error."
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    AppActivate Application.Caption
    MsgBox prefix & ": " & errMsg, vbCritical Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
End Sub
