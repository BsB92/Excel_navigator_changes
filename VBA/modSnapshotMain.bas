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
    MsgBox "Snapshot created:" & vbCrLf & snapshotPath, vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo EH
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
    If VarType(fp) = vbBoolean Then Exit Sub

    reportPath = CompareCurrentWorkbookToSnapshot(wb, CStr(fp))
    MsgBox "Compare report created:" & vbCrLf & reportPath, vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    Exit Sub
EH:
    ShowSnapshotError "Compare failed", Err.Description
End Sub

Public Sub SnapshotCompareTwoSnapshots()
    On Error GoTo EH
    Dim fpA As Variant, fpB As Variant, reportPath As String
    fpA = PickSnapshotFile("Select Snapshot A")
    If VarType(fpA) = vbBoolean Then Exit Sub
    fpB = PickSnapshotFile("Select Snapshot B")
    If VarType(fpB) = vbBoolean Then Exit Sub

    reportPath = CompareTwoSnapshots(CStr(fpA), CStr(fpB))
    MsgBox "Compare report created:" & vbCrLf & reportPath, vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    Exit Sub
EH:
    ShowSnapshotError "Compare 2 snapshots failed", Err.Description
End Sub

Private Function PickSnapshotFile(ByVal titleText As String) As Variant
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo 0
    PickSnapshotFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , titleText)
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Function

Private Sub ShowSnapshotError(ByVal prefix As String, ByVal errMsg As String)
    If Len(errMsg) = 0 Then errMsg = "Unknown error."
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    MsgBox prefix & ": " & errMsg, vbCritical Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
End Sub
