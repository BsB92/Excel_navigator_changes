Attribute VB_Name = "modSnapshotMain"
Option Explicit

Public Sub SnapshotCreateActiveWorkbook()
    On Error GoTo EH

    Dim wb As Workbook
    Dim snap As Object
    Dim jsonText As String
    Dim finalPath As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6101, "SnapshotCreateActiveWorkbook", "No active workbook."

    If Len(wb.Path) = 0 Then
        On Error Resume Next
        modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
        frmExcelNavigator.RestoreNavigatorToFront
        On Error GoTo EH
        MsgBox "Workbook must be saved before creating snapshot.", vbExclamation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
        On Error Resume Next
        modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
        On Error GoTo EH
        Exit Sub
    End If

    Set snap = modSnapshotCapture.CaptureWorkbookSnapshot(wb)
    jsonText = modSnapshotJson.SerializeWorkbookSnapshot(snap)
    finalPath = modSnapshotStorage.SaveSnapshotTransactional(wb, jsonText, snap("CreatedBy"))

    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo EH
    MsgBox "Snapshot created:" & vbCrLf & finalPath, vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo EH
    Exit Sub

EH:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errMsg) = 0 Then errMsg = "Unknown error."
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo 0
    MsgBox "Snapshot creation failed: " & errMsg, vbCritical Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Sub

Public Sub SnapshotCompareActiveWorkbook()
    On Error GoTo EH

    Dim wb As Workbook
    Dim fp As Variant
    Dim oldSnap As Object
    Dim newSnap As Object
    Dim options As Object
    Dim diff As Object

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6102, "SnapshotCompareActiveWorkbook", "No active workbook."

    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo EH
    fp = Application.GetOpenFilename("JSON Files (*.json),*.json", , "Select snapshot JSON")
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo EH
    If VarType(fp) = vbBoolean Then Exit Sub

    Set oldSnap = modSnapshotJson.LoadWorkbookSnapshot(CStr(fp))
    Set newSnap = modSnapshotCapture.CaptureWorkbookSnapshot(wb)

    Set options = CreateObject("Scripting.Dictionary")
    options("IgnoreBlankChanges") = True
    options("IgnoreFormulaResultChanges") = False
    options("IgnoreVolatileFormulas") = False
    Set diff = modSnapshotCompare.CompareSnapshots(oldSnap, newSnap, options)

    modSnapshotReport.GenerateDiffReport diff, wb, CStr(fp)
    Exit Sub

EH:
    Dim compareErr As String
    compareErr = Err.Description
    If Len(compareErr) = 0 Then compareErr = "Unknown error."
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo 0
    MsgBox "Compare failed: " & compareErr, vbCritical Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Sub
