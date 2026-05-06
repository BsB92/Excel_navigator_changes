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
        frmExcelNavigator.RestoreNavigatorToFront
        On Error GoTo EH
        MsgBox "Workbook must be saved before creating snapshot.", vbExclamation, "ExcelNavigator Snapshot"
        Exit Sub
    End If

    Set snap = modSnapshotCapture.CaptureWorkbookSnapshot(wb)
    jsonText = modSnapshotJson.SerializeWorkbookSnapshot(snap)
    finalPath = modSnapshotStorage.SaveSnapshotTransactional(wb, jsonText, snap("CreatedBy"))

    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo EH
    MsgBox "Snapshot created:" & vbCrLf & finalPath, vbInformation, "ExcelNavigator Snapshot"
    Exit Sub

EH:
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo 0
    MsgBox "Snapshot creation failed: " & Err.Description, vbCritical, "ExcelNavigator Snapshot"
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
    On Error GoTo EH
    fp = Application.GetOpenFilename("JSON Files (*.json),*.json", , "Select snapshot JSON")
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
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo 0
    MsgBox "Compare failed: " & Err.Description, vbCritical, "ExcelNavigator Snapshot"
End Sub
