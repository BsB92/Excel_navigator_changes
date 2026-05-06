Attribute VB_Name = "modSnapshotMain"
Option Explicit

Public Sub SnapshotCreateActiveWorkbook()
    On Error GoTo EH
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6101, "SnapshotCreateActiveWorkbook", "No active workbook."
    If Len(wb.Path) = 0 Then
        MsgBox "Workbook must be saved before creating snapshot.", vbExclamation, "ExcelNavigator Snapshot"
        Exit Sub
    End If

    Dim snap As clsWorkbookSnapshot
    Set snap = modSnapshotCapture.CaptureWorkbookSnapshot(wb)

    Dim jsonText As String
    jsonText = modSnapshotJson.SerializeWorkbookSnapshot(snap)

    Dim finalPath As String
    finalPath = modSnapshotStorage.SaveSnapshotTransactional(wb, jsonText, snap.CreatedBy)

    MsgBox "Snapshot created:" & vbCrLf & finalPath, vbInformation, "ExcelNavigator Snapshot"
    Exit Sub
EH:
    MsgBox "Snapshot creation failed: " & Err.Description, vbCritical, "ExcelNavigator Snapshot"
End Sub

Public Sub SnapshotCompareActiveWorkbook()
    On Error GoTo EH
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 6102, "SnapshotCompareActiveWorkbook", "No active workbook."

    Dim fp As Variant
    fp = Application.GetOpenFilename("JSON Files (*.json),*.json", , "Select snapshot JSON")
    If VarType(fp) = vbBoolean Then Exit Sub

    Dim oldSnap As clsWorkbookSnapshot
    Set oldSnap = modSnapshotJson.LoadWorkbookSnapshot(CStr(fp))

    Dim newSnap As clsWorkbookSnapshot
    Set newSnap = modSnapshotCapture.CaptureWorkbookSnapshot(wb)

    Dim options As clsSnapshotOptions
    Set options = New clsSnapshotOptions

    Dim diff As clsSnapshotDiffResult
    Set diff = modSnapshotCompare.CompareSnapshots(oldSnap, newSnap, options)

    modSnapshotReport.GenerateDiffReport diff, wb, CStr(fp)
    Exit Sub
EH:
    MsgBox "Compare failed: " & Err.Description, vbCritical, "ExcelNavigator Snapshot"
End Sub
