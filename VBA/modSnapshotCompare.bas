Attribute VB_Name = "modSnapshotCompare"
Option Explicit

Private Const META_SHEET As String = "_snapshot_meta"
Private Const BLOCK_SIZE As Long = 20

Public Function CreateWorkbookSnapshotFile(ByVal sourceWb As Workbook) As String
    Dim workPath As String, targetPath As String, snapRoot As String
    Dim tempCopy As String, snapWb As Workbook
    Dim prevAlerts As Boolean

    snapRoot = BuildSnapshotFolder(sourceWb)
    If Not modSnapshotStorage.EnsureFolderPath(snapRoot) Then Err.Raise vbObjectError + 6401, , "Cannot create snapshot folder."

    targetPath = snapRoot & "\" & SnapshotBaseName(sourceWb) & "_snapshot_" & Format$(Now, "yymmddhhnnss") & ".xlsx"
    tempCopy = Environ$("TEMP") & "\excelnav_snapshot_" & Format$(Now, "yymmddhhnnss") & ".tmp"
    sourceWb.SaveCopyAs tempCopy

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set snapWb = Workbooks.Open(Filename:=tempCopy, UpdateLinks:=0, ReadOnly:=False)
    BreakExternalLinks snapWb
    SaveSnapshotMeta snapWb, sourceWb.FullName
    snapWb.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook
    snapWb.Close SaveChanges:=False
    Kill tempCopy
    Application.DisplayAlerts = prevAlerts
    CreateWorkbookSnapshotFile = targetPath
End Function

Public Function CompareCurrentWorkbookToSnapshot(ByVal currentWb As Workbook, ByVal snapshotPath As String) As String
    Dim snapWb As Workbook
    Set snapWb = Workbooks.Open(Filename:=snapshotPath, UpdateLinks:=0, ReadOnly:=True)
    CompareCurrentWorkbookToSnapshot = BuildCompareReport(currentWb, snapWb, "Current", ExtractSnapshotStamp(snapshotPath), currentWb)
    snapWb.Close SaveChanges:=False
End Function

Public Function CompareTwoSnapshots(ByVal pathA As String, ByVal pathB As String) As String
    Dim wbA As Workbook, wbB As Workbook
    Set wbA = Workbooks.Open(Filename:=pathA, UpdateLinks:=0, ReadOnly:=True)
    Set wbB = Workbooks.Open(Filename:=pathB, UpdateLinks:=0, ReadOnly:=True)
    CompareTwoSnapshots = BuildCompareReport(wbA, wbB, ExtractSnapshotStamp(pathA), ExtractSnapshotStamp(pathB), wbA)
    wbA.Close False: wbB.Close False
End Function

Private Function BuildCompareReport(ByVal wbA As Workbook, ByVal wbB As Workbook, ByVal nameA As String, ByVal nameB As String, ByVal reportBaseWb As Workbook) As String
    Dim diff As Object
    Set diff = CollectDiffs(wbA, wbB, nameA, nameB)
    BuildCompareReport = modSnapshotReport.GenerateDiffReport(diff, reportBaseWb, nameA, nameB)
End Function

Private Function CollectDiffs(ByVal wbA As Workbook, ByVal wbB As Workbook, ByVal nameA As String, ByVal nameB As String) As Object
    Dim d As Object, ws As Worksheet
    Set d = CreateObject("Scripting.Dictionary")
    Set d("Rows") = New Collection
    For Each ws In wbA.Worksheets
        If ws.Name <> META_SHEET Then
            If WorksheetByName(wbB, ws.Name) Is Nothing Then d("Rows").Add Array("Removed Sheet", ws.Name, "", "", "", "", "") Else CompareSheet ws, WorksheetByName(wbB, ws.Name), d("Rows")
        End If
    Next ws
    For Each ws In wbB.Worksheets
        If ws.Name <> META_SHEET Then If WorksheetByName(wbA, ws.Name) Is Nothing Then d("Rows").Add Array("Added Sheet", ws.Name, "", "", "", "", "")
    Next ws
    Set CollectDiffs = d
End Function

Private Sub CompareSheet(ByVal wsA As Worksheet, ByVal wsB As Worksheet, ByVal rows As Collection)
    Dim rgA As Range, rgB As Range, maxRow As Long, maxCol As Long
    Set rgA = modSnapshotUtils.GetRealUsedRange(wsA)
    Set rgB = modSnapshotUtils.GetRealUsedRange(wsB)
    maxRow = MaxLng(GetRangeLastRow(rgA), GetRangeLastRow(rgB))
    maxCol = MaxLng(GetRangeLastCol(rgA), GetRangeLastCol(rgB))
    If maxRow = 0 Or maxCol = 0 Then Exit Sub
    Dim br As Long, bc As Long
    For br = 1 To maxRow Step BLOCK_SIZE
        For bc = 1 To maxCol Step BLOCK_SIZE
            CompareBlock wsA, wsB, br, bc, maxRow, maxCol, rows
        Next bc
    Next br
End Sub

Private Sub CompareBlock(ByVal wsA As Worksheet, ByVal wsB As Worksheet, ByVal startR As Long, ByVal startC As Long, ByVal maxRow As Long, ByVal maxCol As Long, ByVal rows As Collection)
    Dim endR As Long, endC As Long
    endR = WorksheetFunction.Min(startR + BLOCK_SIZE - 1, maxRow)
    endC = WorksheetFunction.Min(startC + BLOCK_SIZE - 1, maxCol)
    Dim rngA As Range, rngB As Range, valsA As Variant, valsB As Variant, frmA As Variant, frmB As Variant
    Set rngA = wsA.Range(wsA.Cells(startR, startC), wsA.Cells(endR, endC))
    Set rngB = wsB.Range(wsB.Cells(startR, startC), wsB.Cells(endR, endC))
    valsA = rngA.Value2: valsB = rngB.Value2: frmA = rngA.Formula: frmB = rngB.Formula
    If BlockArraysEqual(valsA, valsB) And BlockArraysEqual(frmA, frmB) Then Exit Sub
    Dim r As Long, c As Long, addr As String
    For r = 1 To UBound(valsA, 1)
        For c = 1 To UBound(valsA, 2)
            If CStr(valsA(r, c)) <> CStr(valsB(r, c)) Or CStr(frmA(r, c)) <> CStr(frmB(r, c)) Then
                addr = wsA.Cells(startR + r - 1, startC + c - 1).Address(False, False)
                rows.Add Array("Cell Changed", wsA.Name, addr, CStr(valsA(r, c)), CStr(valsB(r, c)), CStr(frmA(r, c)), CStr(frmB(r, c)))
            End If
        Next c
    Next r
End Sub

Private Function BlockArraysEqual(ByVal dataA As Variant, ByVal dataB As Variant) As Boolean
    Dim r As Long, c As Long

    For r = 1 To UBound(dataA, 1)
        For c = 1 To UBound(dataA, 2)
            If CStr(dataA(r, c)) <> CStr(dataB(r, c)) Then Exit Function
        Next c
    Next r

    BlockArraysEqual = True
End Function

Private Sub SaveSnapshotMeta(ByVal snapWb As Workbook, ByVal sourcePath As String)
    Dim meta As Worksheet, ws As Worksheet, r As Long
    On Error Resume Next: Set meta = snapWb.Worksheets(META_SHEET): On Error GoTo 0
    If meta Is Nothing Then Set meta = snapWb.Worksheets.Add(After:=snapWb.Worksheets(snapWb.Worksheets.Count))
    meta.Name = META_SHEET: meta.Cells.Clear
    meta.Cells(1, 1).Value = "OriginalPath": meta.Cells(1, 2).Value = sourcePath
    meta.Cells(2, 1).Value = "OriginalName": meta.Cells(2, 2).Value = snapWb.Name
    meta.Cells(3, 1).Value = "SnapshotTimestamp": meta.Cells(3, 2).Value = Format$(Now, "yymmddhhnnss")
    meta.Cells(4, 1).Value = "GeneratedUtc": meta.Cells(4, 2).Value = Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")
    r = 6
    For Each ws In snapWb.Worksheets
        If ws.Name <> META_SHEET Then
            meta.Cells(r, 1).Value = "Sheet"
            meta.Cells(r, 2).Value = ws.Name
            r = r + 1
        End If
    Next ws
    meta.Visible = xlSheetVisible
End Sub

Private Sub BreakExternalLinks(ByVal wb As Workbook)
    Dim links As Variant, i As Long
    links = wb.LinkSources(xlLinkTypeExcelLinks)
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            wb.BreakLink Name:=CStr(links(i)), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub
Private Function WorksheetByName(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    On Error Resume Next
    Set WorksheetByName = wb.Worksheets(nm)
    On Error GoTo 0
End Function
Private Function BuildSnapshotFolder(ByVal wb As Workbook) As String
    BuildSnapshotFolder = wb.Path & "\.snapshot\" & modSnapshotStorage.CleanFileName(Left$(wb.Name, InStrRev(wb.Name, ".") - 1))
End Function
Private Function SnapshotBaseName(ByVal wb As Workbook) As String
    SnapshotBaseName = modSnapshotStorage.CleanFileName(Left$(wb.Name, InStrRev(wb.Name, ".") - 1))
End Function
Private Function ExtractSnapshotStamp(ByVal p As String) As String
    Dim fn As String, x As Long
    fn = Mid$(p, InStrRev(p, "\\") + 1)
    x = InStr(1, fn, "_snapshot_")
    If x > 0 Then ExtractSnapshotStamp = Mid$(fn, x + 10, 12) Else ExtractSnapshotStamp = "snapshot"
End Function
Private Function MaxLng(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLng = a Else MaxLng = b
End Function

Private Function GetRangeLastRow(ByVal rg As Range) As Long
    If rg Is Nothing Then Exit Function
    GetRangeLastRow = rg.Row + rg.Rows.Count - 1
End Function
Private Function GetRangeLastCol(ByVal rg As Range) As Long
    If rg Is Nothing Then Exit Function
    GetRangeLastCol = rg.Column + rg.Columns.Count - 1
End Function
