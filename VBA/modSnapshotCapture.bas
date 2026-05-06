Attribute VB_Name = "modSnapshotCapture"
Option Explicit

Public Function CaptureWorkbookSnapshot(ByVal wb As Workbook) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    Set s("Worksheets") = CreateObject("Scripting.Dictionary")
    Set s("Cells") = CreateObject("Scripting.Dictionary")
    Set s("NamedRanges") = CreateObject("Scripting.Dictionary")
    Set s("ExternalLinks") = CreateObject("Scripting.Dictionary")
    Set s("VbaModules") = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, nm As Name, i As Long
    s("WorkbookName") = wb.Name
    s("WorkbookPath") = wb.FullName
    s("CreatedAtUtc") = Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")
    s("CreatedBy") = Environ$("USERNAME")
    s("MachineName") = Environ$("COMPUTERNAME")
    s("ExcelVersion") = Application.Version
    s("SnapshotSchemaVersion") = "1.0"

    For Each ws In wb.Worksheets
        s("Worksheets")(ws.Name) = CStr(ws.Index) & "|" & CStr(ws.Visible)
        CaptureSheetCells ws, s("Cells")
    Next ws

    For Each nm In wb.Names
        s("NamedRanges")(nm.Name) = "Workbook|" & nm.RefersTo
    Next nm

    Dim links As Variant
    links = wb.LinkSources(xlLinkTypeExcelLinks)
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            s("ExternalLinks")(CStr(links(i))) = True
        Next i
    End If
    CaptureVbaHashes wb, s
    Set CaptureWorkbookSnapshot = s
End Function

Private Sub CaptureSheetCells(ByVal ws As Worksheet, ByVal cellsDict As Object)
    Dim rg As Range, vals As Variant, fml As Variant
    Dim r As Long, c As Long, key As String, v As Variant, f As String, t As String
    Set rg = modSnapshotUtils.GetRealUsedRange(ws)
    If rg Is Nothing Then Exit Sub
    vals = rg.Value2
    fml = rg.Formula
    For r = 1 To UBound(vals, 1)
        For c = 1 To UBound(vals, 2)
            v = vals(r, c)
            f = CStr(fml(r, c))
            If Len(f) > 0 Or Len(CStr(v)) > 0 Then
                key = ws.Name & "!" & rg.Cells(r, c).Address(False, False)
                t = TypeName(v)
                cellsDict(key) = f & ChrW(31) & CStr(v) & ChrW(31) & t
            End If
        Next c
    Next r
End Sub

Private Sub CaptureVbaHashes(ByVal wb As Workbook, ByVal s As Object)
    On Error GoTo NO_VBA
    Dim vbProj As Object, comp As Object
    Set vbProj = wb.VBProject
    For Each comp In vbProj.VBComponents
        s("VbaModules")(comp.Name) = CStr(Len(comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)))
    Next comp
    Exit Sub
NO_VBA:
    s("VbaModules")("__warning__") = "VBA access denied"
    Err.Clear
End Sub
