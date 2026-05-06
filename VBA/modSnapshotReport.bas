Attribute VB_Name = "modSnapshotReport"
Option Explicit

Public Sub GenerateDiffReport(ByVal d As clsSnapshotDiffResult, ByVal sourceWb As Workbook, ByVal snapshotPath As String)
    Dim rWb As Workbook: Set rWb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet, i As Long
    Set ws = rWb.Worksheets(1): ws.Name = "Summary"
    ws.Range("A1:B6").Value = Array( _
        Array("Snapshot", snapshotPath), _
        Array("Workbook", sourceWb.FullName), _
        Array("Formula changes", d.FormulaChanges.Count), _
        Array("Value changes", d.ValueChanges.Count), _
        Array("Generated", Now), _
        Array("Tool", "ExcelNavigator Snapshot Compare"))

    Set ws = rWb.Worksheets.Add(After:=rWb.Worksheets(rWb.Worksheets.Count)): ws.Name = "Changed Formulas"
    ws.Range("A1:D1").Value = Array("Sheet", "Cell", "Old Formula", "New Formula")
    For i = 1 To d.FormulaChanges.Count
        ws.Cells(i + 1, 1).Resize(1, 4).Value = d.FormulaChanges(i)
    Next i

    Set ws = rWb.Worksheets.Add(After:=rWb.Worksheets(rWb.Worksheets.Count)): ws.Name = "Changed Values"
    ws.Range("A1:D1").Value = Array("Sheet", "Cell", "Old Value", "New Value")
    For i = 1 To d.ValueChanges.Count
        ws.Cells(i + 1, 1).Resize(1, 4).Value = d.ValueChanges(i)
    Next i
End Sub
