Attribute VB_Name = "modSnapshotReport"
Option Explicit

Public Function GenerateDiffReport(ByVal d As Object, ByVal sourceWb As Workbook, ByVal labelA As String, ByVal labelB As String) As String
    Dim rWb As Workbook: Set rWb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet, i As Long, outPath As String, folderPath As String, baseName As String
    Set ws = rWb.Worksheets(1): ws.Name = "Summary"
    ws.Range("A1:B5").Value = Array(Array("Source Workbook", sourceWb.FullName), Array("Side A", labelA), Array("Side B", labelB), Array("Differences", d("Rows").Count), Array("Generated", Now))

    Set ws = rWb.Worksheets.Add(After:=rWb.Worksheets(rWb.Worksheets.Count)): ws.Name = "Differences"
    ws.Range("A1:G1").Value = Array("Difference Type", "Worksheet", "Cell", "Snapshot A Value", "Snapshot B Value", "Snapshot A Formula", "Snapshot B Formula")
    For i = 1 To d("Rows").Count
        ws.Cells(i + 1, 1).Resize(1, 7).Value = d("Rows")(i)
    Next i

    folderPath = sourceWb.Path & "\.snapshot\" & modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1))
    modSnapshotStorage.EnsureFolderPath folderPath
    baseName = modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1)) & "_compare_" & labelA & "_vs_" & labelB & "_report_" & Format$(Now, "yymmddhhnnss") & ".xlsx"
    outPath = folderPath & "\" & baseName
    Application.DisplayAlerts = False
    rWb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    rWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenerateDiffReport = outPath
End Function
