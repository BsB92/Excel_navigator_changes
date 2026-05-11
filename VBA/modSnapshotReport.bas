Attribute VB_Name = "modSnapshotReport"
Option Explicit

Public Function GenerateDiffReport(ByVal d As Object, ByVal sourceWb As Workbook, ByVal labelA As String, ByVal labelB As String) As String
    Dim rWb As Workbook: Set rWb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet, i As Long, j As Long, outPath As String, folderPath As String, baseName As String
    Dim rowData As Variant, outRow(1 To 1, 1 To 7) As Variant
    Set ws = rWb.Worksheets(1): ws.Name = "Summary"
    ws.Range("A1:B5").Value = Array(Array("Source Workbook", sourceWb.FullName), Array("Side A", labelA), Array("Side B", labelB), Array("Differences", d("Rows").Count), Array("Generated", Now))

    Set ws = rWb.Worksheets.Add(After:=rWb.Worksheets(rWb.Worksheets.Count)): ws.Name = "Differences"
    ws.Range("A1:G1").Value = Array("Difference Type", "Worksheet", "Cell", "Snapshot A Value", "Snapshot B Value", "Snapshot A Formula", "Snapshot B Formula")
    For i = 1 To d("Rows").Count
        rowData = d("Rows")(i)
        For j = 1 To 7
            outRow(1, j) = SanitizeReportCell(CStr(rowData(j - 1)))
        Next j
        ws.Cells(i + 1, 1).Resize(1, 7).Value = outRow
        ApplyFormulaDifferenceBold ws.Cells(i + 1, 6), ws.Cells(i + 1, 7)
    Next i

    folderPath = ResolveReportFolder(sourceWb)
    modSnapshotStorage.EnsureFolderPath folderPath
    baseName = modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1)) & "_compare_" & modSnapshotStorage.CleanFileName(labelA) & "_vs_" & modSnapshotStorage.CleanFileName(labelB) & "_report_" & Format$(Now, "yymmddhhnnss") & ".xlsx"
    outPath = folderPath & "\" & baseName
    Application.DisplayAlerts = False
    rWb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    rWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenerateDiffReport = outPath
End Function


Private Sub ApplyFormulaDifferenceBold(ByVal formulaCellA As Range, ByVal formulaCellB As Range)
    Dim formulaA As String, formulaB As String
    Dim diffStart As Long, diffLenA As Long, diffLenB As Long

    formulaA = DisplayCellText(formulaCellA)
    formulaB = DisplayCellText(formulaCellB)

    If Len(formulaA) = 0 Or Len(formulaB) = 0 Then Exit Sub
    If StrComp(formulaA, formulaB, vbBinaryCompare) = 0 Then Exit Sub

    FindDifferenceBounds formulaA, formulaB, diffStart, diffLenA, diffLenB
    If diffLenA > 0 Then formulaCellA.Characters(diffStart, diffLenA).Font.Bold = True
    If diffLenB > 0 Then formulaCellB.Characters(diffStart, diffLenB).Font.Bold = True
End Sub


Private Sub FindDifferenceBounds(ByVal textA As String, ByVal textB As String, ByRef startPos As Long, ByRef changeLenA As Long, ByRef changeLenB As Long)
    Dim leftMatch As Long, rightMatch As Long
    Dim maxLeft As Long, maxRight As Long

    maxLeft = IIf(Len(textA) < Len(textB), Len(textA), Len(textB))
    leftMatch = 0
    Do While leftMatch < maxLeft
        If Mid$(textA, leftMatch + 1, 1) <> Mid$(textB, leftMatch + 1, 1) Then Exit Do
        leftMatch = leftMatch + 1
    Loop

    maxRight = IIf((Len(textA) - leftMatch) < (Len(textB) - leftMatch), (Len(textA) - leftMatch), (Len(textB) - leftMatch))
    rightMatch = 0
    Do While rightMatch < maxRight
        If Mid$(textA, Len(textA) - rightMatch, 1) <> Mid$(textB, Len(textB) - rightMatch, 1) Then Exit Do
        rightMatch = rightMatch + 1
    Loop

    startPos = leftMatch + 1
    changeLenA = Len(textA) - leftMatch - rightMatch
    changeLenB = Len(textB) - leftMatch - rightMatch
End Sub


Private Function DisplayCellText(ByVal targetCell As Range) As String
    DisplayCellText = CStr(targetCell.Value2)
    If Left$(DisplayCellText, 1) = "'" Then DisplayCellText = Mid$(DisplayCellText, 2)
End Function


Private Function SanitizeReportCell(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) = "=" Or Left$(s, 1) = "+" Or Left$(s, 1) = "-" Or Left$(s, 1) = "@" Then
            SanitizeReportCell = "'" & s
            Exit Function
        End If
    End If
    SanitizeReportCell = s
End Function


Private Function ResolveReportFolder(ByVal sourceWb As Workbook) As String
    If InStr(1, sourceWb.Path, "\.snapshot\", vbTextCompare) > 0 Then
        ResolveReportFolder = sourceWb.Path
    Else
        ResolveReportFolder = sourceWb.Path & "\.snapshot\" & modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1))
    End If
End Function
