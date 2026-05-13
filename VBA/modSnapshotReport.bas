Attribute VB_Name = "modSnapshotReport"
Option Explicit

Public Function GenerateDiffReport(ByVal d As Object, ByVal sourceWb As Workbook, ByVal labelA As String, ByVal labelB As String, Optional ByVal fileA As String = "", Optional ByVal fileB As String = "") As String
    Dim rWb As Workbook: Set rWb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet, i As Long, j As Long, outPath As String, folderPath As String, baseName As String
    Dim rowData As Variant, outRow(1 To 1, 1 To 7) As Variant
    Dim resolvedFileA As String, resolvedFileB As String

    resolvedFileA = ResolveDisplayFileName(fileA, labelA)
    resolvedFileB = ResolveDisplayFileName(fileB, labelB)

    Set ws = rWb.Worksheets(1): ws.Name = "Summary"
    ws.Range("A1:B7").Value = Array(Array("Source Workbook", sourceWb.FullName), Array("Side A", labelA), Array("Side B", labelB), Array("File A", resolvedFileA), Array("File B", resolvedFileB), Array("Differences", d("Rows").Count), Array("Generated", Now))
    ws.Range("A9").Value = "Legend"
    ws.Range("A10").Value = "Changed formula fragment"
    ws.Range("A10").Characters(1, Len(ws.Range("A10").Value2)).Font.Bold = True
    ws.Range("A10").Characters(1, Len(ws.Range("A10").Value2)).Font.Color = RGB(192, 0, 0)

    Set ws = rWb.Worksheets.Add(After:=rWb.Worksheets(rWb.Worksheets.Count)): ws.Name = "Differences"
    ws.Range("A1:G1").Value = Array( _
        "Difference Type", _
        "Worksheet", _
        "Cell", _
        "A Value (" & GetShortFileLabel(resolvedFileA) & ")", _
        "B Value (" & GetShortFileLabel(resolvedFileB) & ")", _
        "A Formula (" & GetShortFileLabel(resolvedFileA) & ")", _
        "B Formula (" & GetShortFileLabel(resolvedFileB) & ")" _
    )
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
    Dim prefixLen As Long, suffixLen As Long
    Dim diffStart As Long, diffLenA As Long, diffLenB As Long

    formulaA = DisplayCellText(formulaCellA)
    formulaB = DisplayCellText(formulaCellB)

    If Len(formulaA) = 0 Or Len(formulaB) = 0 Then Exit Sub
    If Left$(formulaA, 1) <> "=" Or Left$(formulaB, 1) <> "=" Then Exit Sub
    If StrComp(formulaA, formulaB, vbBinaryCompare) = 0 Then Exit Sub

    formulaCellA.Font.Bold = False
    formulaCellB.Font.Bold = False

    prefixLen = CommonPrefixLength(formulaA, formulaB)
    suffixLen = CommonSuffixLength(formulaA, formulaB, prefixLen)

    diffStart = prefixLen + 1
    diffLenA = Len(formulaA) - prefixLen - suffixLen
    diffLenB = Len(formulaB) - prefixLen - suffixLen

    If diffLenA > 0 Then MarkChangedFragment formulaCellA, diffStart, diffLenA
    If diffLenB > 0 Then MarkChangedFragment formulaCellB, diffStart, diffLenB
End Sub


Private Function CommonPrefixLength(ByVal textA As String, ByVal textB As String) As Long
    Dim i As Long, limitLen As Long

    limitLen = Len(textA)
    If Len(textB) < limitLen Then limitLen = Len(textB)

    For i = 1 To limitLen
        If Mid$(textA, i, 1) <> Mid$(textB, i, 1) Then Exit For
    Next i

    CommonPrefixLength = i - 1
End Function


Private Function CommonSuffixLength(ByVal textA As String, ByVal textB As String, ByVal prefixLen As Long) As Long
    Dim i As Long
    Dim maxSuffix As Long

    maxSuffix = Len(textA) - prefixLen
    If Len(textB) - prefixLen < maxSuffix Then maxSuffix = Len(textB) - prefixLen

    For i = 1 To maxSuffix
        If Mid$(textA, Len(textA) - i + 1, 1) <> Mid$(textB, Len(textB) - i + 1, 1) Then Exit For
    Next i

    CommonSuffixLength = i - 1
End Function


Private Sub MarkChangedFragment(ByVal targetCell As Range, ByVal startPos As Long, ByVal fragmentLen As Long)
    With targetCell.Characters(startPos, fragmentLen).Font
        .Bold = True
        .Color = RGB(192, 0, 0)
    End With
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




Private Function GetShortFileLabel(ByVal valueText As String) As String
    Dim normalized As String
    Dim p As Long

    normalized = Replace$(Trim$(valueText), "/", "\")
    p = InStrRev(normalized, "\")

    If p > 0 And p < Len(normalized) Then
        GetShortFileLabel = Mid$(normalized, p + 1)
    Else
        GetShortFileLabel = valueText
    End If
End Function


Private Function ResolveDisplayFileName(ByVal filePath As String, ByVal fallbackLabel As String) As String
    If Len(Trim$(filePath)) > 0 Then
        ResolveDisplayFileName = filePath
    Else
        ResolveDisplayFileName = fallbackLabel
    End If
End Function


Private Function ResolveReportFolder(ByVal sourceWb As Workbook) As String
    If InStr(1, sourceWb.Path, "\.snapshot\", vbTextCompare) > 0 Then
        ResolveReportFolder = sourceWb.Path
    Else
        ResolveReportFolder = sourceWb.Path & "\.snapshot\" & modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1))
    End If
End Function
