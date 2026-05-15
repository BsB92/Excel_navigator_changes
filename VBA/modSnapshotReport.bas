Attribute VB_Name = "modSnapshotReport"
Option Explicit

Public Function GenerateDiffReport(ByVal d As Object, ByVal sourceWb As Workbook, ByVal labelA As String, ByVal labelB As String, Optional ByVal fileA As String = "", Optional ByVal fileB As String = "") As String
    Dim rWb As Workbook: Set rWb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet, i As Long, j As Long, outPath As String, folderPath As String, baseName As String
    Dim rowData As Variant, outRow(1 To 1, 1 To 7) As Variant
    Dim resolvedFileA As String, resolvedFileB As String, compareMode As String

    resolvedFileA = ResolveDisplayFileName(fileA, labelA)
    resolvedFileB = ResolveDisplayFileName(fileB, labelB)
    compareMode = modNavigatorSettings.GetSnapshotCompareMode()

    Set ws = rWb.Worksheets(1): ws.Name = "Summary"
    WriteSummary ws, d, sourceWb, compareMode, labelA, labelB, resolvedFileA, resolvedFileB

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

    ApplyReportLayout rWb.Worksheets("Summary"), ws

    folderPath = ResolveReportFolder(sourceWb)
    modSnapshotStorage.EnsureFolderPath folderPath
    baseName = BuildReportFileName(sourceWb.Name, compareMode, resolvedFileA, resolvedFileB)
    outPath = folderPath & "\" & baseName
    rWb.Worksheets("Summary").Activate
    rWb.Worksheets("Summary").Range("A1").Select
    Application.DisplayAlerts = False
    rWb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
    rWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenerateDiffReport = outPath
End Function


Private Sub WriteSummary(ByVal ws As Worksheet, ByVal d As Object, ByVal sourceWb As Workbook, ByVal compareMode As String, ByVal labelA As String, ByVal labelB As String, ByVal resolvedFileA As String, ByVal resolvedFileB As String)
    Dim totalDiff As Long
    Dim changedCells As Long
    Dim addedSheets As Long
    Dim removedSheets As Long
    Dim changedSheets As Long
    Dim warningCount As Long
    Dim rowNo As Long

    BuildDiffStats d, totalDiff, changedCells, addedSheets, removedSheets, changedSheets, warningCount

    rowNo = 1
    WriteSummaryRow ws, rowNo, "Source Workbook", sourceWb.FullName
    WriteSummaryRow ws, rowNo, "Compare Mode", compareMode
    WriteSummaryRow ws, rowNo, "Side A", labelA
    WriteSummaryRow ws, rowNo, "Side B", labelB
    WriteSummaryRow ws, rowNo, "File A name", GetShortFileLabel(resolvedFileA)
    WriteSummaryRow ws, rowNo, "File B name", GetShortFileLabel(resolvedFileB)
    WriteSummaryRow ws, rowNo, "File A path", resolvedFileA
    WriteSummaryRow ws, rowNo, "File B path", resolvedFileB
    WriteSummaryRow ws, rowNo, "Compared sheets", CStr(CountComparedSheets(d))
    WriteSummaryRow ws, rowNo, "Sheets with differences", CStr(changedSheets)
    WriteSummaryRow ws, rowNo, "Added sheets", CStr(addedSheets)
    WriteSummaryRow ws, rowNo, "Removed sheets", CStr(removedSheets)
    WriteSummaryRow ws, rowNo, "Changed cells", CStr(changedCells)
    WriteSummaryRow ws, rowNo, "Warnings/other change types", CStr(warningCount)
    WriteSummaryRow ws, rowNo, "Total differences", CStr(totalDiff)
    WriteSummaryRow ws, rowNo, "Generated", CStr(Now)
    WriteSummaryRow ws, rowNo, "Legend", ""
    WriteSummaryRow ws, rowNo, "Changed formula fragment", "bold + red text"
    ws.Range("A" & rowNo).Characters(1, Len(ws.Range("A" & rowNo).Value2)).Font.Bold = True
    ws.Range("A" & rowNo).Characters(1, Len(ws.Range("A" & rowNo).Value2)).Font.Color = RGB(192, 0, 0)
    WriteSummaryRow ws, rowNo, "Compare Mode (report)", compareMode
End Sub

Private Sub WriteSummaryRow(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal keyText As String, ByVal valueText As String)
    ws.Cells(rowNo, 1).Value = keyText
    ws.Cells(rowNo, 2).Value = valueText
    rowNo = rowNo + 1
End Sub

Private Sub BuildDiffStats(ByVal d As Object, ByRef totalDiff As Long, ByRef changedCells As Long, ByRef addedSheets As Long, ByRef removedSheets As Long, ByRef changedSheets As Long, ByRef warningCount As Long)
    Dim i As Long
    Dim rowData As Variant
    Dim diffType As String
    Dim sheetName As String
    Dim seen As Object

    totalDiff = d("Rows").Count
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To d("Rows").Count
        rowData = d("Rows")(i)
        diffType = LCase$(Trim$(CStr(rowData(0))))
        sheetName = Trim$(CStr(rowData(1)))

        If Len(sheetName) > 0 Then
            If Not seen.Exists(sheetName) Then seen(sheetName) = True
        End If

        If diffType = "added sheet" Then
            addedSheets = addedSheets + 1
        ElseIf diffType = "removed sheet" Then
            removedSheets = removedSheets + 1
        ElseIf InStr(1, diffType, "cell", vbTextCompare) > 0 Or InStr(1, diffType, "formula", vbTextCompare) > 0 Or InStr(1, diffType, "value", vbTextCompare) > 0 Then
            changedCells = changedCells + 1
        Else
            warningCount = warningCount + 1
        End If
    Next i

    changedSheets = seen.Count
End Sub

Private Function CountComparedSheets(ByVal d As Object) As Long
    Dim i As Long
    Dim rowData As Variant
    Dim diffType As String
    Dim sheetName As String
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To d("Rows").Count
        rowData = d("Rows")(i)
        diffType = LCase$(Trim$(CStr(rowData(0))))
        sheetName = Trim$(CStr(rowData(1)))
        If Len(sheetName) > 0 Then
            If diffType <> "added sheet" And diffType <> "removed sheet" Then
                If Not seen.Exists(sheetName) Then seen(sheetName) = True
            End If
        End If
    Next i
    CountComparedSheets = seen.Count
End Function


Private Sub ApplyReportLayout(ByVal summaryWs As Worksheet, ByVal diffWs As Worksheet)
    summaryWs.Columns("A:B").EntireColumn.AutoFit

    diffWs.Rows(1).Font.Bold = True
    diffWs.Columns("A:G").EntireColumn.AutoFit
    EnsureMinHeaderWidth diffWs, 1, 1
    EnsureMinHeaderWidth diffWs, 2, 1
    EnsureMinHeaderWidth diffWs, 3, 1
    EnsureMinHeaderWidth diffWs, 4, 1
    EnsureMinHeaderWidth diffWs, 5, 1
    EnsureMinHeaderWidth diffWs, 6, 1
    EnsureMinHeaderWidth diffWs, 7, 1
End Sub

Private Sub EnsureMinHeaderWidth(ByVal targetWs As Worksheet, ByVal colIndex As Long, ByVal headerRow As Long)
    Dim headerLen As Long
    Dim minWidth As Double

    headerLen = Len(CStr(targetWs.Cells(headerRow, colIndex).Value2))
    minWidth = headerLen + 2
    If minWidth < 14 Then minWidth = 14

    If targetWs.Columns(colIndex).ColumnWidth < minWidth Then
        targetWs.Columns(colIndex).ColumnWidth = minWidth
    End If
End Sub


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



Private Function BuildReportFileName(ByVal sourceWorkbookName As String, ByVal compareMode As String, ByVal fileA As String, ByVal fileB As String) As String
    Dim aShort As String
    Dim bShort As String

    aShort = GetShortFileLabel(fileA)
    bShort = GetShortFileLabel(fileB)

    BuildReportFileName = "CompareMode_" & modSnapshotStorage.CleanFileName(compareMode) & _
        "_" & modSnapshotStorage.CleanFileName(aShort) & _
        "_vs_" & modSnapshotStorage.CleanFileName(bShort) & _
        "_" & Format$(Now, "yymmdd_hhnnss") & ".xlsx"
End Function


Private Function ResolveReportFolder(ByVal sourceWb As Workbook) As String
    If InStr(1, sourceWb.Path, "\.snapshot\", vbTextCompare) > 0 Then
        ResolveReportFolder = sourceWb.Path
    Else
        ResolveReportFolder = sourceWb.Path & "\.snapshot\" & modSnapshotStorage.CleanFileName(Left$(sourceWb.Name, InStrRev(sourceWb.Name, ".") - 1))
    End If
End Function
