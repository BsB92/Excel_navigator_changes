Attribute VB_Name = "modCF2Static"
Option Explicit

Private Const CF2STATIC_SUFFIX As String = "_without_conditional_formatting"
Private Const XL_DATABAR_TYPE As Long = 4
Private Const XL_ICONSET_TYPE As Long = 6

Public Sub RunCF2Static(ByVal ownerForm As Object)
    Dim selectedWorkbooks As Collection
    Dim errors As Collection
    Dim targetFolder As String
    Dim item As Variant
    Dim srcWb As Workbook
    Dim outPath As String
    Dim detail As String
    Dim okCount As Long
    Dim failCount As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevAskToUpdateLinks As Boolean
    Dim prevCalculation As XlCalculation
    Dim calculationChanged As Boolean
    Dim appStateCaptured As Boolean
    Dim summaryText As String
    Dim i As Long

    On Error GoTo FatalError

    If ownerForm Is Nothing Then Exit Sub
    If Not CBool(ownerForm.tglBatchMode.Value) Then
        MsgBox "Turn ON Selection mode and select at least one file.", vbExclamation, "CF2Static"
        Exit Sub
    End If

    Set errors = New Collection
    Set selectedWorkbooks = GetSelectedWorkbooks(ownerForm, errors)
    If selectedWorkbooks.Count = 0 Then
        MsgBox "Nothing selected.", vbExclamation, "CF2Static"
        Exit Sub
    End If

    targetFolder = PickCF2StaticTargetFolder()
    If Len(targetFolder) = 0 Then Exit Sub

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts
    prevAskToUpdateLinks = Application.AskToUpdateLinks
    prevCalculation = Application.Calculation
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    On Error Resume Next
    Application.Calculation = xlCalculationManual
    calculationChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo FatalError

    For Each item In selectedWorkbooks
        Set srcWb = item
        outPath = vbNullString
        detail = vbNullString
        Application.StatusBar = "CF2Static: " & srcWb.Name

        If Len(srcWb.Path) = 0 Then
            failCount = failCount + 1
            errors.Add srcWb.Name & ": source workbook must be saved first."
        Else
            outPath = BuildCF2StaticOutputPath(targetFolder, srcWb.Name)
            If OutputFileExists(outPath) Then
                failCount = failCount + 1
                errors.Add srcWb.Name & ": target file already exists: " & outPath
            ElseIf ConvertCopiedWorkbook(srcWb, outPath, detail) Then
                okCount = okCount + 1
            Else
                failCount = failCount + 1
                If Len(detail) = 0 Then detail = "Unknown conversion error."
                errors.Add srcWb.Name & ": " & detail
            End If
        End If
    Next item

CleanExit:
    On Error Resume Next
    Application.StatusBar = False
    If appStateCaptured Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEnableEvents
        Application.DisplayAlerts = prevDisplayAlerts
        Application.AskToUpdateLinks = prevAskToUpdateLinks
        If calculationChanged Then Application.Calculation = prevCalculation
    End If
    On Error GoTo 0

    summaryText = "CF2Static finished." & vbCrLf & _
                  "Processed successfully: " & CStr(okCount) & vbCrLf & _
                  "Failed: " & CStr(failCount)

    If Not errors Is Nothing Then
        If errors.Count > 0 Then
            summaryText = summaryText & vbCrLf & vbCrLf & "Errors / limitations:"
            For i = 1 To errors.Count
                summaryText = summaryText & vbCrLf & "- " & CStr(errors(i))
            Next i
        End If
    End If

    MsgBox summaryText, IIf(failCount > 0, vbExclamation, vbInformation), "CF2Static"
    On Error Resume Next
    ownerForm.RestoreNavigatorToFront
    On Error GoTo 0
    Exit Sub

FatalError:
    If errors Is Nothing Then Set errors = New Collection
    failCount = failCount + 1
    errors.Add "Fatal error: " & Err.Description
    Resume CleanExit
End Sub

Private Function GetSelectedWorkbooks(ByVal ownerForm As Object, ByVal errors As Collection) As Collection
    Dim result As Collection
    Dim i As Long
    Dim fullPath As String
    Dim displayName As String
    Dim wb As Workbook

    Set result = New Collection

    For i = 1 To ownerForm.lstWorkbooks.ListCount - 1
        If ownerForm.lstWorkbooks.Selected(i) Then
            fullPath = CStr(ownerForm.lstWorkbooks.List(i, 3))
            displayName = CStr(ownerForm.lstWorkbooks.List(i, 0))
            If Left$(displayName, 2) = "> " Then displayName = Mid$(displayName, 3)

            Set wb = FindOpenWorkbook(fullPath, displayName)
            If wb Is Nothing Then
                errors.Add displayName & ": workbook is no longer open."
            Else
                result.Add wb
            End If
        End If
    Next i

    Set GetSelectedWorkbooks = result
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String, ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If Len(fullPath) > 0 Then
            If StrComp(wb.fullName, fullPath, vbTextCompare) = 0 Then
                Set FindOpenWorkbook = wb
                Exit Function
            End If
        ElseIf StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function PickCF2StaticTargetFolder() As String
    Dim fd As FileDialog
    Dim initialFolder As String

    initialFolder = modNavigatorSettings.ResolveInitialFolder(ThisWorkbook.Path)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "CF2Static - choose target folder for copied files"
        .AllowMultiSelect = False
        On Error Resume Next
        If Len(initialFolder) > 0 Then .InitialFileName = EnsureTrailingSeparator(initialFolder)
        On Error GoTo 0

        If .Show = -1 Then PickCF2StaticTargetFolder = CStr(.SelectedItems(1))
    End With
End Function

Private Function EnsureTrailingSeparator(ByVal folderPath As String) As String
    Dim result As String

    result = Trim$(folderPath)
    If Len(result) = 0 Then Exit Function
    If Right$(result, 1) <> "\" And Right$(result, 1) <> "/" Then result = result & Application.PathSeparator
    EnsureTrailingSeparator = result
End Function

Private Function BuildCF2StaticOutputPath(ByVal targetFolder As String, ByVal sourceName As String) As String
    Dim dotPos As Long
    Dim baseName As String
    Dim extension As String
    Dim normalizedFolder As String

    normalizedFolder = Trim$(targetFolder)
    Do While Right$(normalizedFolder, 1) = "\" Or Right$(normalizedFolder, 1) = "/"
        normalizedFolder = Left$(normalizedFolder, Len(normalizedFolder) - 1)
    Loop

    dotPos = InStrRev(sourceName, ".")
    If dotPos > 0 Then
        baseName = Left$(sourceName, dotPos - 1)
        extension = Mid$(sourceName, dotPos)
    Else
        baseName = sourceName
        extension = ".xlsx"
    End If

    BuildCF2StaticOutputPath = normalizedFolder & Application.PathSeparator & baseName & CF2STATIC_SUFFIX & extension
End Function

Private Function OutputFileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    OutputFileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    On Error GoTo 0
End Function

Private Function ConvertCopiedWorkbook(ByVal srcWb As Workbook, ByVal outPath As String, ByRef detail As String) As Boolean
    Dim copiedWb As Workbook
    Dim expectedWorksheets As Long
    Dim processedWorksheets As Long

    On Error GoTo EH

    srcWb.SaveCopyAs outPath
    Set copiedWb = Workbooks.Open(fileName:=outPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, AddToMru:=False)

    expectedWorksheets = copiedWb.Worksheets.Count
    If Not ConvertWorkbookConditionalFormatting(copiedWb, processedWorksheets, detail) Then GoTo Failed

    If processedWorksheets <> expectedWorksheets Then
        detail = "Worksheet validation failed: expected " & CStr(expectedWorksheets) & _
                 ", processed " & CStr(processedWorksheets) & "."
        GoTo Failed
    End If

    If WorkbookHasConditionalFormatting(copiedWb, detail) Then GoTo Failed

    copiedWb.Save
    copiedWb.Close saveChanges:=False
    Set copiedWb = Nothing
    ConvertCopiedWorkbook = True
    Exit Function

Failed:
    On Error Resume Next
    If Not copiedWb Is Nothing Then copiedWb.Close saveChanges:=False
    Set copiedWb = Nothing
    DeleteFailedCopy outPath
    On Error GoTo 0
    ConvertCopiedWorkbook = False
    Exit Function

EH:
    detail = "Conversion error: " & Err.Description
    Resume Failed
End Function

Private Function ConvertWorkbookConditionalFormatting(ByVal wb As Workbook, ByRef processedWorksheets As Long, ByRef detail As String) As Boolean
    Dim ws As Worksheet

    processedWorksheets = 0

    For Each ws In wb.Worksheets
        If Not ConvertWorksheetConditionalFormatting(ws, detail) Then
            detail = ws.Name & ": " & detail
            Exit Function
        End If
        processedWorksheets = processedWorksheets + 1
    Next ws

    ConvertWorkbookConditionalFormatting = True
End Function

Private Function ConvertWorksheetConditionalFormatting(ByVal ws As Worksheet, ByRef detail As String) As Boolean
    Dim cfCells As Range
    Dim c As Range
    Dim snapshots As Object
    Dim addressKey As Variant
    Dim beforeSignature As String
    Dim afterSignature As String
    Dim unsupportedDetails As String

    On Error GoTo EH

    Set cfCells = GetConditionalFormattingCells(ws)
    If cfCells Is Nothing Then
        ConvertWorksheetConditionalFormatting = True
        Exit Function
    End If

    If HasUnsupportedConditionalFormatting(cfCells, unsupportedDetails) Then
        detail = unsupportedDetails
        Exit Function
    End If

    Set snapshots = CreateObject("Scripting.Dictionary")

    ' Exact DisplayFormat can differ cell by cell because formulas, color scales and priorities
    ' are evaluated per cell. For that reason only cells actually covered by Conditional
    ' Formatting are processed, but those cells are intentionally handled individually.
    For Each c In cfCells.Cells
        beforeSignature = BuildDisplayFormatSignature(c)
        snapshots(c.Address(False, False, xlA1)) = beforeSignature
        ApplyDisplayFormatAsStatic c
    Next c

    ws.Cells.FormatConditions.Delete

    For Each addressKey In snapshots.Keys
        afterSignature = BuildStaticFormatSignature(ws.Range(CStr(addressKey)))
        If StrComp(CStr(snapshots(addressKey)), afterSignature, vbBinaryCompare) <> 0 Then
            detail = "Visual validation failed at " & CStr(addressKey) & "."
            Exit Function
        End If
    Next addressKey

    Set cfCells = GetConditionalFormattingCells(ws)
    If Not cfCells Is Nothing Then
        detail = "Conditional Formatting rules remain after conversion."
        Exit Function
    End If

    ConvertWorksheetConditionalFormatting = True
    Exit Function

EH:
    detail = "Worksheet conversion error: " & Err.Description
    ConvertWorksheetConditionalFormatting = False
End Function

Private Function GetConditionalFormattingCells(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set GetConditionalFormattingCells = ws.Cells.SpecialCells(xlCellTypeAllFormatConditions)
    On Error GoTo 0
End Function

Private Function HasUnsupportedConditionalFormatting(ByVal cfCells As Range, ByRef details As String) As Boolean
    Dim area As Range
    Dim fc As Object

    On Error Resume Next
    For Each area In cfCells.Areas
        For Each fc In area.FormatConditions
            If CLng(fc.Type) = XL_DATABAR_TYPE Then
                details = "Data Bars cannot be reproduced as ordinary static cell formatting by Excel VBA. File was not saved as converted."
                HasUnsupportedConditionalFormatting = True
                Exit Function
            ElseIf CLng(fc.Type) = XL_ICONSET_TYPE Then
                details = "Icon Sets cannot be reproduced as ordinary static cell formatting by Excel VBA. File was not saved as converted."
                HasUnsupportedConditionalFormatting = True
                Exit Function
            End If
        Next fc
    Next area
    On Error GoTo 0
End Function

Private Sub ApplyDisplayFormatAsStatic(ByVal targetCell As Range)
    Dim visibleFormat As Object
    Dim borderIndexes As Variant
    Dim borderIndex As Variant

    Set visibleFormat = targetCell.DisplayFormat

    With targetCell.Interior
        .Pattern = visibleFormat.Interior.Pattern
        .Color = visibleFormat.Interior.Color
        .PatternColor = visibleFormat.Interior.PatternColor
        On Error Resume Next
        .TintAndShade = visibleFormat.Interior.TintAndShade
        .PatternTintAndShade = visibleFormat.Interior.PatternTintAndShade
        On Error GoTo 0
    End With

    With targetCell.Font
        .Name = visibleFormat.Font.Name
        .Size = visibleFormat.Font.Size
        .Bold = visibleFormat.Font.Bold
        .Italic = visibleFormat.Font.Italic
        .Underline = visibleFormat.Font.Underline
        .Strikethrough = visibleFormat.Font.Strikethrough
        .Color = visibleFormat.Font.Color
        On Error Resume Next
        .TintAndShade = visibleFormat.Font.TintAndShade
        On Error GoTo 0
    End With

    targetCell.NumberFormat = visibleFormat.NumberFormat

    borderIndexes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalDown, xlDiagonalUp)
    For Each borderIndex In borderIndexes
        CopyVisibleBorder visibleFormat, targetCell, CLng(borderIndex)
    Next borderIndex
End Sub

Private Sub CopyVisibleBorder(ByVal visibleFormat As Object, ByVal targetCell As Range, ByVal borderIndex As Long)
    On Error Resume Next
    With targetCell.Borders(borderIndex)
        .LineStyle = visibleFormat.Borders(borderIndex).LineStyle
        .Weight = visibleFormat.Borders(borderIndex).Weight
        .Color = visibleFormat.Borders(borderIndex).Color
        .TintAndShade = visibleFormat.Borders(borderIndex).TintAndShade
    End With
    On Error GoTo 0
End Sub

Private Function BuildDisplayFormatSignature(ByVal c As Range) As String
    Dim visibleFormat As Object

    Set visibleFormat = c.DisplayFormat
    BuildDisplayFormatSignature = BuildFormatSignature(visibleFormat)
End Function

Private Function BuildStaticFormatSignature(ByVal c As Range) As String
    BuildStaticFormatSignature = BuildFormatSignature(c)
End Function

Private Function BuildFormatSignature(ByVal formattedObject As Object) As String
    Dim result As String
    Dim borderIndexes As Variant
    Dim borderIndex As Variant

    On Error GoTo EH

    result = "IP=" & CStr(formattedObject.Interior.Pattern) & _
             "|IC=" & CStr(formattedObject.Interior.Color) & _
             "|IPC=" & CStr(formattedObject.Interior.PatternColor) & _
             "|FN=" & CStr(formattedObject.Font.Name) & _
             "|FS=" & CStr(formattedObject.Font.Size) & _
             "|FB=" & CStr(formattedObject.Font.Bold) & _
             "|FI=" & CStr(formattedObject.Font.Italic) & _
             "|FU=" & CStr(formattedObject.Font.Underline) & _
             "|FST=" & CStr(formattedObject.Font.Strikethrough) & _
             "|FC=" & CStr(formattedObject.Font.Color) & _
             "|NF=" & CStr(formattedObject.NumberFormat)

    On Error Resume Next
    result = result & "|ITS=" & CStr(formattedObject.Interior.TintAndShade) & _
             "|IPTS=" & CStr(formattedObject.Interior.PatternTintAndShade) & _
             "|FTS=" & CStr(formattedObject.Font.TintAndShade)
    On Error GoTo EH

    borderIndexes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalDown, xlDiagonalUp)
    For Each borderIndex In borderIndexes
        result = result & BorderSignature(formattedObject, CLng(borderIndex))
    Next borderIndex

    BuildFormatSignature = result
    Exit Function

EH:
    BuildFormatSignature = "#FORMAT_SIGNATURE_ERROR#" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function BorderSignature(ByVal formattedObject As Object, ByVal borderIndex As Long) As String
    Dim result As String

    On Error Resume Next
    result = "|B" & CStr(borderIndex) & "=" & _
             CStr(formattedObject.Borders(borderIndex).LineStyle) & "," & _
             CStr(formattedObject.Borders(borderIndex).Weight) & "," & _
             CStr(formattedObject.Borders(borderIndex).Color) & "," & _
             CStr(formattedObject.Borders(borderIndex).TintAndShade)
    If Err.Number <> 0 Then
        Err.Clear
        result = "|B" & CStr(borderIndex) & "=#NA#"
    End If
    On Error GoTo 0

    BorderSignature = result
End Function

Private Function WorkbookHasConditionalFormatting(ByVal wb As Workbook, ByRef detail As String) As Boolean
    Dim ws As Worksheet
    Dim cfCells As Range

    For Each ws In wb.Worksheets
        Set cfCells = GetConditionalFormattingCells(ws)
        If Not cfCells Is Nothing Then
            detail = ws.Name & ": Conditional Formatting remains after workbook conversion."
            WorkbookHasConditionalFormatting = True
            Exit Function
        End If
    Next ws
End Function

Private Sub DeleteFailedCopy(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill filePath
    End If
    On Error GoTo 0
End Sub
