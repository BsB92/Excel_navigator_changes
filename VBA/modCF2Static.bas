Attribute VB_Name = "modCF2Static"
Option Explicit

Private Const CF2STATIC_SUFFIX As String = "_without_conditional_formatting"
Private Const XL_DATABAR_TYPE As Long = 4
Private Const XL_ICONSET_TYPE As Long = 6
Private Const SIGNATURE_ERROR_PREFIX As String = "#FORMAT_SIGNATURE_ERROR#"

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
    Dim prevAutomationSecurity As MsoAutomationSecurity
    Dim calculationChanged As Boolean
    Dim automationSecurityChanged As Boolean
    Dim appStateCaptured As Boolean
    Dim summaryText As String

    On Error GoTo FatalError

    If ownerForm Is Nothing Then Exit Sub
    If Not CBool(ownerForm.tglBatchMode.Value) Then
        MsgBox "Turn ON Selection mode and select at least one file.", vbExclamation, "CF2Static"
        Exit Sub
    End If

    Set errors = New Collection
    Set selectedWorkbooks = GetSelectedWorkbooks(ownerForm, errors)
    failCount = errors.Count

    If selectedWorkbooks.Count = 0 Then
        summaryText = "No valid selected workbooks were found."
        If errors.Count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & BuildErrorList(errors)
        MsgBox summaryText, vbExclamation, "CF2Static"
        Exit Sub
    End If

    targetFolder = PickCF2StaticTargetFolder()
    If Len(targetFolder) = 0 Then Exit Sub

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts
    prevAskToUpdateLinks = Application.AskToUpdateLinks
    prevCalculation = Application.Calculation
    prevAutomationSecurity = Application.AutomationSecurity
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    On Error Resume Next
    Err.Clear
    Application.Calculation = xlCalculationManual
    calculationChanged = (Err.Number = 0)
    Err.Clear
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    automationSecurityChanged = (Err.Number = 0)
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
        If automationSecurityChanged Then Application.AutomationSecurity = prevAutomationSecurity
    End If
    On Error GoTo 0

    summaryText = "CF2Static finished." & vbCrLf & _
                  "Processed successfully: " & CStr(okCount) & vbCrLf & _
                  "Failed: " & CStr(failCount)
    If Not errors Is Nothing Then
        If errors.Count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & BuildErrorList(errors)
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

Private Function BuildErrorList(ByVal errors As Collection) As String
    Dim result As String
    Dim i As Long

    result = "Errors / limitations:"
    For i = 1 To errors.Count
        result = result & vbCrLf & "- " & CStr(errors(i))
    Next i
    BuildErrorList = result
End Function

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
    Dim failedCopyRemoved As Boolean

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
    On Error GoTo 0

    failedCopyRemoved = DeleteFailedCopy(outPath)
    If Not failedCopyRemoved And OutputFileExists(outPath) Then
        If Len(detail) > 0 Then detail = detail & " "
        detail = detail & "Failed output copy could not be removed: " & outPath
    End If

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
    Dim capturedFormats As Object
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
    Set capturedFormats = CreateObject("Scripting.Dictionary")

    ' Capture every rendered format before changing any cell. This prevents a formula-based
    ' rule from being re-evaluated against static formats already written to an earlier cell.
    For Each c In cfCells.Cells
        beforeSignature = BuildDisplayFormatSignature(c)
        If IsSignatureError(beforeSignature) Then
            detail = "Could not read complete DisplayFormat at " & c.Address(False, False, xlA1) & "."
            Exit Function
        End If

        snapshots(c.Address(False, False, xlA1)) = beforeSignature
        capturedFormats(c.Address(False, False, xlA1)) = CaptureDisplayFormat(c)
    Next c

    ' DisplayFormat can differ in every cell (notably for formula rules and color scales),
    ' so write only the CF-covered cells and intentionally apply their formats one by one.
    For Each c In cfCells.Cells
        ApplyCapturedFormat c, capturedFormats(c.Address(False, False, xlA1))
    Next c

    ws.Cells.FormatConditions.Delete

    For Each addressKey In snapshots.Keys
        afterSignature = BuildStaticFormatSignature(ws.Range(CStr(addressKey)))
        If IsSignatureError(afterSignature) Then
            detail = "Could not validate static formatting at " & CStr(addressKey) & "."
            Exit Function
        End If
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

    On Error GoTo EH

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
    Exit Function

EH:
    details = "Could not safely inspect all Conditional Formatting rule types: " & Err.Description
    HasUnsupportedConditionalFormatting = True
End Function

Private Function CaptureDisplayFormat(ByVal sourceCell As Range) As Variant
    Dim visibleFormat As Object
    Dim borderIndexes As Variant
    Dim borderIndex As Variant
    Dim snapshot(0 To 28) As Variant
    Dim i As Long

    Set visibleFormat = sourceCell.DisplayFormat

    snapshot(0) = visibleFormat.Interior.Pattern
    snapshot(1) = visibleFormat.Interior.Color
    snapshot(2) = visibleFormat.Interior.PatternColor
    snapshot(3) = visibleFormat.Font.Name
    snapshot(4) = visibleFormat.Font.Size
    snapshot(5) = visibleFormat.Font.Bold
    snapshot(6) = visibleFormat.Font.Italic
    snapshot(7) = visibleFormat.Font.Underline
    snapshot(8) = visibleFormat.Font.Strikethrough
    snapshot(9) = visibleFormat.Font.Color
    snapshot(10) = visibleFormat.NumberFormat

    borderIndexes = GetBorderIndexes()
    i = 11
    For Each borderIndex In borderIndexes
        snapshot(i) = visibleFormat.Borders(CLng(borderIndex)).LineStyle
        snapshot(i + 1) = visibleFormat.Borders(CLng(borderIndex)).Weight
        snapshot(i + 2) = visibleFormat.Borders(CLng(borderIndex)).Color
        i = i + 3
    Next borderIndex

    CaptureDisplayFormat = snapshot
End Function

Private Sub ApplyCapturedFormat(ByVal targetCell As Range, ByVal snapshot As Variant)
    Dim borderIndexes As Variant
    Dim borderIndex As Variant
    Dim i As Long

    With targetCell.Interior
        .Pattern = snapshot(0)
        .Color = snapshot(1)
        .PatternColor = snapshot(2)
    End With

    With targetCell.Font
        .Name = snapshot(3)
        .Size = snapshot(4)
        .Bold = snapshot(5)
        .Italic = snapshot(6)
        .Underline = snapshot(7)
        .Strikethrough = snapshot(8)
        .Color = snapshot(9)
    End With

    targetCell.NumberFormat = snapshot(10)

    borderIndexes = GetBorderIndexes()
    i = 11
    For Each borderIndex In borderIndexes
        With targetCell.Borders(CLng(borderIndex))
            .LineStyle = snapshot(i)
            .Weight = snapshot(i + 1)
            .Color = snapshot(i + 2)
        End With
        i = i + 3
    Next borderIndex
End Sub

Private Function GetBorderIndexes() As Variant
    GetBorderIndexes = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                             xlDiagonalDown, xlDiagonalUp)
End Function

Private Function BuildDisplayFormatSignature(ByVal c As Range) As String
    Dim visibleFormat As Object

    On Error GoTo EH
    Set visibleFormat = c.DisplayFormat
    BuildDisplayFormatSignature = BuildFormatSignature(visibleFormat)
    Exit Function

EH:
    BuildDisplayFormatSignature = SIGNATURE_ERROR_PREFIX & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function BuildStaticFormatSignature(ByVal c As Range) As String
    On Error GoTo EH
    BuildStaticFormatSignature = BuildFormatSignature(c)
    Exit Function

EH:
    BuildStaticFormatSignature = SIGNATURE_ERROR_PREFIX & CStr(Err.Number) & ":" & Err.Description
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

    borderIndexes = GetBorderIndexes()
    For Each borderIndex In borderIndexes
        result = result & BorderSignature(formattedObject, CLng(borderIndex))
    Next borderIndex

    BuildFormatSignature = result
    Exit Function

EH:
    BuildFormatSignature = SIGNATURE_ERROR_PREFIX & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function BorderSignature(ByVal formattedObject As Object, ByVal borderIndex As Long) As String
    Dim result As String

    On Error Resume Next
    result = "|B" & CStr(borderIndex) & "=" & _
             CStr(formattedObject.Borders(borderIndex).LineStyle) & "," & _
             CStr(formattedObject.Borders(borderIndex).Weight) & "," & _
             CStr(formattedObject.Borders(borderIndex).Color)
    If Err.Number <> 0 Then
        Err.Clear
        result = "|B" & CStr(borderIndex) & "=#NA#"
    End If
    On Error GoTo 0

    BorderSignature = result
End Function

Private Function IsSignatureError(ByVal signature As String) As Boolean
    IsSignatureError = (Left$(signature, Len(SIGNATURE_ERROR_PREFIX)) = SIGNATURE_ERROR_PREFIX)
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

Private Function DeleteFailedCopy(ByVal filePath As String) As Boolean
    On Error GoTo EH

    If Len(filePath) = 0 Then
        DeleteFailedCopy = True
        Exit Function
    End If
    If Not OutputFileExists(filePath) Then
        DeleteFailedCopy = True
        Exit Function
    End If

    On Error Resume Next
    SetAttr filePath, vbNormal
    On Error GoTo EH
    Kill filePath

    DeleteFailedCopy = Not OutputFileExists(filePath)
    Exit Function

EH:
    DeleteFailedCopy = False
End Function
