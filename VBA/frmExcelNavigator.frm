VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExcelNavigator 
   Caption         =   "ExcelNavigator v5.0"
   ClientHeight    =   9795.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "frmExcelNavigator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExcelNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
' =========================================================
' BATCH SAFETY DESIGN NOTES
'
' - Excel RefreshAll cannot be safely aborted from VBA.
' - Cancel only stops processing of NEXT workbooks.
' - Current refresh must be allowed to finish.
' - No UI manipulation is allowed during RefreshAll.
'
' This design prevents:
' - Excel crashes
' - forced workbook closes
' - Document Recovery scenarios
' =========================================================

' ========= FORM SIZE LIMITS =========
Private mHooked As Boolean
Private Const FORM_MAX_W As Long = 1200
Private Const FORM_MAX_H As Long = 900
Private Const REG_APP As String = "ExcelNavigator_v5.0"
Private Const REG_SEC As String = "FormState"
Private Const REFRESH_TIMEOUT_SEC As Long = 300 ' 300=5min
Private mCancelBatch As Boolean
Private mBatchRunning As Boolean


' ========= STATE =========
Private mActivating As Boolean
Private mAllNames() As String
Private mAllFullNames() As String
Private mAllCount As Long

Private mStatus As Object   ' key=wbName, value=Array(lastRefreshDate As Date, isRefreshing As Boolean)
Private mLastReload As Date
Private mFixingHeader As Boolean
Private mUIBusy As Boolean
Private mLastCopyFolder As String
' ========= RESIZE LAYOUT CACHE =========
Private mBaseInsideW As Single, mBaseInsideH As Single
Private mBaseFileColW As Single
Private mTextMeasureLabel As Object

Private mTopBlockBottom As Single   ' bottom of search row
Private mBottomBlockTop As Single   ' top of bottom buttons row (before shifting)

Private mRightMargin As Single
Private mGap As Single

Private mCtlTop(1 To 21) As Single
Private mMinTrackW As Long
Private mMinTrackH As Long
Private mHookReady As Boolean
Private mOpenCopiedOffsetTop As Single
Private mOpenCopiedOffsetLeft As Single
Private Const HEADER_IMAGE_MARGIN As Single = 6
Private Const KEYBOARD_CTRL_MASK As Integer = 2


' ========= CONSTANTS =========
Private Const ACTIVE_PREFIX As String = "> "
Private Const PANEL_TOGGLE_COLLAPSED As String = ">>>>"
Private Const PANEL_TOGGLE_EXPANDED As String = "<<<<"
Private Const CLOSEMODE_ASK_EACH As Long = 0
Private Const CLOSEMODE_SAVE_ALL As Long = 1
Private Const CLOSEMODE_DONT_SAVE_ALL As Long = 2
Private Const MOUSE_BUTTON_RIGHT As Integer = 2
Private WithEvents mBtnTogglePanel As MSForms.CommandButton
Attribute mBtnTogglePanel.VB_VarHelpID = -1
Private WithEvents mLstSheets As MSForms.ListBox
Attribute mLstSheets.VB_VarHelpID = -1
Private mLblSheetsWorkbook As MSForms.Label
Private mIsExpandedView As Boolean
Private mCollapsedInsideW As Single
Private mPanelWidth As Single
Private mBaseListLeft As Single
Private mUpdatingSheets As Boolean
Private Const SELECTION_OFF_COLOR As Long = &H8080FF
Private mActivatingSheetFromList As Boolean
Private mActivatingWorkbookFromList As Boolean
Private mWorkbookKeyboardNavigation As Boolean
Private mSheetKeyboardNavigation As Boolean
Private WithEvents mBtnSettings As MSForms.CommandButton
Attribute mBtnSettings.VB_VarHelpID = -1
Private WithEvents mBtnHelp As MSForms.CommandButton
Attribute mBtnHelp.VB_VarHelpID = -1
Private WithEvents mBtnSnapshotCreate As MSForms.CommandButton
Attribute mBtnSnapshotCreate.VB_VarHelpID = -1
Private WithEvents mBtnCompareFiles As MSForms.CommandButton
Attribute mBtnCompareFiles.VB_VarHelpID = -1
Private mSnapshotMode As Boolean
Private mSnapshotFrame As MSForms.Frame
Private WithEvents mBtnSnapshotActionCreate As MSForms.CommandButton
Attribute mBtnSnapshotActionCreate.VB_VarHelpID = -1
Private WithEvents mBtnSnapshotActionCompare As MSForms.CommandButton
Attribute mBtnSnapshotActionCompare.VB_VarHelpID = -1
Private WithEvents mBtnSnapshotActionHistory As MSForms.CommandButton
Attribute mBtnSnapshotActionHistory.VB_VarHelpID = -1
Private Const TOP_LEFT_BUTTON_MARGIN As Single = 6
Private Const TOP_LEFT_BUTTON_GAP As Single = 6
Private mSettingsMode As Boolean
Private mSettingsFrame As MSForms.Frame
Private mSettingsLblTitle As MSForms.Label
Private mSettingsLblCopy As MSForms.Label
Private mSettingsTxtCopy As MSForms.TextBox
Private mSettingsLblOpen As MSForms.Label
Private mSettingsTxtOpen As MSForms.TextBox
Private mSettingsLblCompareMode As MSForms.Label
Private mSettingsOptCompareStrict As MSForms.OptionButton
Private mSettingsOptCompareValue As MSForms.OptionButton
Private mSettingsOptCompareHybrid As MSForms.OptionButton
Private WithEvents mBtnSettingsSave As MSForms.CommandButton
Attribute mBtnSettingsSave.VB_VarHelpID = -1
Private WithEvents mBtnSettingsCancel As MSForms.CommandButton
Attribute mBtnSettingsCancel.VB_VarHelpID = -1

Private Sub btnCancel_Click()
    ' Cancel DOES NOT stop an active RefreshAll.
    ' It only prevents starting refresh/save for NEXT workbooks.
    mCancelBatch = True
End Sub
Public Sub RestoreNavigatorToFront()
    On Error Resume Next
    If Not Me.Visible Then Me.Show vbModeless
    modWinAPI.BringFormToFront Me.Caption
    modWinAPI.SetTopMostState Me.Caption, True
    If Me.lstWorkbooks.Visible And Me.lstWorkbooks.enabled Then Me.lstWorkbooks.SetFocus
End Sub

Private Sub ActivateWorkbookThenRestoreNavigator(ByVal wb As Workbook)
    On Error Resume Next

    If wb Is Nothing Then Exit Sub
    If wb.Windows.Count > 0 Then wb.Windows(1).Activate
    wb.Activate
End Sub

Private Function GetSelectedWorkbookNames() As Collection
    Dim col As Collection
    Dim i As Long
    Dim wbName As String

    Set col = New Collection

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            wbName = GetRawNameFromRow(i)
            col.Add wbName
        End If
    Next i

    Set GetSelectedWorkbookNames = col
End Function

Private Function CountUnsavedSelectedClosableWorkbooks() As Long
    Dim i As Long
    Dim wb As Workbook
    Dim wbName As String
    Dim cnt As Long

    cnt = 0

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            wbName = GetRawNameFromRow(i)
            Set wb = GetWorkbookByName(wbName)

            If Not wb Is Nothing Then
                If Not wb.IsAddin Then
                    If UCase$(wb.Name) <> "PERSONAL.XLSB" Then
                        If Not wb.Saved Then
                            If AnyWorkbookRefreshing(wb) Then
                                If LCase$(Left$(wb.Path, 4)) = "http" Then
                                    cnt = cnt + 1
                                End If
                            Else
                                cnt = cnt + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    CountUnsavedSelectedClosableWorkbooks = cnt
End Function

Private Function ResolveCloseMode(ByVal unsavedCount As Long) As Long
    Dim resp As VbMsgBoxResult

    If unsavedCount <= 1 Then
        ResolveCloseMode = CLOSEMODE_ASK_EACH
        Exit Function
    End If

    resp = SafeMsgBox( _
        "Unsaved files detected: " & unsavedCount & vbCrLf & vbCrLf & _
        "Yes = save ALL unsaved files" & vbCrLf & _
        "No = close ALL unsaved files WITHOUT saving" & vbCrLf & _
        "Cancel = ask separately for each file", _
        vbYesNoCancel Or vbQuestion, _
        "Close mode" _
    )

    Select Case resp
        Case vbYes
            ResolveCloseMode = CLOSEMODE_SAVE_ALL
        Case vbNo
            ResolveCloseMode = CLOSEMODE_DONT_SAVE_ALL
        Case Else
            ResolveCloseMode = CLOSEMODE_ASK_EACH
    End Select
End Function


Private Sub btnCopyBreakLinks_Click()

    If mUIBusy Then Exit Sub
    If Not Me.tglBatchMode.Value Then
        SafeMsgBox "Turn ON Selection mode and select at least one file.", vbExclamation
        Exit Sub
    End If

    Dim suffix As String
    Dim targetFolder As String
    Dim i As Long, wbName As String
    Dim srcWb As Workbook
    Dim copiedCount As Long
    Dim copiedPaths As Collection
    Dim outPath As String

    suffix = Trim$(CStr(Me.txtSuffix.Value))
    If Len(suffix) = 0 Then
        SafeMsgBox "Suffix is required (e.g. _without_formulas).", vbExclamation
        Exit Sub
    End If
    ' Use suffix exactly as typed by user (no automatic prefixing)

    targetFolder = PickFolder("Choose target folder for copied files")
    If Len(targetFolder) = 0 Then Exit Sub

    copiedCount = 0
    Set copiedPaths = New Collection

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then

            wbName = GetRawNameFromRow(i)
            Set srcWb = GetWorkbookByName(wbName)

            If srcWb Is Nothing Then
                SafeMsgBox "Workbook not found: " & wbName, vbExclamation
            ElseIf Len(srcWb.Path) = 0 Then
                SafeMsgBox "Workbook must be saved first: " & srcWb.Name, vbExclamation
            Else
                If CopyAndBreakLinks(srcWb, targetFolder, suffix, outPath) Then
                    copiedCount = copiedCount + 1
                    copiedPaths.Add outPath
                End If
            End If

        End If
    Next i

    If copiedCount > 0 Then
        mLastCopyFolder = targetFolder
        Me.btnOpenCopyFolder.enabled = True
        If ShouldOpenCopiedFiles() Then
            OpenCopiedFiles copiedPaths
        End If
    End If

    SafeMsgBox "Done. Copied and processed: " & copiedCount & " file(s).", vbInformation
End Sub

Private Function PickFolder(ByVal titleText As String) As String
    Dim fd As FileDialog
    Dim initialFolder As String

    initialFolder = modNavigatorSettings.ResolveInitialFolder(ThisWorkbook.Path)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = titleText
        .AllowMultiSelect = False
        On Error Resume Next
        .InitialFileName = initialFolder
        On Error GoTo 0
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Private Function CopyAndBreakLinks(ByVal srcWb As Workbook, ByVal targetFolder As String, ByVal suffix As String, ByRef outPath As String) As Boolean
    Dim copiedWb As Workbook
    Dim linksBefore As Variant
    Dim prevAskToUpdate As Boolean
    Dim prevDisplayAlerts As Boolean

    On Error GoTo EH

    outPath = BuildOutputPath(targetFolder, srcWb.Name, suffix)

    srcWb.SaveCopyAs outPath

    prevAskToUpdate = Application.AskToUpdateLinks
    prevDisplayAlerts = Application.DisplayAlerts
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False

    Set copiedWb = Workbooks.Open(fileName:=outPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True)

    linksBefore = copiedWb.LinkSources(Type:=xlExcelLinks)
    AddExternalLinksAuditSheet copiedWb, linksBefore, srcWb.FullName
    BreakExternalLinks copiedWb, srcWb.FullName

    copiedWb.Save
    copiedWb.Close saveChanges:=False

    Application.DisplayAlerts = prevDisplayAlerts
    Application.AskToUpdateLinks = prevAskToUpdate

    CopyAndBreakLinks = True
    Exit Function

EH:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.AskToUpdateLinks = prevAskToUpdate
    SafeMsgBox "Copy/BreakLinks error for: " & srcWb.Name & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If Not copiedWb Is Nothing Then copiedWb.Close saveChanges:=False
    CopyAndBreakLinks = False
End Function

Private Sub AddExternalLinksAuditSheet(ByVal wb As Workbook, ByVal linksBefore As Variant, ByVal sourceWorkbookPath As String)
    Dim wsLog As Worksheet
    Dim ws As Worksheet
    Dim rngFormulas As Range
    Dim c As Range
    Dim r As Long
    Dim i As Long
    Dim formulaText As String
    Dim extRef As String

    On Error GoTo EH

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("_ExternalLinksLog").Delete
    Application.DisplayAlerts = True
    On Error GoTo EH

    Set wsLog = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsLog.Name = "_ExternalLinksLog"

    wsLog.Range("A1:H1").Value = Array("Timestamp", "SourceWorkbook", "Sheet", "CellAddress", "Formula", "ValueAtSnapshot", "ExternalReferenceDetected", "LinkSource")
    r = 2

    If IsArray(linksBefore) Then
        For i = LBound(linksBefore) To UBound(linksBefore)
            wsLog.Cells(r, 1).Value = Now
            wsLog.Cells(r, 2).Value = sourceWorkbookPath
            wsLog.Cells(r, 3).Value = "<workbook-link>"
            wsLog.Cells(r, 4).Value = ""
            wsLog.Cells(r, 5).Value = ""
            wsLog.Cells(r, 6).Value = ""
            wsLog.Cells(r, 7).Value = ""
            wsLog.Cells(r, 8).Value = CStr(linksBefore(i))
            r = r + 1
        Next i
    End If

    For Each ws In wb.Worksheets
        If ws.Name <> wsLog.Name Then
            Set rngFormulas = Nothing
            On Error Resume Next
            Set rngFormulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo EH

            If Not rngFormulas Is Nothing Then
                For Each c In rngFormulas.Cells
                    formulaText = CStr(c.Formula)
                    If InStr(1, formulaText, "[", vbTextCompare) > 0 Then
                        extRef = ExtractBracketReference(formulaText)
                        wsLog.Cells(r, 1).Value = Now
                        wsLog.Cells(r, 2).Value = sourceWorkbookPath
                        wsLog.Cells(r, 3).Value = ws.Name
                        wsLog.Cells(r, 4).Value = c.Address(False, False)
                        wsLog.Cells(r, 5).Value = formulaText
                        wsLog.Cells(r, 6).Value = c.Value2
                        wsLog.Cells(r, 7).Value = extRef
                        wsLog.Cells(r, 8).Value = ""
                        r = r + 1
                    End If
                Next c
            End If
        End If
    Next ws

    wsLog.Columns("A:H").EntireColumn.AutoFit
    Exit Sub

EH:
    On Error Resume Next
    Application.DisplayAlerts = True
End Sub

Private Function ExtractBracketReference(ByVal formulaText As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, formulaText, "[", vbTextCompare)
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, formulaText, "]", vbTextCompare)
    If p2 = 0 Then
        ExtractBracketReference = Mid$(formulaText, p1)
    Else
        ExtractBracketReference = Mid$(formulaText, p1, p2 - p1 + 1)
    End If
End Function

Private Function BuildOutputPath(ByVal folderPath As String, ByVal fileName As String, ByVal suffix As String) As String
    Dim baseName As String, ext As String
    Dim dotPos As Long
    Dim normalizedFolder As String

    normalizedFolder = Trim$(folderPath)
    Do While Right$(normalizedFolder, 1) = Application.PathSeparator
        normalizedFolder = Left$(normalizedFolder, Len(normalizedFolder) - 1)
    Loop

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        ext = ".xlsx"
    End If

    BuildOutputPath = normalizedFolder & Application.PathSeparator & baseName & suffix & ext
End Function

Private Sub ForceBreakExternalFormulas(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim rngFormulas As Range
    Dim c As Range
    Dim chObj As ChartObject
    Dim srs As Series

    On Error Resume Next

    For Each ws In wb.Worksheets
        Set rngFormulas = Nothing
        Set rngFormulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Not rngFormulas Is Nothing Then
            For Each c In rngFormulas.Cells
                If InStr(1, c.Formula, "[", vbTextCompare) > 0 Then
                    c.Value = c.Value
                End If
            Next c
        End If

        For Each chObj In ws.ChartObjects
            For Each srs In chObj.Chart.SeriesCollection
                If InStr(1, srs.Formula, "[", vbTextCompare) > 0 Then
                    srs.Values = srs.Values
                    srs.XValues = srs.XValues
                End If
            Next srs
        Next chObj
    Next ws

    On Error GoTo 0
End Sub

Private Sub BreakExternalLinks(ByVal wb As Workbook, ByVal sourceWorkbookPath As String)
    Dim links As Variant
    Dim i As Long
    Dim passNo As Long
    Dim linkName As String
    Dim prevAskToUpdate As Boolean
    Dim prevDisplayAlerts As Boolean

    prevAskToUpdate = Application.AskToUpdateLinks
    prevDisplayAlerts = Application.DisplayAlerts
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False

    On Error GoTo CleanExit

    For passNo = 1 To 3
        links = wb.LinkSources(Type:=xlExcelLinks)
        If IsEmpty(links) Then Exit For
        If Not IsArray(links) Then Exit For

        For i = LBound(links) To UBound(links)
            linkName = CStr(links(i))
            On Error Resume Next
            wb.BreakLink Name:=linkName, Type:=xlLinkTypeExcelLinks
            If Err.Number <> 0 Then
                Err.Clear
                wb.ChangeLink Name:=linkName, NewName:=sourceWorkbookPath, Type:=xlLinkTypeExcelLinks
                wb.BreakLink Name:=sourceWorkbookPath, Type:=xlLinkTypeExcelLinks
            End If
            Err.Clear
            On Error GoTo CleanExit
        Next i
    Next passNo

    links = wb.LinkSources(Type:=xlExcelLinks)
    If Not IsEmpty(links) Then
        ForceBreakExternalFormulas wb
    End If

CleanExit:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.AskToUpdateLinks = prevAskToUpdate
End Sub

Private Sub btnCopyWithSuffix_Click()
    If mUIBusy Then Exit Sub
    If Not Me.tglBatchMode.Value Then
        SafeMsgBox "Turn ON Selection mode and select at least one file.", vbExclamation
        Exit Sub
    End If

    Dim suffix As String
    Dim targetFolder As String
    Dim i As Long, wbName As String
    Dim srcWb As Workbook
    Dim copiedCount As Long
    Dim copiedPaths As Collection
    Dim outPath As String

    suffix = Trim$(CStr(Me.txtSuffix.Value))
    If Len(suffix) = 0 Then
        SafeMsgBox "Suffix is required (e.g. _without_formulas).", vbExclamation
        Exit Sub
    End If
    If Left$(suffix, 1) <> "_" Then
        suffix = "_" & suffix
        Me.txtSuffix.Value = suffix
    End If

    targetFolder = PickFolder("Choose target folder for copied files")
    If Len(targetFolder) = 0 Then Exit Sub

    copiedCount = 0
    Set copiedPaths = New Collection

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            wbName = GetRawNameFromRow(i)
            Set srcWb = GetWorkbookByName(wbName)

            If srcWb Is Nothing Then
                SafeMsgBox "Workbook not found: " & wbName, vbExclamation
            ElseIf Len(srcWb.Path) = 0 Then
                SafeMsgBox "Workbook must be saved first: " & srcWb.Name, vbExclamation
            Else
                If CopyWithSuffixOnly(srcWb, targetFolder, suffix, outPath) Then
                    copiedCount = copiedCount + 1
                    copiedPaths.Add outPath
                End If
            End If
        End If
    Next i

    If copiedCount > 0 Then
        mLastCopyFolder = targetFolder
        Me.btnOpenCopyFolder.enabled = True
        If ShouldOpenCopiedFiles() Then
            OpenCopiedFiles copiedPaths
        End If
    End If

    SafeMsgBox "Done. Copied: " & copiedCount & " file(s).", vbInformation
End Sub

Private Sub btnOpenFile_Click()
    OpenFilesAndOptionalRefresh False
End Sub

Private Sub btnOpenAndRefresh_Click()
    OpenFilesAndOptionalRefresh True
End Sub

Private Sub OpenFilesAndOptionalRefresh(ByVal doRefresh As Boolean)
    Dim files As Collection
    Dim fp As Variant
    Dim wb As Workbook
    Dim openedCount As Long
    Dim refOK As Long, refTO As Long

    If mUIBusy Then Exit Sub
    If mBatchRunning Then
        SafeMsgBox "Batch is running. Finish/Cancel batch first.", vbExclamation
        Exit Sub
    End If

    Set files = PickFilesMulti(IIf(doRefresh, "Open & Refresh - select Excel files", "Open file(s) - select Excel files"))
    If files Is Nothing Then Exit Sub
    If files.Count = 0 Then Exit Sub

    mBatchRunning = True
    SetBatchUI True
    mCancelBatch = False

    openedCount = 0: refOK = 0: refTO = 0

    For Each fp In files
        If mCancelBatch Then Exit For

        Set wb = OpenWorkbookSafe(CStr(fp))
        If Not wb Is Nothing Then
            openedCount = openedCount + 1

            If doRefresh Then
                If RefreshOneWorkbook(wb) Then
                    refOK = refOK + 1
                ElseIf IsTimedOut(wb.Name) Then
                    refTO = refTO + 1
                End If

                If mCancelBatch Then
                    btnCancel.enabled = False
                    Exit For
                End If
            End If
        End If
    Next fp

    ReloadListPreserveSelection
    RefreshVisuals

    If doRefresh Then
        SafeMsgBox "Done. Opened: " & openedCount & ", refreshed: " & refOK & ", timed out: " & refTO & ".", vbInformation
    Else
        SafeMsgBox "Done. Opened: " & openedCount & ".", vbInformation
    End If

FINALLY:
    mCancelBatch = False
    mBatchRunning = False
    Application.StatusBar = False
    SetBatchUI False
    ForceTopMost
End Sub


Private Sub btnOpenCopyFolder_Click()

    If Len(mLastCopyFolder) = 0 Then
        SafeMsgBox "No folder to open yet. Use Copy + Break Links first.", vbExclamation
        Exit Sub
    End If

    If Dir(mLastCopyFolder, vbDirectory) = "" Then
        SafeMsgBox "Folder does not exist:" & vbCrLf & mLastCopyFolder, vbExclamation
        Exit Sub
    End If

    Shell "explorer.exe " & """" & mLastCopyFolder & """", vbNormalFocus
End Sub

' =========================================================
' CLOSE SELECTED
' - closes ONLY selected workbooks
' - NEVER closes while refreshing
' - prompts Yes/No/Cancel for unsaved changes
' - NEVER closes add-ins and PERSONAL.XLSB
' =========================================================
Private Sub btnCloseSelected_Click()
    CloseSelectedWorkbooks
End Sub

Private Sub CloseSelectedWorkbooks()
    Dim i As Long
    Dim wb As Workbook
    Dim wbName As String
    Dim cntSel As Long
    Dim cntClosed As Long
    Dim cntSkippedRefreshing As Long
    Dim cntSkippedProtected As Long
    Dim resp As VbMsgBoxResult
    Dim okContinue As Boolean
    Dim closeMode As Long
    Dim unsavedCount As Long

    On Error GoTo EH

    If mUIBusy Then Exit Sub
    If mBatchRunning Then
        SafeMsgBox "Batch is running. Finish/Cancel batch first.", vbExclamation
        Exit Sub
    End If

    If Not Me.tglBatchMode.Value Then
        SafeMsgBox "Turn ON Selection mode and select at least one file.", vbExclamation
        Exit Sub
    End If

    resp = SafeMsgBox( _
        "Close selected workbooks?", _
        vbOKCancel Or vbQuestion, _
        "Close selected" _
    )
    If resp <> vbOK Then Exit Sub

    unsavedCount = CountUnsavedSelectedClosableWorkbooks()
    closeMode = ResolveCloseMode(unsavedCount)

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            cntSel = cntSel + 1
            wbName = GetRawNameFromRow(i)
            Set wb = GetWorkbookByName(wbName)

            If wb Is Nothing Then GoTo NEXT_I

            If wb.IsAddin Then
                cntSkippedProtected = cntSkippedProtected + 1
                GoTo NEXT_I
            End If

            If UCase$(wb.Name) = "PERSONAL.XLSB" Then
                cntSkippedProtected = cntSkippedProtected + 1
                GoTo NEXT_I
            End If

            If AnyWorkbookRefreshing(wb) Then
                If LCase$(Left$(wb.Path, 4)) <> "http" Then
                    cntSkippedRefreshing = cntSkippedRefreshing + 1
                    GoTo NEXT_I
                End If
            End If

            okContinue = CloseOneWorkbookWithPrompt(wb, closeMode)
            If Not okContinue Then Exit For

            If Not WorkbookIsOpen(wbName) Then
                cntClosed = cntClosed + 1

                On Error Resume Next
                If mStatus.Exists(wbName) Then mStatus.Remove wbName
                On Error GoTo 0
            End If

NEXT_I:
        End If
    Next i

    ReloadListPreserveSelection
    RefreshVisuals

    If cntSel = 0 Then
        SafeMsgBox "Nothing selected.", vbExclamation
    Else
        SafeMsgBox "Done. Closed: " & cntClosed & _
                   ", skipped (refreshing): " & cntSkippedRefreshing & _
                   ", skipped (protected): " & cntSkippedProtected & ".", vbInformation
    End If

    Exit Sub

EH:
    SafeMsgBox "CloseSelected error: " & Err.Description, vbCritical
End Sub

Private Function CloseOneWorkbookWithPrompt(ByVal wb As Workbook, ByVal closeMode As Long) As Boolean
    Dim resp As VbMsgBoxResult
    Dim saveChanges As Boolean

    On Error GoTo EH

    If wb.Saved Then
        On Error Resume Next
        modWinAPI.SetTopMostState Me.Caption, False
        On Error GoTo 0

        wb.Close saveChanges:=False

        On Error Resume Next
        modWinAPI.SetTopMostState Me.Caption, True
        On Error GoTo 0

        CloseOneWorkbookWithPrompt = True
        Exit Function
    End If

    Select Case closeMode
        Case CLOSEMODE_SAVE_ALL
            saveChanges = True

        Case CLOSEMODE_DONT_SAVE_ALL
            saveChanges = False

        Case Else
            resp = SafeMsgBox( _
                "Save changes before closing?" & vbCrLf & wb.Name, _
                vbYesNoCancel Or vbQuestion, _
                "Close workbook" _
            )

            If resp = vbCancel Then
                CloseOneWorkbookWithPrompt = False
                Exit Function
            End If

            saveChanges = (resp = vbYes)
    End Select

    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, False
    On Error GoTo 0

    wb.Close saveChanges:=saveChanges

    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, True
    On Error GoTo 0

    CloseOneWorkbookWithPrompt = True
    Exit Function

EH:
    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, True
    On Error GoTo 0

    SafeMsgBox "Close error in: " & wb.Name & vbCrLf & Err.Description, vbCritical
    CloseOneWorkbookWithPrompt = True
End Function
Private Function WorkbookIsOpen(ByVal wbName As String) As Boolean
    Dim w As Workbook
    On Error Resume Next
    Set w = Application.Workbooks(wbName)
    WorkbookIsOpen = Not (w Is Nothing)
    On Error GoTo 0
End Function

Private Sub CheckBox1_Click()

End Sub

Private Function ShouldOpenCopiedFiles() As Boolean
    Dim ctl As Object

    Set ctl = GetControlIfExists("ChckBox1")
    If ctl Is Nothing Then Set ctl = GetControlIfExists("CheckBox1")
    If ctl Is Nothing Then Exit Function

    On Error Resume Next
    ShouldOpenCopiedFiles = CBool(ctl.Value)
    On Error GoTo 0
End Function

Private Sub OpenCopiedFiles(ByVal copiedPaths As Collection)
    Dim fp As Variant
    Dim wb As Workbook
    Dim openedCount As Long
    Dim failedCount As Long

    If copiedPaths Is Nothing Then Exit Sub
    If copiedPaths.Count = 0 Then Exit Sub

    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, False
    On Error GoTo 0

    For Each fp In copiedPaths
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(fileName:=CStr(fp), UpdateLinks:=0, ReadOnly:=False)
        If wb Is Nothing Then
            failedCount = failedCount + 1
        Else
            openedCount = openedCount + 1
        End If
        On Error GoTo 0
    Next fp

    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, True
    On Error GoTo 0

    If failedCount > 0 Then
        SafeMsgBox "Opened " & openedCount & " copied file(s)." & vbCrLf & _
                   "Could not open " & failedCount & " file(s).", vbExclamation
    End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtFullPath_Enter()
    SelectAllFullPath
End Sub

Private Sub txtFullPath_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SelectAllFullPath
End Sub

Private Sub SelectAllFullPath()
    On Error Resume Next
    Me.txtFullPath.SelStart = 0
    Me.txtFullPath.SelLength = Len(Me.txtFullPath.Value)
    On Error GoTo 0
End Sub

Private Function TryHookResize() As Boolean
    Dim h As LongPtr

    On Error GoTo EH

    h = modWinAPI.FindWindow("ThunderDFrame", Me.Caption)
    If h = 0 Then h = modWinAPI.FindWindow("ThunderXFrame", Me.Caption)
    If h = 0 Then
        TryHookResize = False
        Exit Function
    End If

    modFormSizeHook.EnableFormResize Me.Caption

    Me.Repaint
    DoEvents

    TryHookResize = True
    Exit Function

EH:
    TryHookResize = False
End Function


Private Sub UserForm_Activate()
    If mActivating Then Exit Sub
    mActivating = True

    On Error GoTo FINALLY

    If TryHookResize() Then mHookReady = True

    ApplyLayout
    PositionTopButtons
    Me.Repaint

FINALLY:
    mActivating = False
End Sub


' =========================================================
' INIT
' =========================================================
Private Sub UserForm_Initialize()

    Set mStatus = CreateObject("Scripting.Dictionary")

    ' ListBox:
    ' col0 = File (display)
    ' col1 = Dir
    ' col2 = Status
    ' col3 = FullPath (hidden, for search)
    Me.lstWorkbooks.ColumnCount = 4
    Me.lstWorkbooks.ColumnWidths = "240.5 pt;30 pt;55 pt;0 pt"
    Me.lstWorkbooks.MultiSelect = fmMultiSelectSingle

    Me.txtSearch.Value = ""

    Me.tglBatchMode.Value = False
    Me.tglBatchMode.Caption = "Selection mode: OFF"
    Me.tglBatchMode.BackColor = SELECTION_OFF_COLOR

    SetActionButtonsEnabled False
    SetOptionalControlEnabled "btnMaximize", True
    SetOptionalControlEnabled "btnScreen1", True
    SetOptionalControlEnabled "btnScreen2", True
    SetOptionalControlEnabled "btnScreen3", True
    mLastCopyFolder = ""
    Me.btnOpenCopyFolder.enabled = False
    Me.txtSuffix.Value = "_without_formulas"
    PinImage1AndTopRow


    ReloadListPreserveSelection
    RefreshVisuals

    
  
Dim w As Variant
Dim defaultW As Single

w = GetSetting(REG_APP, REG_SEC, "W", 0)
defaultW = Me.Width

If w > 0 Then
    Me.Width = CSng(w)
    If Me.Width < defaultW Then Me.Width = defaultW
Else
    Me.Width = defaultW
End If
If Me.Width > FORM_MAX_W Then Me.Width = FORM_MAX_W

If mMinTrackW = 0 Then mMinTrackW = CLng(Me.Width)
If mMinTrackH = 0 Then mMinTrackH = CLng(Me.Height)
btnCancel.Visible = False
btnCancel.enabled = True
EnsureRightPanelControls
mIsExpandedView = False
    mCollapsedInsideW = Me.InsideWidth
mPanelWidth = GetSheetPanelWidthPt()
SetExpandedView False
EnsureTopLeftButtons
EnsureSnapshotOverlay
ApplySnapshotOverlayVisibility
PositionTopButtons

' --- layout only (resize hook will be done in Activate when hwnd exists) ---
mHookReady = False
CacheLayout
ApplyLayout



End Sub

Private Sub EnsureTopLeftButtons()
    If mBtnSettings Is Nothing Then
        Set mBtnSettings = GetControlIfExists("btnSettings")
        If mBtnSettings Is Nothing Then
            Set mBtnSettings = Me.Controls.Add("Forms.CommandButton.1", "btnSettings", True)
        End If
    End If

    If mBtnHelp Is Nothing Then
        Set mBtnHelp = GetControlIfExists("btnHelp")
        If mBtnHelp Is Nothing Then
            Set mBtnHelp = Me.Controls.Add("Forms.CommandButton.1", "btnHelp", True)
        End If
    End If

    If mBtnSnapshotCreate Is Nothing Then
        Set mBtnSnapshotCreate = GetControlIfExists("btnSnapshotCreate")
        If mBtnSnapshotCreate Is Nothing Then Set mBtnSnapshotCreate = Me.Controls.Add("Forms.CommandButton.1", "btnSnapshotCreate", True)
    End If

    If mBtnCompareFiles Is Nothing Then
        Set mBtnCompareFiles = GetControlIfExists("btnCompareFiles")
        If mBtnCompareFiles Is Nothing Then Set mBtnCompareFiles = Me.Controls.Add("Forms.CommandButton.1", "btnCompareFiles", True)
    End If

    With mBtnSettings
        .Caption = "Settings"
        .Top = TOP_LEFT_BUTTON_MARGIN
        .Left = TOP_LEFT_BUTTON_MARGIN
        .Height = 18
        .Width = 54
        .Visible = True
    End With

    With mBtnHelp
        .Caption = "?"
        .Top = mBtnSettings.Top
        .Height = mBtnSettings.Height
        .Width = 18
        .Left = mBtnSettings.Left + mBtnSettings.Width + TOP_LEFT_BUTTON_GAP
        .Visible = True
    End With

    With mBtnSnapshotCreate
        .Caption = "Snapshot"
        .Top = mBtnSettings.Top
        .Height = mBtnSettings.Height
        .Width = 66
        .Left = mBtnHelp.Left + mBtnHelp.Width + TOP_LEFT_BUTTON_GAP
        .Visible = True
    End With

    With mBtnCompareFiles
        .Caption = "Compare Files"
        .Top = mBtnSettings.Top
        .Height = mBtnSettings.Height
        .Width = 78
        .Left = mBtnSnapshotCreate.Left + mBtnSnapshotCreate.Width + TOP_LEFT_BUTTON_GAP
        .Visible = True
    End With

End Sub

Private Sub EnsureSettingsOverlay()
    If mSettingsFrame Is Nothing Then
        Set mSettingsFrame = Me.Controls.Add("Forms.Frame.1", "fraNavigatorSettings", True)
        Set mSettingsLblTitle = mSettingsFrame.Controls.Add("Forms.Label.1", "lblSettingsTitle", True)
        Set mSettingsLblCopy = mSettingsFrame.Controls.Add("Forms.Label.1", "lblSettingsCopy", True)
        Set mSettingsTxtCopy = mSettingsFrame.Controls.Add("Forms.TextBox.1", "txtSettingsCopy", True)
        Set mSettingsLblOpen = mSettingsFrame.Controls.Add("Forms.Label.1", "lblSettingsOpen", True)
        Set mSettingsTxtOpen = mSettingsFrame.Controls.Add("Forms.TextBox.1", "txtSettingsOpen", True)
        Set mSettingsLblCompareMode = mSettingsFrame.Controls.Add("Forms.Label.1", "lblSettingsCompareMode", True)
        Set mSettingsOptCompareStrict = mSettingsFrame.Controls.Add("Forms.OptionButton.1", "optCompareStrict", True)
        Set mSettingsOptCompareValue = mSettingsFrame.Controls.Add("Forms.OptionButton.1", "optCompareValue", True)
        Set mSettingsOptCompareHybrid = mSettingsFrame.Controls.Add("Forms.OptionButton.1", "optCompareHybrid", True)
        Set mBtnSettingsSave = mSettingsFrame.Controls.Add("Forms.CommandButton.1", "btnSettingsSave", True)
        Set mBtnSettingsCancel = mSettingsFrame.Controls.Add("Forms.CommandButton.1", "btnSettingsCancel", True)
    End If

    With mSettingsFrame
        .Caption = ""
        .Left = 4
        .Top = Me.Image1.Top + Me.Image1.Height + 2
        .Width = Me.InsideWidth - 8
        .Height = Me.btnClose.Top - .Top - 6
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With mSettingsLblTitle
        .Caption = "Settings:"
        .Left = 8
        .Top = 8
        .Width = 220
        .Height = 20
        .Font.Bold = True
        .Font.Size = 11
    End With

    With mSettingsLblCopy
        .Caption = "Set default path for copy operations:"
        .Left = 8
        .Top = 40
        .Width = 260
        .Height = 16
    End With

    With mSettingsTxtCopy
        .Left = 8
        .Top = 58
        .Width = mSettingsFrame.Width - 16
        .Height = 22
    End With

    With mSettingsLblOpen
        .Caption = "Set default path for open operations:"
        .Left = 8
        .Top = 90
        .Width = 260
        .Height = 16
    End With

    With mSettingsTxtOpen
        .Left = 8
        .Top = 108
        .Width = mSettingsFrame.Width - 16
        .Height = 22
    End With

    With mSettingsLblCompareMode
        .Caption = "Snapshot compare mode:"
        .Left = 8
        .Top = 138
        .Width = 220
        .Height = 16
    End With

    With mSettingsOptCompareStrict
        .Caption = "Strict (Value + Formula)"
        .Left = 12
        .Top = 154
        .Width = 210
        .Height = 16
    End With

    With mSettingsOptCompareValue
        .Caption = "Value-only (recommended)"
        .Left = 12
        .Top = 170
        .Width = 210
        .Height = 16
    End With

    With mSettingsOptCompareHybrid
        .Caption = "Hybrid (ignore external-link formula noise)"
        .Left = 12
        .Top = 186
        .Width = mSettingsFrame.Width - 16
        .Height = 16
    End With

    With mBtnSettingsSave
        .Caption = "Save settings"
        .Width = 72
        .Height = 24
        .Top = mSettingsFrame.Height - 32
        .Left = (mSettingsFrame.Width / 2) - .Width - 10
    End With

    With mBtnSettingsCancel
        .Caption = "Cancel"
        .Width = 72
        .Height = 24
        .Top = mSettingsFrame.Height - 32
        .Left = (mSettingsFrame.Width / 2) + 10
    End With
End Sub

Private Sub ApplySettingsOverlayVisibility()
    If mSettingsFrame Is Nothing Then Exit Sub

    mSettingsFrame.Visible = mSettingsMode

    If mSettingsMode Then
        mBtnSettings.BackColor = RGB(0, 176, 80)
        mBtnSettings.Font.Bold = True
        mSettingsTxtCopy.Text = modNavigatorSettings.GetDefaultWorkingFolder()
        mSettingsTxtOpen.Text = modNavigatorSettings.GetOpenFilesFolder()
        Select Case modNavigatorSettings.GetSnapshotCompareMode()
            Case modNavigatorSettings.SNAP_COMPARE_MODE_STRICT: mSettingsOptCompareStrict.Value = True
            Case modNavigatorSettings.SNAP_COMPARE_MODE_HYBRID: mSettingsOptCompareHybrid.Value = True
            Case Else: mSettingsOptCompareValue.Value = True
        End Select
        mSettingsFrame.ZOrder 0
    Else
        mBtnSettings.BackColor = vbButtonFace
        mBtnSettings.Font.Bold = False
    End If
End Sub

Private Sub mBtnSettings_Click()
    mSettingsMode = Not mSettingsMode
    EnsureSettingsOverlay
    ApplySettingsOverlayVisibility
End Sub

Private Sub mBtnSettingsSave_Click()
    Dim copyFolder As String
    Dim openFolder As String

    copyFolder = Trim$(mSettingsTxtCopy.Text)
    openFolder = Trim$(mSettingsTxtOpen.Text)

    If Len(copyFolder) = 0 Or Len(Dir$(copyFolder, vbDirectory)) = 0 Then
        SafeMsgBox "Copy folder does not exist: " & copyFolder, vbExclamation
        Exit Sub
    End If

    If Len(openFolder) = 0 Or Len(Dir$(openFolder, vbDirectory)) = 0 Then
        SafeMsgBox "Open folder does not exist: " & openFolder, vbExclamation
        Exit Sub
    End If

    modNavigatorSettings.SaveDefaultWorkingFolder copyFolder
    modNavigatorSettings.SaveOpenFilesFolder openFolder

    If mSettingsOptCompareStrict.Value Then
        modNavigatorSettings.SaveSnapshotCompareMode modNavigatorSettings.SNAP_COMPARE_MODE_STRICT
    ElseIf mSettingsOptCompareHybrid.Value Then
        modNavigatorSettings.SaveSnapshotCompareMode modNavigatorSettings.SNAP_COMPARE_MODE_HYBRID
    Else
        modNavigatorSettings.SaveSnapshotCompareMode modNavigatorSettings.SNAP_COMPARE_MODE_VALUE_ONLY
    End If

    mSettingsMode = False
    ApplySettingsOverlayVisibility
    SafeMsgBox "Settings saved.", vbInformation
End Sub

Private Sub mBtnSettingsCancel_Click()
    mSettingsMode = False
    ApplySettingsOverlayVisibility
End Sub

Private Sub mBtnHelp_Click()
    On Error GoTo EH
    modNavigatorSettings.OpenHelpInstructions
    Exit Sub
EH:
    SafeMsgBox "Could not open help file: " & Err.Description, vbExclamation
End Sub


Private Sub mBtnCompareFiles_Click()
    On Error GoTo EH
    modSnapshotMain.SnapshotCompareAnyTwoFiles
    Exit Sub
EH:
    SafeMsgBox "Compare Files failed: " & Err.Description, vbExclamation
End Sub

Private Sub mBtnSnapshotCreate_Click()
    mSnapshotMode = Not mSnapshotMode
    EnsureSnapshotOverlay
    ApplySnapshotOverlayVisibility
End Sub



Private Sub EnsureSnapshotOverlay()
    If mSnapshotFrame Is Nothing Then
        Set mSnapshotFrame = Me.Controls.Add("Forms.Frame.1", "fraSnapshotActions", True)
        Set mBtnSnapshotActionCreate = mSnapshotFrame.Controls.Add("Forms.CommandButton.1", "btnSnapshotActionCreate", True)
        Set mBtnSnapshotActionCompare = mSnapshotFrame.Controls.Add("Forms.CommandButton.1", "btnSnapshotActionCompare", True)
        Set mBtnSnapshotActionHistory = mSnapshotFrame.Controls.Add("Forms.CommandButton.1", "btnSnapshotActionHistory", True)
    End If

    With mSnapshotFrame
        .Caption = "Snapshot / Compare"
        .Left = mBtnSnapshotCreate.Left
        .Top = mBtnSnapshotCreate.Top + mBtnSnapshotCreate.Height + 4
        .Width = 170
        .Height = 88
        .SpecialEffect = fmSpecialEffectSunken
        .BorderStyle = fmBorderStyleSingle
    End With

    With mBtnSnapshotActionCreate
        .Caption = "Create Snapshot"
        .Left = 8: .Top = 14: .Width = 150: .Height = 20
    End With
    With mBtnSnapshotActionCompare
        .Caption = "Compare Snapshot"
        .Left = 8: .Top = 36: .Width = 150: .Height = 20
    End With
    With mBtnSnapshotActionHistory
        .Caption = "Compare 2 Snapshots (A/B)"
        .Left = 8: .Top = 58: .Width = 150: .Height = 20
    End With
End Sub

Private Sub ApplySnapshotOverlayVisibility()
    If mSnapshotFrame Is Nothing Then Exit Sub
    mSnapshotFrame.Visible = mSnapshotMode
    If mSnapshotMode Then mSnapshotFrame.ZOrder 0
End Sub

Private Sub mBtnSnapshotActionCreate_Click()
    On Error GoTo EH
    modSnapshotMain.SnapshotCreateActiveWorkbook
    Exit Sub
EH:
    SafeMsgBox "Create Snapshot failed: " & Err.Description, vbExclamation
End Sub

Private Sub mBtnSnapshotActionCompare_Click()
    On Error GoTo EH
    modSnapshotMain.SnapshotCompareActiveWorkbook
    Exit Sub
EH:
    SafeMsgBox "Compare Snapshot failed: " & Err.Description, vbExclamation
End Sub

Private Sub mBtnSnapshotActionHistory_Click()
    On Error GoTo EH
    modSnapshotHistory.SnapshotHistoryManager
    Exit Sub
EH:
    SafeMsgBox "Snapshot History failed: " & Err.Description, vbExclamation
End Sub

Private Sub PinImage1AndTopRow()
    Dim img As Object
    Dim minTop As Single

    Set img = GetControlIfExists("Image1")
    If img Is Nothing Then Exit Sub

    img.TOP = HEADER_IMAGE_MARGIN
    img.Left = Me.InsideWidth - HEADER_IMAGE_MARGIN - img.Width

    minTop = img.TOP + img.Height + HEADER_IMAGE_MARGIN

    Me.txtSearch.TOP = minTop
    Me.btnReload.TOP = minTop
    Me.Label1.TOP = minTop
End Sub

' =========================================================
' SELECTION MODE
' =========================================================
Private Sub tglBatchMode_Click()
    ApplySelectionModeState
End Sub

Private Sub ApplySelectionModeState()
    On Error GoTo ModeSwitchError

    mUIBusy = True

    ' Najpierw zdejmij bieżące zaznaczenie i fokus, potem zmień MultiSelect.
    ' W niektórych konfiguracjach Excela taka kolejność zapobiega runtime error.
    Me.lstWorkbooks.ListIndex = -1
    ClearAllSelections

    If Me.tglBatchMode.Value Then
        Me.tglBatchMode.Caption = "Selection mode: ON"
        Me.lstWorkbooks.MultiSelect = fmMultiSelectMulti
        SetActionButtonsEnabled True
    Else
        Me.tglBatchMode.Caption = "Selection mode: OFF"
        Me.lstWorkbooks.MultiSelect = fmMultiSelectSingle
        SetActionButtonsEnabled False
    End If

    ShowFullPathForIndex -1
    Me.tglBatchMode.BackColor = IIf(Me.tglBatchMode.Value, RGB(0, 176, 80), SELECTION_OFF_COLOR)
    RefreshVisuals
    RefreshSheetList

SafeExit:
    mUIBusy = False
    Exit Sub

ModeSwitchError:
    mUIBusy = False
    SafeMsgBox "Cannot switch Selection mode: " & Err.Description, vbExclamation
End Sub

Private Sub SetActionButtonsEnabled(ByVal enabled As Boolean)
    Me.btnRefresh.enabled = enabled
    Me.btnSave.enabled = enabled
    Me.btnRefreshSave.enabled = enabled
    Me.btnCopyBreakLinks.enabled = enabled
    Me.btnSelectAll.enabled = enabled
    Me.btnClearAll.enabled = enabled
    Me.btnCloseSelected.enabled = enabled
    Me.btnCopyWithSuffix.enabled = enabled
    SetOptionalControlEnabled "btnMaximize", True
    SetOptionalControlEnabled "btnScreen1", True
    SetOptionalControlEnabled "btnScreen2", True
    SetOptionalControlEnabled "btnScreen3", True


End Sub

' =========================================================
' LIST EVENTS (header block)
' =========================================================
Private Sub lstWorkbooks_Click()

    Dim idx As Long

    If mUIBusy Then Exit Sub
    On Error Resume Next
    Me.lstWorkbooks.SetFocus
    On Error GoTo 0

    idx = Me.lstWorkbooks.ListIndex

    ' Klik w naglówek albo puste miejsce = nic nie rób
    If idx <= 0 Then
        mUIBusy = True
        On Error Resume Next
        Me.lstWorkbooks.Selected(0) = False
        Me.lstWorkbooks.ListIndex = -1
        Me.txtFullPath.Value = ""
        On Error GoTo 0
        mUIBusy = False
        Exit Sub
    End If

    ShowFullPathForIndex idx

    ' Click only selects row and refreshes dependent UI.
    ' Workbook activation is handled by arrow-key navigation
    ' (and double-click) to keep keyboard focus stable.
    If Me.tglBatchMode.Value Then
        RefreshVisuals
        RefreshSheetList
        Exit Sub
    End If

    RefreshVisuals
    RefreshSheetList
End Sub



Private Sub lstWorkbooks_Change()

    Dim idx As Long

    If mUIBusy Then Exit Sub

    idx = Me.lstWorkbooks.ListIndex

    If idx <= 0 Then
        mUIBusy = True
        On Error Resume Next
        Me.lstWorkbooks.Selected(0) = False
        Me.lstWorkbooks.ListIndex = -1
        Me.txtFullPath.Value = ""
        On Error GoTo 0
        mUIBusy = False
        Exit Sub
    End If

    ShowFullPathForIndex idx
    RefreshSheetList

End Sub


Private Function HandleGlobalKeyboardShortcuts(ByRef KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) As Boolean
    Dim handled As Boolean

    handled = False

    If mSettingsMode Then
        HandleGlobalKeyboardShortcuts = False
        Exit Function
    End If

    If KeyCode = vbKeyA Then
        If (Shift And KEYBOARD_CTRL_MASK) <> 0 Then
            btnSelectAll_Click
            KeyCode = 0
            handled = True
        End If
    ElseIf KeyCode = vbKeyS Then
        If Shift = 0 Then
            If Not Me.tglBatchMode.Value Then Me.tglBatchMode.Value = True
            KeyCode = 0
            handled = True
        End If
    ElseIf KeyCode = vbKeyRight Then
        If Shift = 0 Then
            If Not mIsExpandedView Then SetExpandedView True
            If Not mLstSheets Is Nothing Then
                If mLstSheets.Visible And mLstSheets.Enabled And mLstSheets.ListCount > 0 Then
                    On Error Resume Next
                    mLstSheets.SetFocus
                    On Error GoTo 0
                End If
            End If
            KeyCode = 0
            handled = True
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If Shift = 0 Then
            If mIsExpandedView Then SetExpandedView False
            If Me.lstWorkbooks.Visible And Me.lstWorkbooks.Enabled Then
                On Error Resume Next
                Me.lstWorkbooks.SetFocus
                On Error GoTo 0
            End If
            KeyCode = 0
            handled = True
        End If
    End If

    HandleGlobalKeyboardShortcuts = handled
End Function

Private Sub lstWorkbooks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If HandleGlobalKeyboardShortcuts(KeyCode, Shift) Then Exit Sub

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        mWorkbookKeyboardNavigation = True
        Exit Sub
    End If

    If Me.tglBatchMode.Value And Shift = 0 Then
        If KeyCode = vbKeyA Then
            If Me.lstWorkbooks.ListIndex > 0 Then
                Me.lstWorkbooks.Selected(Me.lstWorkbooks.ListIndex) = True
                RefreshVisuals
                RefreshSheetList
            End If
            KeyCode = 0
            Exit Sub
        ElseIf KeyCode = vbKeyD Then
            If Me.lstWorkbooks.ListIndex > 0 Then
                Me.lstWorkbooks.Selected(Me.lstWorkbooks.ListIndex) = False
                RefreshVisuals
                RefreshSheetList
            End If
            KeyCode = 0
            Exit Sub
        End If
    End If

    If KeyCode = vbKeyReturn Then
        mWorkbookKeyboardNavigation = False
        If Me.lstWorkbooks.ListIndex > 0 Then ActivateWorkbookFromListIndex Me.lstWorkbooks.ListIndex
        KeyCode = 0
        Exit Sub
    End If

    FixHeaderSelection
End Sub

Private Sub lstWorkbooks_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mWorkbookKeyboardNavigation = False
End Sub

Private Sub lstWorkbooks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mUIBusy Then Exit Sub
    If Me.lstWorkbooks.ListIndex > 0 Then ActivateWorkbookFromListIndex Me.lstWorkbooks.ListIndex
End Sub

Private Sub MoveWorkbookSelectionBy(ByVal delta As Long)
    Dim idx As Long

    If mUIBusy Then Exit Sub
    If delta = 0 Then Exit Sub
    If Me.lstWorkbooks.ListCount <= 1 Then Exit Sub

    idx = Me.lstWorkbooks.ListIndex
    If idx < 1 Then idx = 1
    idx = idx + delta

    If idx < 1 Then idx = 1
    If idx > Me.lstWorkbooks.ListCount - 1 Then idx = Me.lstWorkbooks.ListCount - 1

    If idx <> Me.lstWorkbooks.ListIndex Then
        Me.lstWorkbooks.ListIndex = idx
    End If

    ShowFullPathForIndex idx
End Sub

Private Sub ActivateWorkbookFromListIndex(ByVal idx As Long)
    Dim wbName As String
    Dim wb As Workbook

    If idx <= 0 Then Exit Sub
    If Me.tglBatchMode.Value Then
        RefreshVisuals
        RefreshSheetList
        Exit Sub
    End If

    If mActivatingWorkbookFromList Then Exit Sub
    mActivatingWorkbookFromList = True

    wbName = GetRawNameFromRow(idx)
    Set wb = GetWorkbookByName(wbName)
    If wb Is Nothing Then GoTo SafeExit

    On Error Resume Next
    If wb.Windows.Count > 0 Then wb.Windows(1).Activate
    wb.Activate
    Me.lstWorkbooks.SetFocus
    On Error GoTo 0

SafeExit:
    RefreshVisuals
    RefreshSheetList
    mActivatingWorkbookFromList = False
End Sub

Private Sub lstWorkbooks_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If mUIBusy Then Exit Sub

    If Button = MOUSE_BUTTON_RIGHT Then
        Me.tglBatchMode.Value = Not Me.tglBatchMode.Value
        ApplySelectionModeState
    End If
End Sub

Private Sub FixHeaderSelection()
    If mFixingHeader Then Exit Sub
    mFixingHeader = True

    On Error GoTo SafeExit

    If Me.lstWorkbooks.ListIndex = 0 Then
        On Error Resume Next
        Me.lstWorkbooks.Selected(0) = False
        Me.lstWorkbooks.ListIndex = -1
        Me.txtFullPath.Value = ""
        On Error GoTo 0
    End If

SafeExit:
    mFixingHeader = False
End Sub


Private Sub SoftReloadIfNeeded()
    ' throttle: not more often than every 2 seconds
    If mLastReload <> 0 Then
        If DateDiff("s", mLastReload, Now) < 2 Then Exit Sub
    End If
    mLastReload = Now
    ReloadListPreserveSelection
End Sub

' =========================================================
' SEARCH / RELOAD
' =========================================================
Private Sub txtSearch_Change()
    Dim caretPos As Long
    Dim txtLen As Long

    If mUIBusy Then Exit Sub

    caretPos = Me.txtSearch.SelStart
    txtLen = Len(Me.txtSearch.Text)

    ReloadListPreserveSelection False
    RefreshVisuals

    On Error Resume Next
    Me.txtSearch.SetFocus
    If caretPos > txtLen Then caretPos = txtLen
    Me.txtSearch.SelStart = caretPos
    Me.txtSearch.SelLength = 0
    On Error GoTo 0
End Sub

Private Sub btnReload_Click()
    ReloadListPreserveSelection
    AutoGrowFormWidthForFileColumn
    ApplyLayout
    RefreshVisuals

End Sub

Private Sub btnMaximize_Click()
    ApplyWindowAction 0
End Sub

Private Sub btnScreen1_Click()
    ApplyWindowAction 1
End Sub

Private Sub btnScreen2_Click()
    ApplyWindowAction 2
End Sub

Private Sub btnScreen3_Click()
    ApplyWindowAction 3
End Sub


' =========================================================
' SELECT / CLEAR
' =========================================================
Private Sub btnSelectAll_Click()
    Dim i As Long
    If Not Me.tglBatchMode.Value Then Exit Sub

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        Me.lstWorkbooks.Selected(i) = True
    Next i

    RefreshVisuals
    
End Sub

Private Sub btnClearAll_Click()
    If Not Me.tglBatchMode.Value Then Exit Sub
    ClearAllSelections
    RefreshVisuals
    
End Sub

Private Sub ClearAllSelections()
    Dim i As Long
    For i = 1 To Me.lstWorkbooks.ListCount - 1
        Me.lstWorkbooks.Selected(i) = False
    Next i
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' =========================================================
' ACTIONS
' =========================================================
Private Sub btnRefresh_Click()
    If Me.tglBatchMode.Value Then RunSelected False, True
End Sub

Private Sub btnSave_Click()
    If Me.tglBatchMode.Value Then RunSelected True, False
End Sub

Private Sub btnRefreshSave_Click()
    If Me.tglBatchMode.Value Then RunSelected True, True
End Sub

Private Sub ApplyWindowAction(ByVal targetScreen As Long)
    Dim targets As Collection
    Dim item As Variant
    Dim wb As Workbook
    Dim workLeft As Long, workTop As Long
    Dim workWidth As Long, workHeight As Long
    Dim prevScreenUpdating As Boolean
    Dim screenUpdatingCaptured As Boolean

    On Error GoTo EH

    Set targets = GetWorkbooksForWindowAction()
    If targets.Count = 0 Then
        RestoreNavigatorToFront
        Exit Sub
    End If

    If targetScreen > 0 Then
        If Not modWinAPI.TryGetMonitorWorkArea(targetScreen, workLeft, workTop, workWidth, workHeight) Then
            SafeMsgBox "Screen " & CStr(targetScreen) & " not available.", vbExclamation
            RestoreNavigatorToFront
            Exit Sub
        End If
    End If

    prevScreenUpdating = Application.ScreenUpdating
    screenUpdatingCaptured = True
    Application.ScreenUpdating = False
    modWinAPI.SuspendExcelWindowUpdates

    For Each item In targets
        Set wb = item
        If targetScreen <= 0 Then
            MaximizeWorkbookWindows wb
        Else
            MoveWorkbookToScreenAndMaximize wb, workLeft, workTop, workWidth, workHeight
        End If
    Next item

    modWinAPI.ResumeExcelWindowUpdates
    Application.ScreenUpdating = prevScreenUpdating
    RestoreNavigatorToFront
    Exit Sub

EH:
    modWinAPI.ResumeExcelWindowUpdates
    If screenUpdatingCaptured Then
        Application.ScreenUpdating = prevScreenUpdating
    Else
        Application.ScreenUpdating = True
    End If
    SafeMsgBox "Window action error: " & Err.Description, vbCritical
    RestoreNavigatorToFront
End Sub

Private Function GetWorkbooksForWindowAction() As Collection
    Dim col As Collection
    Dim i As Long
    Dim wb As Workbook
    Dim wbName As String

    Set col = New Collection

    If Me.tglBatchMode.Value Then
        For i = 1 To Me.lstWorkbooks.ListCount - 1
            If Me.lstWorkbooks.Selected(i) Then
                wbName = GetRawNameFromRow(i)
                Set wb = GetWorkbookByName(wbName)
                If Not wb Is Nothing Then
                    If Not IsWorkbookSkippable(wb) Then col.Add wb
                End If
            End If
        Next i

        If col.Count = 0 Then
            SafeMsgBox "Turn ON Selection mode and select at least one file.", vbExclamation
        End If
    Else
        For Each wb In Application.Workbooks
            If Not IsWorkbookSkippable(wb) Then col.Add wb
        Next wb
    End If

    Set GetWorkbooksForWindowAction = col
End Function

Private Function MaximizeWorkbookWindows(ByVal wb As Workbook) As Long
    Dim win As Window

    On Error Resume Next
    For Each win In wb.Windows
        win.WindowState = xlMaximized
        If Err.Number = 0 Then MaximizeWorkbookWindows = MaximizeWorkbookWindows + 1
        Err.Clear
    Next win
    On Error GoTo 0
End Function

Private Function MoveWorkbookToScreenAndMaximize(ByVal wb As Workbook, _
                                                 ByVal workLeft As Long, _
                                                 ByVal workTop As Long, _
                                                 ByVal workWidth As Long, _
                                                 ByVal workHeight As Long) As Long
    Dim win As Window

    On Error Resume Next
    For Each win In wb.Windows
        win.WindowState = xlNormal
        win.Left = workLeft
        win.TOP = workTop
        win.WindowState = xlMaximized
        If Err.Number = 0 Then MoveWorkbookToScreenAndMaximize = MoveWorkbookToScreenAndMaximize + 1
        Err.Clear
    Next win
    On Error GoTo 0
End Function

Private Sub RunSelected(ByVal doSave As Boolean, ByVal doRefresh As Boolean)
    Dim selectedNames As Collection
    Dim item As Variant
    Dim wb As Workbook
    Dim wbName As String
    Dim cntSel As Long
    Dim cntRefOK As Long
    Dim cntTO As Long
    Dim cntSav As Long
    Dim cntSkip As Long

    On Error GoTo EH

    If mBatchRunning Then Exit Sub

    Set selectedNames = GetSelectedWorkbookNames()
    cntSel = selectedNames.Count

    If cntSel = 0 Then
        SafeMsgBox "Nothing selected.", vbExclamation
        Exit Sub
    End If

    mBatchRunning = True
    SetBatchUI True
    mCancelBatch = False

    If doRefresh And doSave Then
        For Each item In selectedNames
            wbName = CStr(item)
            Set wb = GetWorkbookByName(wbName)

            If Not wb Is Nothing Then
                If RefreshOneWorkbook(wb) Then
                    cntRefOK = cntRefOK + 1
                ElseIf IsTimedOut(wb.Name) Then
                    cntTO = cntTO + 1
                End If

                If mCancelBatch Then
                    btnCancel.enabled = False
                    Exit For
                End If
            End If
        Next item

        If Not mCancelBatch Then
            For Each item In selectedNames
                wbName = CStr(item)
                Set wb = GetWorkbookByName(wbName)

                If Not wb Is Nothing Then
                    If IsTimedOut(wb.Name) Then
                        cntSkip = cntSkip + 1
                    Else
                        If SaveOneWorkbook(wb) Then
                            cntSav = cntSav + 1
                        End If
                    End If
                End If
            Next item
        End If

        SafeMsgBox "Done: refreshed " & cntRefOK & _
                   ", timed out " & cntTO & _
                   ", saved " & cntSav & _
                   ", skipped (timeout) " & cntSkip & " file(s).", vbInformation

    ElseIf doRefresh Then
        For Each item In selectedNames
            wbName = CStr(item)
            Set wb = GetWorkbookByName(wbName)

            If Not wb Is Nothing Then
                If RefreshOneWorkbook(wb) Then
                    cntRefOK = cntRefOK + 1
                ElseIf IsTimedOut(wb.Name) Then
                    cntTO = cntTO + 1
                End If

                If mCancelBatch Then
                    btnCancel.enabled = False
                    Exit For
                End If
            End If
        Next item

        SafeMsgBox "Done: refreshed " & cntRefOK & _
                   ", timed out " & cntTO & " file(s).", vbInformation

    ElseIf doSave Then
        For Each item In selectedNames
            wbName = CStr(item)
            Set wb = GetWorkbookByName(wbName)

            If Not wb Is Nothing Then
                If IsTimedOut(wb.Name) Then
                    cntSkip = cntSkip + 1
                Else
                    If SaveOneWorkbook(wb) Then
                        cntSav = cntSav + 1
                    End If
                End If
            End If
        Next item

        SafeMsgBox "Done: saved " & cntSav & _
                   ", skipped (timeout) " & cntSkip & " file(s).", vbInformation
    End If

    ReloadListPreserveSelection
    RefreshVisuals

FINALLY:
    mCancelBatch = False
    mBatchRunning = False
    Application.StatusBar = False
    SetBatchUI False
    ForceTopMost
    Exit Sub

EH:
    mCancelBatch = False
    mBatchRunning = False
    Application.StatusBar = False
    SetBatchUI False
    SafeMsgBox "RunSelected error: " & Err.Description, vbCritical
End Sub



Private Function RefreshOneWorkbook(ByVal wb As Workbook) As Boolean
' NOTE:
' RefreshAll is controlled by Excel engine, not VBA.
' Once started, it cannot be safely stopped from code.
'
' Any attempt to abort RefreshAll, manipulate UI,
' or unload forms during refresh may crash Excel.
'
' This procedure waits for refresh completion
' or timeout, but never forces termination.


    Dim key As String
    key = wb.Name

    SetRefreshing key, True
    UpdateRowStatus key
    DoEvents

    On Error GoTo EH

    wb.RefreshAll

    If WaitForRefreshToFinish(wb, REFRESH_TIMEOUT_SEC) Then
        SetRefreshed key
        RefreshOneWorkbook = True
    Else
        SetTimedOut key
        RefreshOneWorkbook = False
    End If

    UpdateRowStatus key
    Exit Function

EH:
    SetRefreshing key, False
    UpdateRowStatus key
    SafeMsgBox "Refresh error in: " & wb.Name & vbCrLf & Err.Description, vbCritical
    RefreshOneWorkbook = False
End Function




Private Function SaveOneWorkbook(ByVal wb As Workbook) As Boolean
    If IsTimedOut(wb.Name) Then Exit Function
    Dim chosenPath As Variant
    Dim defaultExt As String
    Dim defaultName As String

    On Error GoTo EH

    If wb.ReadOnly Then
        SafeMsgBox "Workbook is read-only, cannot save: " & wb.Name, vbExclamation
        Exit Function
    End If

    If Len(wb.Path) = 0 Then
        defaultExt = IIf(wb.HasVBProject, ".xlsm", ".xlsx")
        defaultName = wb.Name
        If InStrRev(defaultName, ".") = 0 Then defaultName = defaultName & defaultExt

        chosenPath = Application.GetSaveAsFilename( _
            InitialFileName:=defaultName, _
            FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb" _
        )

        If chosenPath = False Then Exit Function

        wb.SaveAs fileName:=CStr(chosenPath)
    Else
        wb.Save
    End If

    SaveOneWorkbook = True
    Exit Function

EH:
    SafeMsgBox "Save error in: " & wb.Name & vbCrLf & Err.Description, vbCritical
    SaveOneWorkbook = False
End Function


' =========================================================
' LIST BUILD
' =========================================================
Private Sub ReloadListPreserveSelection(Optional ByVal selectFirstDataRow As Boolean = True)
    Dim dict As Object
    Dim i As Long
    Dim nm As String

    If mUIBusy Then Exit Sub
    mUIBusy = True

    On Error GoTo SafeExit

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            dict(GetRawNameFromRow(i)) = True
        End If
    Next i

    LoadWorkbookCache
    ApplyFilterAndFillList selectFirstDataRow

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        nm = GetRawNameFromRow(i)
        If dict.Exists(nm) Then
            Me.lstWorkbooks.Selected(i) = True
        End If
    Next i

    If selectFirstDataRow Then
        FixHeaderSelection
    Else
        Me.lstWorkbooks.ListIndex = -1
    End If

SafeExit:
    mUIBusy = False
End Sub


Private Sub LoadWorkbookCache()

    Dim wb As Workbook
    Dim cnt As Long

    cnt = 0
    For Each wb In Application.Workbooks
        If Not IsWorkbookSkippable(wb) Then cnt = cnt + 1
    Next wb

    mAllCount = cnt
    If mAllCount = 0 Then Exit Sub

    ReDim mAllNames(1 To cnt)
    ReDim mAllFullNames(1 To cnt)

    cnt = 0
    For Each wb In Application.Workbooks
        If Not IsWorkbookSkippable(wb) Then
            cnt = cnt + 1
            mAllNames(cnt) = wb.Name
            mAllFullNames(cnt) = wb.fullName
        End If
    Next wb

    If mAllCount > 1 Then SortWorkbooks 1, mAllCount
End Sub

Private Sub SortWorkbooks(ByVal lo As Long, ByVal hi As Long)

    Dim i As Long, j As Long
    Dim pivot As String
    Dim t1 As String, t2 As String

    i = lo: j = hi
    pivot = LCase$(mAllNames((lo + hi) \ 2))

    Do While i <= j
        Do While LCase$(mAllNames(i)) < pivot: i = i + 1: Loop
        Do While LCase$(mAllNames(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            t1 = mAllNames(i): mAllNames(i) = mAllNames(j): mAllNames(j) = t1
            t2 = mAllFullNames(i): mAllFullNames(i) = mAllFullNames(j): mAllFullNames(j) = t2
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then SortWorkbooks lo, j
    If i < hi Then SortWorkbooks i, hi
End Sub

Private Sub ApplyFilterAndFillList(Optional ByVal selectFirstDataRow As Boolean = True)
    Dim q As String
    Dim i As Long
    Dim nm As String
    Dim fullP As String

    q = LCase$(Trim$(CStr(Me.txtSearch.Value)))

    Me.lstWorkbooks.Clear

    Me.lstWorkbooks.AddItem "File"
    Me.lstWorkbooks.List(0, 1) = "Dir"
    Me.lstWorkbooks.List(0, 2) = "Sync"
    Me.lstWorkbooks.List(0, 3) = ""

    If mAllCount = 0 Then
        Me.lstWorkbooks.ListIndex = -1
        UpdateFileCounterLabel
        Exit Sub
    End If

    For i = 1 To mAllCount
        nm = mAllNames(i)
        fullP = mAllFullNames(i)

        If (q = "") _
           Or (InStr(1, LCase$(nm), q, vbTextCompare) > 0) _
           Or (InStr(1, LCase$(fullP), q, vbTextCompare) > 0) Then

            Me.lstWorkbooks.AddItem nm
            Me.lstWorkbooks.List(Me.lstWorkbooks.ListCount - 1, 1) = GetLocationTag(fullP)
            Me.lstWorkbooks.List(Me.lstWorkbooks.ListCount - 1, 2) = GetStatusText(nm)
            Me.lstWorkbooks.List(Me.lstWorkbooks.ListCount - 1, 3) = fullP
        End If
    Next i

    If selectFirstDataRow Then
        If Me.lstWorkbooks.ListCount > 1 Then
            Me.lstWorkbooks.ListIndex = 1
            Me.lstWorkbooks.Selected(0) = False
        Else
            Me.lstWorkbooks.ListIndex = -1
        End If
    Else
        Me.lstWorkbooks.ListIndex = -1
    End If

    UpdateFileCounterLabel
End Sub

Private Sub UpdateFileCounterLabel()
    Dim visibleCount As Long

    visibleCount = Me.lstWorkbooks.ListCount - 1
    If visibleCount < 0 Then visibleCount = 0

    Me.Label3.Caption = "Files: " & CStr(visibleCount) & " / " & CStr(mAllCount)
    If mBaseInsideW > 0 Then ApplyLayout

End Sub

Private Sub RefreshVisuals()

    Dim i As Long
    Dim activeName As String
    Dim rawName As String

    activeName = ""
    On Error Resume Next
    If Not Application.ActiveWorkbook Is Nothing Then activeName = Application.ActiveWorkbook.Name
    On Error GoTo 0

    ' Skip header (0)
    For i = 1 To Me.lstWorkbooks.ListCount - 1

        rawName = GetRawNameFromRow(i)

        ' File column: show prefix only for active workbook, but never modify stored raw name
        If rawName = activeName Then
            Me.lstWorkbooks.List(i, 0) = ACTIVE_PREFIX & rawName
        Else
            Me.lstWorkbooks.List(i, 0) = rawName
        End If

        ' Status column based on RAW name
        Me.lstWorkbooks.List(i, 2) = GetStatusText(rawName)
    Next i

    RefreshSheetList
End Sub

Private Function GetRawNameFromRow(ByVal rowIndex As Long) As String
    Dim s As String
    s = CStr(Me.lstWorkbooks.List(rowIndex, 0))
    If Left$(s, Len(ACTIVE_PREFIX)) = ACTIVE_PREFIX Then
        GetRawNameFromRow = Mid$(s, Len(ACTIVE_PREFIX) + 1)
    Else
        GetRawNameFromRow = s
    End If
End Function

' =========================================================
' TOPMOST (requires WinAPI in modWinAPI)
' =========================================================
Private Sub ForceTopMost()
End Sub

' =========================================================
' STATUS / LOCATION
' =========================================================
Private Sub SetRefreshing(ByVal wbName As String, ByVal isRefreshing As Boolean)
    Dim arr As Variant

    If Not mStatus.Exists(wbName) Then
        arr = Array(CDate(0), isRefreshing, False)
        mStatus.Add wbName, arr
    Else
        arr = mStatus(wbName)

        If Not IsArray(arr) Or UBound(arr) < 2 Then
            arr = Array(CDate(0), False, False)
        End If

        arr(1) = isRefreshing
        If isRefreshing Then
            arr(2) = False
            If UBound(arr) >= 3 Then arr(3) = False
        End If

        mStatus(wbName) = arr
    End If
End Sub



Private Sub SetRefreshed(ByVal wbName As String)
    Dim arr As Variant

    If Not mStatus.Exists(wbName) Then
        arr = Array(Now, False, False)
        mStatus.Add wbName, arr
    Else
        arr = mStatus(wbName)
        If Not IsArray(arr) Or UBound(arr) < 2 Then
            arr = Array(CDate(0), False, False)
        End If

        arr(0) = Now
        arr(1) = False
        arr(2) = False
        mStatus(wbName) = arr
    End If
End Sub

Private Function GetStatusText(ByVal wbName As String) As String
    Dim v As Variant

    If Not mStatus.Exists(wbName) Then
        GetStatusText = "Never"
        Exit Function
    End If

    v = mStatus(wbName)
    If IsArray(v) And UBound(v) >= 3 Then
    If CBool(v(3)) Then
        GetStatusText = "Cancelled"
        Exit Function
    End If
End If


    ' wstecznie: jak ktos ma tylko 2 pola
    If Not IsArray(v) Or UBound(v) < 1 Then
        GetStatusText = "Never"
        Exit Function
    End If

    ' Timed out ma priorytet
    If IsArray(v) And UBound(v) >= 2 Then
        If CBool(v(2)) Then
            GetStatusText = "Timed out"
            Exit Function
        End If
    End If

    If CBool(v(1)) Then
        GetStatusText = "Refreshing..."
    ElseIf CDate(v(0)) = 0 Then
        GetStatusText = "Never"
    Else
        GetStatusText = FormatAge(CDate(v(0)))
    End If
End Function

Private Function FormatAge(ByVal d As Date) As String
    Dim m As Long, h As Long, dy As Long
    m = DateDiff("n", d, Now)
    If m < 60 Then FormatAge = m & " min ago": Exit Function
    h = DateDiff("h", d, Now)
    If h < 24 Then FormatAge = h & " h ago": Exit Function
    dy = DateDiff("d", d, Now)
    FormatAge = dy & " d ago"
End Function

Private Function GetLocationTag(ByVal fullPath As String) As String
    If LCase$(Left$(fullPath, 4)) = "http" Then
        GetLocationTag = "WEB"
    ElseIf Len(fullPath) >= 2 And Mid$(fullPath, 2, 1) = ":" Then
        GetLocationTag = Left$(fullPath, 2)
    Else
        GetLocationTag = "-"
    End If
End Function

Private Sub UpdateRowStatus(ByVal wbName As String)
    Dim i As Long
    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If GetRawNameFromRow(i) = wbName Then
            Me.lstWorkbooks.List(i, 2) = GetStatusText(wbName)
            Exit Sub
        End If
    Next i
End Sub

Private Sub SetTimedOut(ByVal wbName As String)
    Dim arr As Variant

    If Not mStatus.Exists(wbName) Then
        arr = Array(CDate(0), False, True)
        mStatus.Add wbName, arr
    Else
        arr = mStatus(wbName)
        If Not IsArray(arr) Or UBound(arr) < 2 Then
            arr = Array(CDate(0), False, True)
        Else
            arr(1) = False
            arr(2) = True
        End If
        mStatus(wbName) = arr
    End If
End Sub
Private Sub SetCancelled(ByVal wbName As String)
' STATUS: Cancelled
' Means: batch execution was stopped by user request.
' It does NOT mean refresh failed or timed out.
'
' The workbook may be fully refreshed,
' partially refreshed, or unchanged 
' Excel decides, not VBA.

    Dim arr As Variant

    If Not mStatus.Exists(wbName) Then
        arr = Array(CDate(0), False, False, True)
        mStatus.Add wbName, arr
    Else
        arr = mStatus(wbName)

        If Not IsArray(arr) Or UBound(arr) < 3 Then
            ReDim Preserve arr(0 To 3)
        End If

        arr(1) = False
        arr(3) = True
        mStatus(wbName) = arr
    End If
End Sub

Private Function IsCancelled(ByVal wbName As String) As Boolean
    Dim v As Variant
    If Not mStatus.Exists(wbName) Then Exit Function

    v = mStatus(wbName)
    If IsArray(v) And UBound(v) >= 3 Then
        IsCancelled = CBool(v(3))
    End If
End Function

Private Function IsTimedOut(ByVal wbName As String) As Boolean
    Dim v As Variant
    If Not mStatus.Exists(wbName) Then Exit Function

    v = mStatus(wbName)
    If IsArray(v) Then
        If UBound(v) >= 2 Then IsTimedOut = CBool(v(2))
    End If
End Function


' =========================================================
' HELPERS
' =========================================================
Private Function GetControlIfExists(ByVal controlName As String) As Object
    On Error Resume Next
    Set GetControlIfExists = Me.Controls(controlName)
    On Error GoTo 0
End Function

Private Sub SetOptionalControlEnabled(ByVal controlName As String, ByVal enabled As Boolean)
    Dim ctl As Object
    Set ctl = GetControlIfExists(controlName)
    If ctl Is Nothing Then Exit Sub
    ctl.enabled = enabled
End Sub

Private Function GetOptionalControlTop(ByVal controlName As String, ByVal fallbackTop As Single) As Single
    Dim ctl As Object
    Set ctl = GetControlIfExists(controlName)
    If ctl Is Nothing Then
        GetOptionalControlTop = fallbackTop
    Else
        GetOptionalControlTop = ctl.TOP
    End If
End Function

Private Function GetOptionalControlLeft(ByVal controlName As String, ByVal fallbackLeft As Single) As Single
    Dim ctl As Object
    Set ctl = GetControlIfExists(controlName)
    If ctl Is Nothing Then
        GetOptionalControlLeft = fallbackLeft
    Else
        GetOptionalControlLeft = ctl.Left
    End If
End Function

Private Sub SetOptionalControlTop(ByVal controlName As String, ByVal newTop As Single)
    Dim ctl As Object
    Set ctl = GetControlIfExists(controlName)
    If ctl Is Nothing Then Exit Sub
    ctl.TOP = newTop
End Sub

Private Function GetWorkbookByName(ByVal wbName As String) As Workbook
    On Error Resume Next
    Set GetWorkbookByName = Application.Workbooks(wbName)
    On Error GoTo 0
End Function

Private Function IsWorkbookSkippable(ByVal wb As Workbook) As Boolean
    Dim nm As String
    nm = UCase$(wb.Name)
    IsWorkbookSkippable = (nm = "PERSONAL.XLSB" Or Right$(nm, 5) = ".XLAM")
End Function

Private Sub EnsureRightPanelControls()
    Dim ctlBtn As Object
    Dim ctlList As Object
    Dim ctlLbl As Object

    Set ctlBtn = GetControlIfExists("btnTogglePanel")
    If ctlBtn Is Nothing Then
        Set ctlBtn = Me.Controls.Add("Forms.CommandButton.1", "btnTogglePanel", True)
    End If
    Set mBtnTogglePanel = ctlBtn
    With mBtnTogglePanel
        .Caption = PANEL_TOGGLE_COLLAPSED
        .Width = Me.btnClose.Width
        .Height = Me.btnClose.Height
        .Visible = True
    End With

    Set ctlList = GetControlIfExists("lstSheets")
    If ctlList Is Nothing Then
        Set ctlList = Me.Controls.Add("Forms.ListBox.1", "lstSheets", True)
    End If
    Set mLstSheets = ctlList
    With mLstSheets
        .Visible = False
        .IntegralHeight = False
        .ColumnCount = 1
        .BoundColumn = 1
    End With

    Set ctlLbl = GetControlIfExists("lblSheetsWorkbook")
    If ctlLbl Is Nothing Then
        Set ctlLbl = Me.Controls.Add("Forms.Label.1", "lblSheetsWorkbook", True)
    End If
    Set mLblSheetsWorkbook = ctlLbl
    With mLblSheetsWorkbook
        .Caption = ""
        .Visible = False
        .WordWrap = False
    End With
End Sub

Private Function GetSheetPanelWidthPt() As Single
    Dim charWidth As Single
    charWidth = MeasureTextWidthPt(String$(33, "W"))
    GetSheetPanelWidthPt = (charWidth + 20) * 0.45
    If GetSheetPanelWidthPt < 90 Then GetSheetPanelWidthPt = 90
End Function

Private Sub SetExpandedView(ByVal expanded As Boolean)
    Dim frameW As Single
    Dim targetInsideW As Single

    mIsExpandedView = expanded
    frameW = Me.Width - Me.InsideWidth
    If mCollapsedInsideW <= 0 Then mCollapsedInsideW = Me.InsideWidth
    If mPanelWidth <= 0 Then mPanelWidth = GetSheetPanelWidthPt()

    If expanded Then
        mCollapsedInsideW = Me.InsideWidth
        targetInsideW = mCollapsedInsideW + mGap + mPanelWidth
    Else
        targetInsideW = mCollapsedInsideW
    End If

    Me.Width = targetInsideW + frameW

    If mBtnTogglePanel Is Nothing Then EnsureRightPanelControls
    mBtnTogglePanel.Caption = IIf(expanded, PANEL_TOGGLE_EXPANDED, PANEL_TOGGLE_COLLAPSED)
    mLstSheets.Visible = expanded
    If Not mLblSheetsWorkbook Is Nothing Then mLblSheetsWorkbook.Visible = expanded

    ApplyLayout
    PositionTopButtons
    RefreshSheetList
End Sub

Private Function GetCurrentWorkbookForSheets() As Workbook
    Dim idx As Long
    Dim wbName As String

    idx = Me.lstWorkbooks.ListIndex
    If idx <= 0 Then Exit Function

    wbName = GetRawNameFromRow(idx)
    Set GetCurrentWorkbookForSheets = GetWorkbookByName(wbName)
End Function

Private Sub RefreshSheetList()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim activeSheetName As String
    Dim rowText As String
    Dim selectedSheetName As String
    Dim selectedIndex As Long
    Dim restoreWorkbookFocus As Boolean

    If mLstSheets Is Nothing Then Exit Sub

    On Error Resume Next
    restoreWorkbookFocus = (Not Me.ActiveControl Is Nothing And Me.ActiveControl.Name = "lstWorkbooks")
    On Error GoTo SafeExit

    mUpdatingSheets = True
    selectedSheetName = GetRawSheetNameFromRow(mLstSheets.ListIndex)
    mLstSheets.Clear
    If Not mLblSheetsWorkbook Is Nothing Then mLblSheetsWorkbook.Caption = ""
    Set wb = GetCurrentWorkbookForSheets()
    If wb Is Nothing Then
        mUpdatingSheets = False
        Exit Sub
    End If

    If Not mLblSheetsWorkbook Is Nothing Then mLblSheetsWorkbook.Caption = wb.Name

    On Error Resume Next
    activeSheetName = CStr(wb.ActiveSheet.Name)
    On Error GoTo 0

    For Each ws In wb.Worksheets
        rowText = ws.Name
        If ws.Name = activeSheetName Then rowText = ACTIVE_PREFIX & rowText
        mLstSheets.AddItem rowText
    Next ws

    selectedIndex = -1
    If Len(selectedSheetName) > 0 Then
        For selectedIndex = 0 To mLstSheets.ListCount - 1
            If GetRawSheetNameFromRow(selectedIndex) = selectedSheetName Then Exit For
        Next selectedIndex
        If selectedIndex >= mLstSheets.ListCount Then selectedIndex = -1
    End If

    If selectedIndex = -1 And Len(activeSheetName) > 0 Then
        For selectedIndex = 0 To mLstSheets.ListCount - 1
            If GetRawSheetNameFromRow(selectedIndex) = activeSheetName Then Exit For
        Next selectedIndex
        If selectedIndex >= mLstSheets.ListCount Then selectedIndex = -1
    End If

    If selectedIndex >= 0 Then mLstSheets.ListIndex = selectedIndex

SafeExit:
    mUpdatingSheets = False

    If restoreWorkbookFocus Then
        On Error Resume Next
        Me.lstWorkbooks.SetFocus
        On Error GoTo 0
    End If
End Sub

Private Function GetRawSheetNameFromRow(ByVal rowIndex As Long) As String
    Dim s As String

    If mLstSheets Is Nothing Then Exit Function
    If rowIndex < 0 Or rowIndex >= mLstSheets.ListCount Then Exit Function

    s = CStr(mLstSheets.List(rowIndex, 0))
    If Left$(s, Len(ACTIVE_PREFIX)) = ACTIVE_PREFIX Then
        GetRawSheetNameFromRow = Mid$(s, Len(ACTIVE_PREFIX) + 1)
    Else
        GetRawSheetNameFromRow = s
    End If
End Function

Private Sub mBtnTogglePanel_Click()
    SetExpandedView Not mIsExpandedView
End Sub

Private Sub ActivateSheetFromSheetList()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String

    If mUpdatingSheets Then Exit Sub
    If mUIBusy Then Exit Sub
    If mActivatingSheetFromList Then Exit Sub
    If mLstSheets.ListIndex < 0 Then Exit Sub

    mActivatingSheetFromList = True

    Set wb = GetCurrentWorkbookForSheets()
    If wb Is Nothing Then GoTo SafeExit

    sheetName = GetRawSheetNameFromRow(mLstSheets.ListIndex)
    If Len(sheetName) = 0 Then GoTo SafeExit

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then GoTo SafeExit

    On Error Resume Next
    If wb.Windows.Count > 0 Then wb.Windows(1).Activate
    wb.Activate
    ws.Activate
    On Error GoTo 0

    RefreshVisuals
    On Error Resume Next
    mLstSheets.SetFocus
    On Error GoTo 0

SafeExit:
    mActivatingSheetFromList = False
End Sub

Private Sub mLstSheets_Click()
    If mUpdatingSheets Then Exit Sub
    If mUIBusy Then Exit Sub
    If mSheetKeyboardNavigation Then Exit Sub

    On Error Resume Next
    mLstSheets.SetFocus
    On Error GoTo 0
    ActivateSheetFromSheetList
End Sub

Private Sub mLstSheets_Change()
End Sub

Private Sub mLstSheets_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If HandleGlobalKeyboardShortcuts(KeyCode, Shift) Then Exit Sub

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        mSheetKeyboardNavigation = True
        Exit Sub
    End If

    If KeyCode = vbKeyReturn Then
        mSheetKeyboardNavigation = False
        ActivateSheetFromSheetList
        KeyCode = 0
    End If
End Sub

Private Sub mLstSheets_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mSheetKeyboardNavigation = False
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC, "W", IIf(mIsExpandedView, Me.Width - (mGap + mPanelWidth), Me.Width)
    SaveSetting REG_APP, REG_SEC, "H", Me.Height
End Sub


Private Sub CacheLayout()
    ' Cache once after size is restored
    mBaseInsideW = Me.InsideWidth
    mBaseInsideH = Me.InsideHeight

    mRightMargin = 8
    mGap = 8
    mBaseListLeft = Me.lstWorkbooks.Left

    ' Bottom block anchor (we shift these by deltaH)
    mCtlTop(1) = Me.tglBatchMode.TOP
    mCtlTop(2) = Me.btnSelectAll.TOP
    mCtlTop(3) = Me.btnClearAll.TOP
    mCtlTop(4) = Me.btnRefresh.TOP
    mCtlTop(5) = Me.btnSave.TOP
    mCtlTop(6) = Me.btnRefreshSave.TOP
    mCtlTop(7) = Me.txtSuffix.TOP
    mCtlTop(8) = Me.btnCopyBreakLinks.TOP
    mCtlTop(9) = Me.btnOpenCopyFolder.TOP
    mCtlTop(10) = Me.btnClose.TOP
    mCtlTop(11) = Me.Label2.TOP   ' Suffix label
    mCtlTop(12) = Me.btnCloseSelected.TOP
    mCtlTop(13) = Me.txtFullPath.TOP
    mCtlTop(14) = Me.btnOpenFile.TOP
    mCtlTop(15) = Me.btnOpenAndRefresh.TOP
    mCtlTop(16) = Me.btnCopyWithSuffix.TOP
    mCtlTop(17) = GetOptionalControlTop("btnMaximize", mCtlTop(16))
    mCtlTop(18) = GetOptionalControlTop("btnScreen1", mCtlTop(17))
    mCtlTop(19) = GetOptionalControlTop("btnScreen2", mCtlTop(18))
    mCtlTop(20) = GetOptionalControlTop("btnScreen3", mCtlTop(19))
    mCtlTop(21) = GetOptionalControlTop("ChckBox1", GetOptionalControlTop("CheckBox1", mCtlTop(16)))
    mOpenCopiedOffsetTop = mCtlTop(21) - mCtlTop(16)
    mOpenCopiedOffsetLeft = GetOptionalControlLeft("ChckBox1", GetOptionalControlLeft("CheckBox1", Me.btnCopyWithSuffix.Left)) - Me.btnCopyWithSuffix.Left

    
    mBottomBlockTop = Me.tglBatchMode.TOP
    mTopBlockBottom = Me.txtSearch.TOP + Me.txtSearch.Height

    ' Start file column width from current setting
    mBaseFileColW = 240.5
End Sub

Private Sub ApplyLayout()
    Dim deltaH As Single
    Dim newFileW As Single
    Dim rightX As Single
    Dim topBlockBottom As Single
    Dim actionBottomMargin As Single
    Dim actionGap As Single
    Dim actionTop As Single
    Dim ctlMax As Object
    Dim ctlS1 As Object
    Dim ctlS2 As Object
    Dim ctlS3 As Object
    Dim ctlOpenCopied As Object
    Dim reservedRight As Single
    Dim sheetBottom As Single

    If mBaseInsideW = 0 Or mBaseInsideH = 0 Then Exit Sub

    deltaH = Me.InsideHeight - mBaseInsideH

    PinImage1AndTopRow

    ' --- Top row: Search stretches, Reload stays right
    rightX = Me.InsideWidth - mRightMargin
    Me.btnReload.Left = rightX - Me.btnReload.Width
    Me.txtSearch.Width = (Me.btnReload.Left - mGap) - Me.txtSearch.Left
    Me.Label1.TextAlign = fmTextAlignRight
    Me.Label1.Left = Me.txtSearch.Left - Me.Label1.Width

    ' Label3: left-aligned to Reload, below it, between top row and list
    Me.Label3.Left = Me.btnReload.Left
    Me.Label3.TOP = Me.btnReload.TOP + Me.btnReload.Height + 2

    ' Dynamic top block bottom so list stays below Label3
    topBlockBottom = Me.txtSearch.TOP + Me.txtSearch.Height
    If Me.btnReload.TOP + Me.btnReload.Height > topBlockBottom Then
        topBlockBottom = Me.btnReload.TOP + Me.btnReload.Height
    End If
    If Me.Label3.TOP + Me.Label3.Height > topBlockBottom Then
        topBlockBottom = Me.Label3.TOP + Me.Label3.Height
    End If

    ' Label3: left-aligned to Reload, below it, between top row and list
    Me.Label3.Left = Me.btnReload.Left
    Me.Label3.TOP = Me.btnReload.TOP + Me.btnReload.Height + 2

    ' Dynamic top block bottom so list stays below Label3
    topBlockBottom = Me.txtSearch.TOP + Me.txtSearch.Height
    If Me.btnReload.TOP + Me.btnReload.Height > topBlockBottom Then
        topBlockBottom = Me.btnReload.TOP + Me.btnReload.Height
    End If
    If Me.Label3.TOP + Me.Label3.Height > topBlockBottom Then
        topBlockBottom = Me.Label3.TOP + Me.Label3.Height
    End If

    ' Label3: left-aligned to Reload, below it, between top row and list
    Me.Label3.Left = Me.btnReload.Left
    Me.Label3.TOP = Me.btnReload.TOP + Me.btnReload.Height + 2

    ' Dynamic top block bottom so list stays below Label3
    topBlockBottom = Me.txtSearch.TOP + Me.txtSearch.Height
    If Me.btnReload.TOP + Me.btnReload.Height > topBlockBottom Then
        topBlockBottom = Me.btnReload.TOP + Me.btnReload.Height
    End If
    If Me.Label3.TOP + Me.Label3.Height > topBlockBottom Then
        topBlockBottom = Me.Label3.TOP + Me.Label3.Height
    End If

    ' --- Shift bottom controls by deltaH (keep logical places)
    Me.tglBatchMode.TOP = mCtlTop(1) + deltaH
    Me.btnSelectAll.TOP = mCtlTop(2) + deltaH
    Me.btnClearAll.TOP = mCtlTop(3) + deltaH
    Me.btnRefresh.TOP = mCtlTop(4) + deltaH
    Me.btnSave.TOP = mCtlTop(5) + deltaH
    Me.btnRefreshSave.TOP = mCtlTop(6) + deltaH
    Me.txtSuffix.TOP = mCtlTop(7) + deltaH
    Me.btnCopyBreakLinks.TOP = mCtlTop(8) + deltaH
    Me.btnOpenCopyFolder.TOP = mCtlTop(9) + deltaH
    Me.btnClose.TOP = mCtlTop(10) + deltaH
    Me.Label2.TOP = mCtlTop(11) + deltaH
    Me.btnCloseSelected.TOP = mCtlTop(12) + deltaH
    Me.btnOpenFile.TOP = mCtlTop(14) + deltaH
    Me.btnOpenAndRefresh.TOP = mCtlTop(15) + deltaH
    Me.btnCopyWithSuffix.TOP = mCtlTop(16) + deltaH
    SetOptionalControlTop "btnMaximize", mCtlTop(17) + deltaH
    SetOptionalControlTop "btnScreen1", mCtlTop(18) + deltaH
    SetOptionalControlTop "btnScreen2", mCtlTop(19) + deltaH
    SetOptionalControlTop "btnScreen3", mCtlTop(20) + deltaH
    ' Keep bottom action buttons in one line:
    ' btnMaximize btnScreen1 btnScreen2 btnScreen3 btnCancel btnClose
    actionBottomMargin = 6
    actionGap = 6
    actionTop = Me.InsideHeight - actionBottomMargin - Me.btnClose.Height

    Me.btnClose.TOP = actionTop
    Me.btnCancel.TOP = actionTop
    If Not mBtnTogglePanel Is Nothing Then
        mBtnTogglePanel.TOP = actionTop - actionGap - mBtnTogglePanel.Height
    End If

    Me.btnClose.Left = Me.InsideWidth - actionBottomMargin - Me.btnClose.Width
    Me.btnCancel.Left = Me.btnClose.Left - actionGap - Me.btnCancel.Width
    If Not mBtnTogglePanel Is Nothing Then
        mBtnTogglePanel.Left = Me.btnClose.Left
        mBtnTogglePanel.Width = Me.btnClose.Width
        mBtnTogglePanel.Height = Me.btnClose.Height
    End If

    Set ctlMax = GetControlIfExists("btnMaximize")
    Set ctlS1 = GetControlIfExists("btnScreen1")
    Set ctlS2 = GetControlIfExists("btnScreen2")
    Set ctlS3 = GetControlIfExists("btnScreen3")
    Set ctlOpenCopied = GetControlIfExists("ChckBox1")
    If ctlOpenCopied Is Nothing Then Set ctlOpenCopied = GetControlIfExists("CheckBox1")

    If Not ctlMax Is Nothing Then
        ctlMax.TOP = actionTop
        ctlMax.Left = Me.btnCopyWithSuffix.Left
    End If

    If Not ctlS1 Is Nothing Then
        ctlS1.TOP = actionTop
        If Not ctlMax Is Nothing Then
            ctlS1.Left = ctlMax.Left + ctlMax.Width + actionGap
        End If
    End If

    If Not ctlS2 Is Nothing Then
        ctlS2.TOP = actionTop
        If Not ctlS1 Is Nothing Then
            ctlS2.Left = ctlS1.Left + ctlS1.Width + actionGap
        End If
    End If

    If Not ctlS3 Is Nothing Then
        ctlS3.TOP = actionTop
        If Not ctlS2 Is Nothing Then
            ctlS3.Left = ctlS2.Left + ctlS2.Width + actionGap
        End If
    End If

    If Not ctlOpenCopied Is Nothing Then
        Dim gapTop As Single
        Dim gapHeight As Single

        ctlOpenCopied.Left = Me.btnCopyWithSuffix.Left + mOpenCopiedOffsetLeft

        If Not ctlMax Is Nothing Then
            gapTop = Me.btnCopyWithSuffix.TOP + Me.btnCopyWithSuffix.Height
            gapHeight = ctlMax.TOP - gapTop

            If gapHeight > ctlOpenCopied.Height Then
                ctlOpenCopied.TOP = gapTop + ((gapHeight - ctlOpenCopied.Height) / 2)
            Else
                ctlOpenCopied.TOP = Me.btnCopyWithSuffix.TOP + mOpenCopiedOffsetTop
            End If
        Else
            ctlOpenCopied.TOP = Me.btnCopyWithSuffix.TOP + mOpenCopiedOffsetTop
        End If
    End If

    ' --- ListBox: stretch to fill between top row and bottom block
    Me.lstWorkbooks.Left = mBaseListLeft
    If mIsExpandedView Then
        reservedRight = mRightMargin + mGap + mPanelWidth
    Else
        reservedRight = mRightMargin
    End If
    Me.lstWorkbooks.Width = Me.InsideWidth - Me.lstWorkbooks.Left - reservedRight
    Me.lstWorkbooks.TOP = topBlockBottom + mGap
    ' FullPath bar aligns with listbox width
Me.txtFullPath.Left = Me.lstWorkbooks.Left
Me.txtFullPath.Width = Me.lstWorkbooks.Width

' Place FullPath bar just above the bottom block
Me.txtFullPath.TOP = Me.tglBatchMode.TOP - mGap - Me.txtFullPath.Height

' ListBox ends above FullPath bar (leave gap)
Me.lstWorkbooks.Height = (Me.txtFullPath.TOP - mGap) - Me.lstWorkbooks.TOP

Me.Label3.Left = Me.lstWorkbooks.Left + Me.lstWorkbooks.Width - Me.Label3.Width

If Not mLstSheets Is Nothing Then
    mLstSheets.Left = Me.lstWorkbooks.Left + Me.lstWorkbooks.Width + mGap
    If Not mLblSheetsWorkbook Is Nothing Then
        mLblSheetsWorkbook.Left = mLstSheets.Left
        mLstSheets.TOP = Me.lstWorkbooks.TOP
        mLblSheetsWorkbook.TOP = mLstSheets.TOP - mLblSheetsWorkbook.Height
        mLblSheetsWorkbook.Width = mPanelWidth
        mLblSheetsWorkbook.Visible = mIsExpandedView
        mLblSheetsWorkbook.TextAlign = fmTextAlignLeft
    Else
        mLstSheets.TOP = Me.lstWorkbooks.TOP
    End If
    mLstSheets.Width = mPanelWidth
    sheetBottom = Me.btnOpenFile.TOP + Me.btnOpenFile.Height
    If Not mBtnTogglePanel Is Nothing Then
        If mBtnTogglePanel.TOP - mGap < sheetBottom Then sheetBottom = mBtnTogglePanel.TOP - mGap
    End If
    If sheetBottom < mLstSheets.TOP + 24 Then sheetBottom = mLstSheets.TOP + 24
    mLstSheets.Height = sheetBottom - mLstSheets.TOP
    mLstSheets.Visible = mIsExpandedView
End If


    ' --- ListBox columns: File width based on average display width of file names
    newFileW = ResolveFileColumnWidthPt(Me.lstWorkbooks.Width)
    Me.lstWorkbooks.ColumnWidths = CStr(newFileW) & " pt;30 pt;55 pt;0 pt"
End Sub

Private Function ResolveFileColumnWidthPt(ByVal listWidthPt As Single) As Single
    Const DIR_COL_W As Single = 30
    Const SYNC_COL_W As Single = 55
    Const SCROLLBAR_FUDGE As Single = 18
    Const MIN_FILE_W As Single = 80

    Dim maxFileW As Single
    Dim avgNameW As Single

    maxFileW = listWidthPt - DIR_COL_W - SYNC_COL_W - SCROLLBAR_FUDGE
    If maxFileW < MIN_FILE_W Then
        ResolveFileColumnWidthPt = MIN_FILE_W
        Exit Function
    End If

    ' If there are no file rows, keep the previous behavior (wide default / layout-driven width).
    If Me.lstWorkbooks.ListCount <= 1 Then
        ResolveFileColumnWidthPt = maxFileW
        Exit Function
    End If

    avgNameW = GetAverageFileNameDisplayWidthPt()
    If avgNameW < MIN_FILE_W Then avgNameW = MIN_FILE_W
    If avgNameW > maxFileW Then avgNameW = maxFileW

    ' Always give remaining horizontal space to File so resizing widens File (not Sync).
    If avgNameW < maxFileW Then avgNameW = maxFileW

    ResolveFileColumnWidthPt = avgNameW
End Function

Private Sub AutoGrowFormWidthForFileColumn()
    Const DIR_COL_W As Single = 30
    Const SYNC_COL_W As Single = 55
    Const SCROLLBAR_FUDGE As Single = 18
    Const MIN_FILE_W As Single = 80

    Dim desiredFileW As Single
    Dim currentMaxFileW As Single
    Dim extraNeeded As Single
    Dim frameW As Single
    Dim targetInsideW As Single
    Dim targetFormW As Single

    If Me.lstWorkbooks.ListCount <= 1 Then Exit Sub

    desiredFileW = GetAverageFileNameDisplayWidthPt()
    If desiredFileW < MIN_FILE_W Then desiredFileW = MIN_FILE_W

    currentMaxFileW = Me.lstWorkbooks.Width - DIR_COL_W - SYNC_COL_W - SCROLLBAR_FUDGE
    If desiredFileW <= currentMaxFileW Then Exit Sub

    extraNeeded = desiredFileW - currentMaxFileW
    frameW = Me.Width - Me.InsideWidth
    targetInsideW = Me.InsideWidth + extraNeeded
    targetFormW = targetInsideW + frameW

    If targetFormW > FORM_MAX_W Then targetFormW = FORM_MAX_W
    If targetFormW > Me.Width Then Me.Width = targetFormW
End Sub

Private Function GetAverageFileNameDisplayWidthPt() As Single
    Dim i As Long
    Dim cnt As Long
    Dim nameText As String
    Dim totalWidthTw As Double
    Dim avgWidthPt As Single

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        nameText = CStr(Me.lstWorkbooks.List(i, 0))
        totalWidthTw = totalWidthTw + (MeasureTextWidthPt(nameText) * 20#)
        cnt = cnt + 1
    Next i

    If cnt = 0 Then
        GetAverageFileNameDisplayWidthPt = 240.5
        Exit Function
    End If

    avgWidthPt = CSng((totalWidthTw / cnt) / 20#)
    ' left/right breathing room so average filename is not glued to cell borders
    avgWidthPt = avgWidthPt + 12

    GetAverageFileNameDisplayWidthPt = avgWidthPt
End Function

Private Function MeasureTextWidthPt(ByVal valueText As String) As Single
    On Error GoTo FALLBACK

    If mTextMeasureLabel Is Nothing Then
        Set mTextMeasureLabel = GetControlIfExists("lblMeasureTextHidden")
    End If

    If mTextMeasureLabel Is Nothing Then
        Set mTextMeasureLabel = Me.Controls.Add("Forms.Label.1", "lblMeasureTextHidden", True)
        With mTextMeasureLabel
            .Visible = False
            .AutoSize = True
            .WordWrap = False
            .Left = -10000
            .TOP = -10000
        End With
    End If

    On Error Resume Next
    mTextMeasureLabel.Font.Name = Me.lstWorkbooks.Font.Name
    mTextMeasureLabel.Font.Size = Me.lstWorkbooks.Font.Size
    mTextMeasureLabel.Font.Bold = Me.lstWorkbooks.Font.Bold
    mTextMeasureLabel.Font.Italic = Me.lstWorkbooks.Font.Italic
    On Error GoTo FALLBACK

    mTextMeasureLabel.Caption = valueText
    MeasureTextWidthPt = mTextMeasureLabel.Width
    Exit Function

FALLBACK:
    ' Safe approximation if dynamic label/font sync is unavailable
    MeasureTextWidthPt = Len(valueText) * 5.5
End Function

Private Sub UserForm_Resize()
    On Error Resume Next
    If mMinTrackW > 0 And Me.Width < mMinTrackW Then Me.Width = mMinTrackW
    If mMinTrackH > 0 And Me.Height < mMinTrackH Then Me.Height = mMinTrackH
    If Me.Width > FORM_MAX_W Then Me.Width = FORM_MAX_W
    If Me.Height > FORM_MAX_H Then Me.Height = FORM_MAX_H
    If Not mIsExpandedView Then mCollapsedInsideW = Me.InsideWidth
    ApplyLayout
    On Error GoTo 0
    PositionTopButtons
End Sub

Private Function WaitForRefreshToFinish(ByVal wb As Workbook, ByVal timeoutSec As Long) As Boolean
' We only WAIT for refresh completion.
' We do NOT try to cancel it mid-flight.

    Dim t0 As Double, tNow As Double
    t0 = Timer

    Do
        DoEvents
        Sleep 50

        tNow = Timer
        If tNow < t0 Then tNow = tNow + 86400#

        If (Application.CalculationState <> xlDone) Or AnyWorkbookRefreshing(wb) Then
            If (tNow - t0) >= timeoutSec Then
                WaitForRefreshToFinish = False
                Exit Function
            End If
        Else
            WaitForRefreshToFinish = True
            Exit Function
        End If
    Loop
End Function

Private Function AnyWorkbookRefreshing(ByVal wb As Workbook) As Boolean
    Dim cn As WorkbookConnection
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject

    On Error Resume Next

    For Each cn In wb.Connections
        If Not cn.OLEDBConnection Is Nothing Then
            If cn.OLEDBConnection.Refreshing Then AnyWorkbookRefreshing = True: Exit Function
        End If
        If Not cn.ODBCConnection Is Nothing Then
            If cn.ODBCConnection.Refreshing Then AnyWorkbookRefreshing = True: Exit Function
        End If
    Next cn

    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            If qt.Refreshing Then AnyWorkbookRefreshing = True: Exit Function
        Next qt

        For Each lo In ws.ListObjects
            If Not lo.QueryTable Is Nothing Then
                If lo.QueryTable.Refreshing Then AnyWorkbookRefreshing = True: Exit Function
            End If
        Next lo
    Next ws

    On Error GoTo 0
End Function

Private Function SafeMsgBox(ByVal Prompt As String, _
                            Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                            Optional ByVal Title As String = vbNullString) As VbMsgBoxResult
    On Error Resume Next
    modWinAPI.BringFormToFront Me.Caption
    modWinAPI.SetTopMostState Me.Caption, False
    On Error GoTo 0

    SafeMsgBox = MsgBox(Prompt, Buttons Or vbMsgBoxSetForeground, Title)

    On Error Resume Next
    modWinAPI.BringFormToFront Me.Caption
    modWinAPI.SetTopMostState Me.Caption, True
    On Error GoTo 0
End Function

Private Sub PositionTopButtons()
    Const MARGIN_R As Single = 6
    Const MARGIN_B As Single = 6
    Const GAP As Single = 6
    Dim ctlMax As Object
    Dim ctlS1 As Object
    Dim ctlS2 As Object
    Dim ctlS3 As Object

    ' Close  kotwica: prawy dolny róg
    btnClose.Left = Me.InsideWidth - btnClose.Width - MARGIN_R
    btnClose.TOP = Me.InsideHeight - btnClose.Height - MARGIN_B

    ' Cancel  zawsze obok Close (z lewej)
    btnCancel.TOP = btnClose.TOP
    btnCancel.Left = btnClose.Left - GAP - btnCancel.Width
    If Not mBtnTogglePanel Is Nothing Then
        mBtnTogglePanel.Width = btnClose.Width
        mBtnTogglePanel.Height = btnClose.Height
        mBtnTogglePanel.Left = btnClose.Left
        mBtnTogglePanel.TOP = btnClose.TOP - GAP - mBtnTogglePanel.Height
    End If

    Set ctlMax = GetControlIfExists("btnMaximize")
    Set ctlS1 = GetControlIfExists("btnScreen1")
    Set ctlS2 = GetControlIfExists("btnScreen2")
    Set ctlS3 = GetControlIfExists("btnScreen3")

    If Not ctlMax Is Nothing Then
        ctlMax.TOP = btnClose.TOP
        ctlMax.Left = Me.btnCopyWithSuffix.Left
    End If

    If Not ctlS1 Is Nothing Then
        ctlS1.TOP = btnClose.TOP
        If Not ctlMax Is Nothing Then
            ctlS1.Left = ctlMax.Left + ctlMax.Width + GAP
        End If
    End If

    If Not ctlS2 Is Nothing Then
        ctlS2.TOP = btnClose.TOP
        If Not ctlS1 Is Nothing Then
            ctlS2.Left = ctlS1.Left + ctlS1.Width + GAP
        End If
    End If

    If Not ctlS3 Is Nothing Then
        ctlS3.TOP = btnClose.TOP
        If Not ctlS2 Is Nothing Then
            ctlS3.Left = ctlS2.Left + ctlS2.Width + GAP
        End If
    End If

    EnsureTopLeftButtons
    EnsureSnapshotOverlay
    ApplySnapshotOverlayVisibility
    If mSettingsMode Then EnsureSettingsOverlay

    ' Na wierzch (jesli cos przykrywa)
    btnClose.ZOrder 0
    btnCancel.ZOrder 0
    If Not mBtnTogglePanel Is Nothing Then mBtnTogglePanel.ZOrder 0
End Sub



Private Sub SetBatchUI(ByVal running As Boolean)
    btnCancel.Visible = running
    btnCancel.enabled = True
    btnClose.enabled = Not running
    PositionTopButtons
End Sub
Private Sub ShowFullPathForIndex(ByVal idx As Long)
    Dim rawName As String
    Dim fullP As String

    If idx <= 0 Then
        Me.txtFullPath.Value = ""
        Exit Sub
    End If

    rawName = GetRawNameFromRow(idx)
    fullP = CStr(Me.lstWorkbooks.List(idx, 3)) ' col3 = FullPath (hidden)

    ' Bezpiecznik: jesli z jakiegos powodu col3 puste, pokaz chociaz nazwe
    If Len(fullP) = 0 Then fullP = rawName

    Me.txtFullPath.Value = fullP
End Sub

' =========================================================
' COPY WITH SUFFIX (NO LINK BREAK) + FILE OPENERS
' =========================================================

Private Function CopyWithSuffixOnly(ByVal srcWb As Workbook, ByVal targetFolder As String, ByVal suffix As String, ByRef outPath As String) As Boolean
    On Error GoTo EH

    outPath = BuildOutputPath(targetFolder, srcWb.Name, suffix)
    srcWb.SaveCopyAs outPath

    CopyWithSuffixOnly = True
    Exit Function

EH:
    SafeMsgBox "Copy with suffix error for: " & srcWb.Name & vbCrLf & Err.Description, vbCritical
    CopyWithSuffixOnly = False
End Function

Private Function PickFilesMulti(ByVal titleText As String) As Collection
    Dim fd As FileDialog
    Dim col As Collection
    Dim i As Long
    Dim initialFolder As String

    Set col = New Collection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    initialFolder = modNavigatorSettings.ResolveOpenFilesInitialFolder(ThisWorkbook.Path)

    With fd
        .Title = titleText
        .AllowMultiSelect = True
        On Error Resume Next
        .InitialFileName = initialFolder
        On Error GoTo 0
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xlsb;*.xls", 1
        .Filters.Add "All files", "*.*"

        If .Show <> -1 Then
            Set PickFilesMulti = col
            Exit Function
        End If

        For i = 1 To .SelectedItems.Count
            col.Add CStr(.SelectedItems(i))
        Next i
    End With

    Set PickFilesMulti = col
End Function

Private Function OpenWorkbookSafe(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    On Error GoTo EH

    Set wb = Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=False)
    Set OpenWorkbookSafe = wb
    Exit Function

EH:
    SafeMsgBox "Open file error:" & vbCrLf & filePath & vbCrLf & Err.Description, vbExclamation
    Set OpenWorkbookSafe = Nothing
End Function

