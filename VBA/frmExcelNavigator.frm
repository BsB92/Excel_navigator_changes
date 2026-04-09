VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExcelNavigator 
   Caption         =   "ExcelNavigator v3.1"
   ClientHeight    =   9240.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
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
Private Const FORM_MIN_W As Long = 420
Private Const FORM_MIN_H As Long = 520
Private Const FORM_MAX_W As Long = 1200
Private Const FORM_MAX_H As Long = 900
Private Const REG_APP As String = "ExcelNavigator"
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

Private mTopBlockBottom As Single   ' bottom of search row
Private mBottomBlockTop As Single   ' top of bottom buttons row (before shifting)

Private mRightMargin As Single
Private mGap As Single

Private mCtlTop(1 To 20) As Single
Private mHookReady As Boolean


' ========= CONSTANTS =========
Private Const ACTIVE_PREFIX As String = "> "
Private Const CLOSEMODE_ASK_EACH As Long = 0
Private Const CLOSEMODE_SAVE_ALL As Long = 1
Private Const CLOSEMODE_DONT_SAVE_ALL As Long = 2

Private Sub btnCancel_Click()
    ' Cancel DOES NOT stop an active RefreshAll.
    ' It only prevents starting refresh/save for NEXT workbooks.
    mCancelBatch = True
End Sub
Public Sub RestoreNavigatorToFront()
    On Error Resume Next
    If Not Me.Visible Then Me.Show vbModeless
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

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then

            wbName = GetRawNameFromRow(i)
            Set srcWb = GetWorkbookByName(wbName)

            If srcWb Is Nothing Then
                SafeMsgBox "Workbook not found: " & wbName, vbExclamation
            ElseIf Len(srcWb.Path) = 0 Then
                SafeMsgBox "Workbook must be saved first: " & srcWb.Name, vbExclamation
            Else
                If CopyAndBreakLinks(srcWb, targetFolder, suffix) Then
                    copiedCount = copiedCount + 1
                End If
            End If

        End If
    Next i

    If copiedCount > 0 Then
        mLastCopyFolder = targetFolder
        Me.btnOpenCopyFolder.enabled = True
    End If

    SafeMsgBox "Done. Copied and processed: " & copiedCount & " file(s).", vbInformation
End Sub

Private Function PickFolder(ByVal titleText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = titleText
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function

Private Function CopyAndBreakLinks(ByVal srcWb As Workbook, ByVal targetFolder As String, ByVal suffix As String) As Boolean
    Dim outPath As String
    Dim copiedWb As Workbook

    On Error GoTo EH

    outPath = BuildOutputPath(targetFolder, srcWb.Name, suffix)

    srcWb.SaveCopyAs outPath
    Set copiedWb = Workbooks.Open(fileName:=outPath, UpdateLinks:=0, ReadOnly:=False)

    BreakExternalLinks copiedWb

    copiedWb.Save
    copiedWb.Close saveChanges:=False

    CopyAndBreakLinks = True
    Exit Function

EH:
    SafeMsgBox "Copy/BreakLinks error for: " & srcWb.Name & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If Not copiedWb Is Nothing Then copiedWb.Close saveChanges:=False
    CopyAndBreakLinks = False
End Function

Private Function BuildOutputPath(ByVal folderPath As String, ByVal fileName As String, ByVal suffix As String) As String
    Dim baseName As String, ext As String
    Dim dotPos As Long

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        ext = ".xlsx"
    End If

    BuildOutputPath = folderPath & Application.PathSeparator & baseName & suffix & ext
End Function

Private Sub BreakExternalLinks(ByVal wb As Workbook)
    Dim links As Variant
    Dim i As Long

    links = wb.LinkSources(Type:=xlExcelLinks)
    If IsEmpty(links) Then Exit Sub

    For i = LBound(links) To UBound(links)
        wb.BreakLink Name:=CStr(links(i)), Type:=xlLinkTypeExcelLinks
    Next i
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

    For i = 1 To Me.lstWorkbooks.ListCount - 1
        If Me.lstWorkbooks.Selected(i) Then
            wbName = GetRawNameFromRow(i)
            Set srcWb = GetWorkbookByName(wbName)

            If srcWb Is Nothing Then
                SafeMsgBox "Workbook not found: " & wbName, vbExclamation
            ElseIf Len(srcWb.Path) = 0 Then
                SafeMsgBox "Workbook must be saved first: " & srcWb.Name, vbExclamation
            Else
                If CopyWithSuffixOnly(srcWb, targetFolder, suffix) Then
                    copiedCount = copiedCount + 1
                End If
            End If
        End If
    Next i

    If copiedCount > 0 Then
        mLastCopyFolder = targetFolder
        Me.btnOpenCopyFolder.enabled = True
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

    If TryHookResize() Then
        mHookReady = True
    End If

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
    Me.lstWorkbooks.ColumnWidths = "185 pt;30 pt;55 pt;0 pt"
    Me.lstWorkbooks.MultiSelect = fmMultiSelectSingle

    Me.txtSearch.Value = ""

    Me.tglBatchMode.Value = False
    Me.tglBatchMode.Caption = "Selection mode: OFF"

    SetActionButtonsEnabled False
    Me.btnMaximize.enabled = True
    Me.btnScreen1.enabled = True
    Me.btnScreen2.enabled = True
    Me.btnScreen3.enabled = True
    mLastCopyFolder = ""
    Me.btnOpenCopyFolder.enabled = False
    Me.txtSuffix.Value = "_without_formulas"


    ReloadListPreserveSelection
    RefreshVisuals

    
  
Dim w As Variant, h As Variant

w = GetSetting(REG_APP, REG_SEC, "W", 0)
h = GetSetting(REG_APP, REG_SEC, "H", 0)

If w > 0 And h > 0 Then
    Me.Width = w
    Me.Height = h
End If
btnCancel.Visible = False
btnCancel.enabled = True
PositionTopButtons

' --- layout only (resize hook will be done in Activate when hwnd exists) ---
mHookReady = False
CacheLayout
ApplyLayout



End Sub

' =========================================================
' SELECTION MODE
' =========================================================
Private Sub tglBatchMode_Click()
    If Me.tglBatchMode.Value Then
        Me.tglBatchMode.Caption = "Selection mode: ON"
        Me.lstWorkbooks.MultiSelect = fmMultiSelectMulti
        ClearAllSelections
        Me.lstWorkbooks.ListIndex = -1
        ShowFullPathForIndex -1
        SetActionButtonsEnabled True
    Else
        Me.tglBatchMode.Caption = "Selection mode: OFF"
        Me.lstWorkbooks.MultiSelect = fmMultiSelectSingle
        ClearAllSelections
        Me.lstWorkbooks.ListIndex = -1
        ShowFullPathForIndex -1
        SetActionButtonsEnabled False
    End If

    Me.tglBatchMode.BackColor = IIf(Me.tglBatchMode.Value, RGB(0, 176, 80), vbButtonFace)
    RefreshVisuals
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
    Me.btnMaximize.enabled = True
    Me.btnScreen1.enabled = True
    Me.btnScreen2.enabled = True
    Me.btnScreen3.enabled = True


End Sub

' =========================================================
' LIST EVENTS (header block)
' =========================================================
Private Sub lstWorkbooks_Click()

    Dim idx As Long
    Dim wbName As String
    Dim wb As Workbook

    If mUIBusy Then Exit Sub

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

    ' Selection mode: tylko zaznaczanie, bez aktywacji
    If Me.tglBatchMode.Value Then
        RefreshVisuals
        Exit Sub
    End If

    wbName = GetRawNameFromRow(idx)
    Set wb = GetWorkbookByName(wbName)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If wb.Windows.Count > 0 Then wb.Windows(1).Activate
    wb.Activate
    On Error GoTo 0

    RefreshVisuals
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
End Sub


Private Sub lstWorkbooks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    FixHeaderSelection
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
    Dim moved As Long
    Dim info As String

    On Error GoTo EH

    Set targets = GetWorkbooksForWindowAction()
    If targets.Count = 0 Then Exit Sub

    If targetScreen > 0 Then
        If Not modWinAPI.TryGetMonitorWorkArea(targetScreen, workLeft, workTop, workWidth, workHeight) Then
            SafeMsgBox "Screen " & CStr(targetScreen) & " not available.", vbExclamation
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
    Exit Sub

EH:
    modWinAPI.ResumeExcelWindowUpdates
    If screenUpdatingCaptured Then
        Application.ScreenUpdating = prevScreenUpdating
    Else
        Application.ScreenUpdating = True
    End If
    For Each item In targets
        Set wb = item
        If targetScreen <= 0 Then
            moved = moved + MaximizeWorkbookWindows(wb)
        Else
            moved = moved + MoveWorkbookToScreenAndMaximize(wb, targetScreen)
        End If
    Next item

    If targetScreen <= 0 Then
        info = "Maximized windows: " & CStr(moved)
    Else
        info = "Moved+maximized windows on screen " & CStr(targetScreen) & ": " & CStr(moved)
    End If
    SafeMsgBox info, vbInformation
    Exit Sub

EH:
    SafeMsgBox "Window action error: " & Err.Description, vbCritical
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

Private Function MoveWorkbookToScreenAndMaximize(ByVal wb As Workbook, ByVal targetScreen As Long) As Long
    Dim workLeft As Long, workTop As Long
    Dim workWidth As Long, workHeight As Long
    Dim win As Window

    If Not modWinAPI.TryGetMonitorWorkArea(targetScreen, workLeft, workTop, workWidth, workHeight) Then
        SafeMsgBox "Screen " & CStr(targetScreen) & " not available.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    For Each win In wb.Windows
        win.WindowState = xlNormal
        win.Left = workLeft
        win.TOP = workTop
        win.Width = workWidth
        win.Height = workHeight
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

Private Sub UserForm_Terminate()
    On Error Resume Next
    SaveSetting REG_APP, REG_SEC, "W", Me.Width
    SaveSetting REG_APP, REG_SEC, "H", Me.Height
End Sub


Private Sub CacheLayout()
    ' Cache once after size is restored
    mBaseInsideW = Me.InsideWidth
    mBaseInsideH = Me.InsideHeight

    mRightMargin = 8
    mGap = 8

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
    mCtlTop(17) = Me.btnMaximize.TOP
    mCtlTop(18) = Me.btnScreen1.TOP
    mCtlTop(19) = Me.btnScreen2.TOP
    mCtlTop(20) = Me.btnScreen3.TOP

    
    mBottomBlockTop = Me.tglBatchMode.TOP
    mTopBlockBottom = Me.txtSearch.TOP + Me.txtSearch.Height

    ' Start file column width from current setting
    mBaseFileColW = 185
End Sub

Private Sub ApplyLayout()
    Dim deltaH As Single
    Dim newFileW As Single
    Dim rightX As Single
    Dim topBlockBottom As Single

    If mBaseInsideW = 0 Or mBaseInsideH = 0 Then Exit Sub

    deltaH = Me.InsideHeight - mBaseInsideH

    ' --- Top row: Search stretches, Reload stays right
    rightX = Me.InsideWidth - mRightMargin
    Me.btnReload.Left = rightX - Me.btnReload.Width
    Me.txtSearch.Width = (Me.btnReload.Left - mGap) - Me.txtSearch.Left

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
    Me.btnMaximize.TOP = mCtlTop(17) + deltaH
    Me.btnScreen1.TOP = mCtlTop(18) + deltaH
    Me.btnScreen2.TOP = mCtlTop(19) + deltaH
    Me.btnScreen3.TOP = mCtlTop(20) + deltaH



    ' Keep Close button on the right edge
    Me.btnClose.Left = rightX - Me.btnClose.Width

    ' --- ListBox: stretch to fill between top row and bottom block
    Me.lstWorkbooks.Left = Me.Label1.Left
    Me.lstWorkbooks.Width = Me.InsideWidth - Me.lstWorkbooks.Left - mRightMargin
    Me.lstWorkbooks.TOP = topBlockBottom + mGap
    ' FullPath bar aligns with listbox width
Me.txtFullPath.Left = Me.lstWorkbooks.Left
Me.txtFullPath.Width = Me.lstWorkbooks.Width

' Place FullPath bar just above the bottom block
Me.txtFullPath.TOP = Me.tglBatchMode.TOP - mGap - Me.txtFullPath.Height

' ListBox ends above FullPath bar (leave gap)
Me.lstWorkbooks.Height = (Me.txtFullPath.TOP - mGap) - Me.lstWorkbooks.TOP


    ' --- ListBox columns: File grows
    newFileW = Me.lstWorkbooks.Width - 30 - 55 - 18 ' Dir(30) + Sync(55) + scrollbar fudge
    If newFileW < 80 Then newFileW = 80
    Me.lstWorkbooks.ColumnWidths = CStr(newFileW) & " pt;30 pt;55 pt;0 pt"
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
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
    modWinAPI.SetTopMostState Me.Caption, False
    On Error GoTo 0

    SafeMsgBox = MsgBox(Prompt, Buttons, Title)

    On Error Resume Next
    modWinAPI.SetTopMostState Me.Caption, True
    On Error GoTo 0
End Function

Private Sub PositionTopButtons()
    Const MARGIN_R As Single = 10
    Const MARGIN_B As Single = 10
    Const GAP As Single = 6

    ' Close  kotwica: prawy dolny róg
    btnClose.Left = Me.InsideWidth - btnClose.Width - MARGIN_R
    btnClose.TOP = Me.InsideHeight - btnClose.Height - MARGIN_B

    ' Cancel  zawsze obok Close (z lewej)
    btnCancel.TOP = btnClose.TOP
    btnCancel.Left = btnClose.Left - GAP - btnCancel.Width

    ' Na wierzch (jesli cos przykrywa)
    btnClose.ZOrder 0
    btnCancel.ZOrder 0
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

Private Function CopyWithSuffixOnly(ByVal srcWb As Workbook, ByVal targetFolder As String, ByVal suffix As String) As Boolean
    Dim outPath As String
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

    Set col = New Collection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = titleText
        .AllowMultiSelect = True
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

