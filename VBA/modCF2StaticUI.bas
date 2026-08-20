Attribute VB_Name = "modCF2StaticUI"
Option Explicit

Private mCF2StaticHandler As clsCF2StaticButton

Public Sub InitializeCF2StaticButton(ByVal ownerForm As Object)
    Const CONTROL_NAME As String = "CF2Static"
    Const CONTROL_CAPTION As String = "CF2Static"
    Const CONTROL_GAP As Single = 6
    Const TOOLTIP_TEXT As String = "Creates copies of selected files and converts conditional formatting to static formatting in all worksheets while preserving the exact visible appearance."
    Dim copyFrame As Object
    Dim cfButton As MSForms.CommandButton
    Dim copyBreakButton As Object
    Dim openCopiedCheckBox As Object

    If ownerForm Is Nothing Then Exit Sub

    On Error Resume Next
    ownerForm.Caption = "ExcelNavigator " & modExcelNavigator.EXCEL_NAVIGATOR_VERSION
    Set copyFrame = ownerForm.Controls("fraCopy")
    On Error GoTo 0
    If copyFrame Is Nothing Then Exit Sub

    On Error Resume Next
    Set cfButton = copyFrame.Controls(CONTROL_NAME)
    Set copyBreakButton = copyFrame.Controls("btnFrameCopyBreakLinks")
    Set openCopiedCheckBox = copyFrame.Controls("chkFrameOpenCopiedFiles")
    On Error GoTo 0

    If copyBreakButton Is Nothing Or openCopiedCheckBox Is Nothing Then Exit Sub

    If cfButton Is Nothing Then
        Set cfButton = copyFrame.Controls.Add("Forms.CommandButton.1", CONTROL_NAME, True)
    End If

    With cfButton
        .Caption = CONTROL_CAPTION
        .ControlTipText = TOOLTIP_TEXT
        .Left = copyBreakButton.Left
        .TOP = openCopiedCheckBox.TOP + openCopiedCheckBox.Height + CONTROL_GAP
        .Width = copyBreakButton.Width
        .Height = copyBreakButton.Height
        .Font.Name = copyBreakButton.Font.Name
        .Font.Size = copyBreakButton.Font.Size
        .Font.Bold = copyBreakButton.Font.Bold
        .Font.Italic = copyBreakButton.Font.Italic
        .SpecialEffect = copyBreakButton.SpecialEffect
        .Visible = True
        .enabled = True
    End With

    Set mCF2StaticHandler = New clsCF2StaticButton
    mCF2StaticHandler.Attach cfButton, ownerForm
End Sub
