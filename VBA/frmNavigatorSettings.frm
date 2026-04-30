VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigatorSettings
   Caption         =   "Excel Navigator Settings"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNavigatorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    BuildLayout
    txtDefaultFolder.Value = modNavigatorSettings.GetDefaultWorkingFolder()
End Sub

Private Sub btnBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Wybierz domyślny folder"
        .AllowMultiSelect = False
        On Error Resume Next
        .InitialFileName = txtDefaultFolder.Value
        On Error GoTo 0
        If .Show = -1 Then txtDefaultFolder.Value = .SelectedItems(1)
    End With
End Sub

Private Sub btnSave_Click()
    Dim folderPath As String
    folderPath = Trim$(txtDefaultFolder.Value)

    If Len(folderPath) = 0 Then
        MsgBox "Podaj folder.", vbExclamation
        Exit Sub
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MsgBox "Folder nie istnieje.", vbExclamation
        Exit Sub
    End If

    modNavigatorSettings.SaveDefaultWorkingFolder folderPath
    MsgBox "Ustawienia zapisane.", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub BuildLayout()
    Dim lblTitle As MSForms.Label
    Dim lblDesc As MSForms.Label
    Dim lblFolder As MSForms.Label
    Dim txt As MSForms.TextBox
    Dim btnBrowse As MSForms.CommandButton
    Dim btnSave As MSForms.CommandButton
    Dim btnClose As MSForms.CommandButton

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    lblTitle.Caption = "Ustawienia dodatku"
    lblTitle.Font.Bold = True
    lblTitle.Left = 12
    lblTitle.Top = 12
    lblTitle.Width = 220

    Set lblDesc = Me.Controls.Add("Forms.Label.1", "lblDesc", True)
    lblDesc.Caption = "Domyślna lokalizacja dla Open / Open&Refresh / Copy i podobnych operacji:"
    lblDesc.Left = 12
    lblDesc.Top = 36
    lblDesc.Width = 520
    lblDesc.Height = 24

    Set lblFolder = Me.Controls.Add("Forms.Label.1", "lblFolder", True)
    lblFolder.Caption = "Folder:"
    lblFolder.Left = 12
    lblFolder.Top = 72
    lblFolder.Width = 60

    Set txt = Me.Controls.Add("Forms.TextBox.1", "txtDefaultFolder", True)
    txt.Left = 12
    txt.Top = 90
    txt.Width = 450

    Set btnBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnBrowse", True)
    btnBrowse.Caption = "Browse..."
    btnBrowse.Left = 468
    btnBrowse.Top = 88
    btnBrowse.Width = 78

    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    btnSave.Caption = "Save settings"
    btnSave.Left = 354
    btnSave.Top = 126
    btnSave.Width = 94

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    btnClose.Caption = "Close"
    btnClose.Left = 452
    btnClose.Top = 126
    btnClose.Width = 94
End Sub
