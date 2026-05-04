VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigatorSettings 
   Caption         =   "Navigator settings"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5940
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNavigatorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTxtCopy As MSForms.TextBox
Attribute mTxtCopy.VB_VarHelpID = -1
Private WithEvents mTxtOpen As MSForms.TextBox
Attribute mTxtOpen.VB_VarHelpID = -1
Private WithEvents mBtnSave As MSForms.CommandButton
Attribute mBtnSave.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCopy", True)
    lbl.Caption = "Default folder for copy operations"
    lbl.Left = 12: lbl.Top = 12: lbl.Width = 300: lbl.Height = 16

    Set mTxtCopy = Me.Controls.Add("Forms.TextBox.1", "txtCopy", True)
    mTxtCopy.Left = 12: mTxtCopy.Top = 30: mTxtCopy.Width = 570: mTxtCopy.Height = 20

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblOpen", True)
    lbl.Caption = "Default folder for Open/Open&Refresh"
    lbl.Left = 12: lbl.Top = 64: lbl.Width = 300: lbl.Height = 16

    Set mTxtOpen = Me.Controls.Add("Forms.TextBox.1", "txtOpen", True)
    mTxtOpen.Left = 12: mTxtOpen.Top = 82: mTxtOpen.Width = 570: mTxtOpen.Height = 20

    Set mBtnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    mBtnSave.Caption = "Save"
    mBtnSave.Width = 72: mBtnSave.Height = 24: mBtnSave.Left = 430: mBtnSave.Top = 116

    Set mBtnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    mBtnCancel.Caption = "Cancel"
    mBtnCancel.Width = 72: mBtnCancel.Height = 24: mBtnCancel.Left = 510: mBtnCancel.Top = 116

    mTxtCopy.Text = modNavigatorSettings.GetDefaultWorkingFolder()
    mTxtOpen.Text = modNavigatorSettings.GetOpenFilesFolder()

    Me.Width = 6060
    Me.Height = 1960
End Sub

Private Sub mBtnSave_Click()
    Dim copyFolder As String
    Dim openFolder As String

    copyFolder = Trim$(mTxtCopy.Text)
    openFolder = Trim$(mTxtOpen.Text)

    If Len(copyFolder) = 0 Or Len(Dir$(copyFolder, vbDirectory)) = 0 Then
        MsgBox "Copy folder does not exist: " & copyFolder, vbExclamation, "Navigator settings"
        Exit Sub
    End If

    If Len(openFolder) = 0 Or Len(Dir$(openFolder, vbDirectory)) = 0 Then
        MsgBox "Open folder does not exist: " & openFolder, vbExclamation, "Navigator settings"
        Exit Sub
    End If

    modNavigatorSettings.SaveDefaultWorkingFolder copyFolder
    modNavigatorSettings.SaveOpenFilesFolder openFolder
    Unload Me
End Sub

Private Sub mBtnCancel_Click()
    Unload Me
End Sub
