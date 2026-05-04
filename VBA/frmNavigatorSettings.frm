VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigatorSettings
   Caption         =   "Excel Navigator Settings"
   ClientHeight    =   2100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblTitle
      Caption         =   "Ustawienia dodatku"
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   2200
   End
   Begin VB.Label lblDesc
      Caption         =   "Domyślna lokalizacja dla Open / Open&Refresh / Copy i podobnych operacji:"
      Height          =   375
      Left            =   120
      Top             =   420
      Width           =   6100
   End
   Begin VB.Label lblFolder
      Caption         =   "Folder:"
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtDefaultFolder
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton btnBrowse
      Caption         =   "Browse..."
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton btnSave
      Caption         =   "Save settings"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton btnClose
      Caption         =   "Close"
      Height          =   345
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmNavigatorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
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
