VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCF2StaticButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mButton As MSForms.CommandButton
Attribute mButton.VB_VarHelpID = -1
Private mOwnerForm As Object

Public Sub Attach(ByVal targetButton As MSForms.CommandButton, ByVal ownerForm As Object)
    Set mButton = targetButton
    Set mOwnerForm = ownerForm
End Sub

Private Sub mButton_Click()
    modCF2Static.RunCF2Static mOwnerForm
End Sub
