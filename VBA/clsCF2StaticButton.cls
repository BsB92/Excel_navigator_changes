Option Explicit

Private WithEvents mButton As MSForms.CommandButton
Private mOwnerForm As Object

Public Sub Attach(ByVal targetButton As MSForms.CommandButton, ByVal ownerForm As Object)
    Set mButton = targetButton
    Set mOwnerForm = ownerForm
End Sub

Private Sub mButton_Click()
    modCF2Static.RunCF2Static mOwnerForm
End Sub
