Attribute VB_Name = "modExcelNavigator"
Option Explicit

Private Sub ShowExcelNavigatorForm()
    Dim uf As Object

    On Error Resume Next

    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmExcelNavigator" Then
            uf.Show vbModeless
            Exit Sub
        End If
    Next uf

    frmExcelNavigator.Show vbModeless
End Sub

Public Sub ExcelNavigator()
    ShowExcelNavigatorForm
End Sub
