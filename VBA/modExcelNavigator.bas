Attribute VB_Name = "modExcelNavigator"
Option Explicit

Public Sub ExcelNavigatorV46_alpha()
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
    ExcelNavigatorV46_alpha
End Sub

Public Sub ExcelNavigatorV45()
    ExcelNavigatorV46_alpha
End Sub
