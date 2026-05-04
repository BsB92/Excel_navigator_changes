Attribute VB_Name = "modExcelNavigator"
Option Explicit

Public Sub ExcelNavigator_v_4_7_alpha()
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
    ExcelNavigator_v_4_7_alpha
End Sub

Public Sub ExcelNavigatorV45()
    ExcelNavigator_v_4_7_alpha
End Sub

Public Sub ExcelNavigatorV46_alpha()
    ExcelNavigator_v_4_7_alpha
End Sub
