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

Public Sub ExcelNavigator_v5_2()
    ShowExcelNavigatorForm
End Sub



Public Sub ExcelNavigatorCreateSnapshot()
    modSnapshotMain.SnapshotCreateActiveWorkbook
End Sub

Public Sub ExcelNavigatorCompareSnapshot()
    modSnapshotMain.SnapshotCompareActiveWorkbook
End Sub

Public Sub ExcelNavigatorSnapshotHistory()
    modSnapshotHistory.SnapshotHistoryManager
End Sub
