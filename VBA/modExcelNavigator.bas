Attribute VB_Name = "modExcelNavigator"
Option Explicit

Public Const EXCEL_NAVIGATOR_VERSION As String = "v6.1"

Private Sub ShowExcelNavigatorForm()
    Dim uf As Object

    On Error Resume Next

    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmExcelNavigator" Then
            uf.Caption = "ExcelNavigator " & EXCEL_NAVIGATOR_VERSION
            uf.Show vbModeless
            modCF2StaticUI.InitializeCF2StaticButton uf
            Exit Sub
        End If
    Next uf

    frmExcelNavigator.Caption = "ExcelNavigator " & EXCEL_NAVIGATOR_VERSION
    frmExcelNavigator.Show vbModeless
    modCF2StaticUI.InitializeCF2StaticButton frmExcelNavigator
End Sub

Public Sub ExcelNavigator_v6_1()
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
