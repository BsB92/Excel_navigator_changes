Attribute VB_Name = "modExcelNavigator"
Option Explicit

Public Const EXCEL_NAVIGATOR_VERSION As String = "v6.1"

Private Sub ShowExcelNavigatorForm()
    Dim uf As Object
    Dim launchStage As String

    On Error GoTo LaunchError

    launchStage = "checking loaded forms"
    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmExcelNavigator" Then
            launchStage = "showing the loaded Excel Navigator form"
            uf.Caption = "ExcelNavigator " & EXCEL_NAVIGATOR_VERSION
            uf.Show vbModeless
            Exit Sub
        End If
    Next uf

    launchStage = "initializing the Excel Navigator form"
    frmExcelNavigator.Caption = "ExcelNavigator " & EXCEL_NAVIGATOR_VERSION
    launchStage = "showing the Excel Navigator form"
    frmExcelNavigator.Show vbModeless
    Exit Sub

LaunchError:
    MsgBox "Excel Navigator could not start while " & launchStage & "." & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, _
           vbCritical, "ExcelNavigator " & EXCEL_NAVIGATOR_VERSION
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
