Attribute VB_Name = "modSnapshotHistory"
Option Explicit
Public Sub SnapshotHistoryManager()
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    On Error GoTo 0
    MsgBox "Snapshot History UI to be added as UserForm in next iteration.", vbInformation, "ExcelNavigator Snapshot"
End Sub
