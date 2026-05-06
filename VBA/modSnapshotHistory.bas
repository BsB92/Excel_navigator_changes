Attribute VB_Name = "modSnapshotHistory"
Option Explicit
Public Sub SnapshotHistoryManager()
    On Error Resume Next
    frmExcelNavigator.RestoreNavigatorToFront
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, False
    On Error GoTo 0
    MsgBox "Snapshot History UI to be added as UserForm in next iteration.", vbInformation Or vbMsgBoxSetForeground, "ExcelNavigator Snapshot"
    On Error Resume Next
    modWinAPI.SetTopMostState frmExcelNavigator.Caption, True
    On Error GoTo 0
End Sub
