Attribute VB_Name = "modSnapshotNavigation"
Option Explicit
Public Sub NavigateToCell(ByVal wb As Workbook, ByVal wsName As String, ByVal addr As String)
    On Error Resume Next
    wb.Worksheets(wsName).Activate
    wb.Worksheets(wsName).Range(addr).Select
End Sub
