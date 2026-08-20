Attribute VB_Name = "modCF2StaticUI"
Option Explicit

Public Sub InitializeCF2StaticButton(ByVal ownerForm As Object)
    If ownerForm Is Nothing Then Exit Sub

    On Error Resume Next
    ownerForm.Caption = "ExcelNavigator " & modExcelNavigator.EXCEL_NAVIGATOR_VERSION
    ownerForm.InitializeCF2StaticButton
    On Error GoTo 0
End Sub
