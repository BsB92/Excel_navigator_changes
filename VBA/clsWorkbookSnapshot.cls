VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsWorkbookSnapshot"
Option Explicit

Public WorkbookName As String
Public WorkbookPath As String
Public CreatedAtUtc As String
Public CreatedBy As String
Public MachineName As String
Public ExcelVersion As String
Public SnapshotSchemaVersion As String
Public Worksheets As Object
Public Cells As Object
Public NamedRanges As Object
Public ExternalLinks As Object
Public VbaModules As Object

Private Sub Class_Initialize()
    Set Worksheets = CreateObject("Scripting.Dictionary")
    Set Cells = CreateObject("Scripting.Dictionary")
    Set NamedRanges = CreateObject("Scripting.Dictionary")
    Set ExternalLinks = CreateObject("Scripting.Dictionary")
    Set VbaModules = CreateObject("Scripting.Dictionary")
End Sub
