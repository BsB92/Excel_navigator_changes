Attribute VB_Name = "modSnapshotUtils"
Option Explicit

Public Function GetRealUsedRange(ByVal ws As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then Exit Function
    lastRow = f.Row
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then Exit Function
    lastCol = f.Column
    Set GetRealUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Public Function IsVolatileFormula(ByVal formulaText As String) As Boolean
    Dim f As String
    f = UCase$(formulaText)
    IsVolatileFormula = (InStr(f, "NOW(") > 0) Or (InStr(f, "TODAY(") > 0) Or (InStr(f, "RAND(") > 0) Or (InStr(f, "RANDBETWEEN(") > 0) Or (InStr(f, "OFFSET(") > 0) Or (InStr(f, "INDIRECT(") > 0) Or (InStr(f, "CELL(") > 0) Or (InStr(f, "INFO(") > 0)
End Function

Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\"")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function
