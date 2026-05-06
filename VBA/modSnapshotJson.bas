Attribute VB_Name = "modSnapshotJson"
Option Explicit

Public Function SerializeWorkbookSnapshot(ByVal s As Object) As String
    Dim b As String, k As Variant, first As Boolean
    b = "{""snapshotSchemaVersion"":""" & s("SnapshotSchemaVersion") & ""","
    b = b & """workbookName"":""" & modSnapshotUtils.JsonEscape(s("WorkbookName")) & ""","
    b = b & """workbookPath"":""" & modSnapshotUtils.JsonEscape(s("WorkbookPath")) & ""","
    b = b & """createdAtUtc"":""" & s("CreatedAtUtc") & ""","
    b = b & """createdBy"":""" & modSnapshotUtils.JsonEscape(s("CreatedBy")) & ""","
    b = b & """machineName"":""" & modSnapshotUtils.JsonEscape(s("MachineName")) & ""","
    b = b & """excelVersion"":""" & modSnapshotUtils.JsonEscape(s("ExcelVersion")) & ""","
    b = b & """worksheets"":{"
    first = True
    For Each k In s("Worksheets").Keys
        If Not first Then b = b & ","
        b = b & """" & modSnapshotUtils.JsonEscape(CStr(k)) & """:""" & modSnapshotUtils.JsonEscape(CStr(s("Worksheets")(k))) & """"
        first = False
    Next k
    b = b & "},""cells"":{"
    first = True
    For Each k In s("Cells").Keys
        If Not first Then b = b & ","
        b = b & """" & modSnapshotUtils.JsonEscape(CStr(k)) & """:""" & modSnapshotUtils.JsonEscape(CStr(s("Cells")(k))) & """"
        first = False
    Next k
    b = b & "}}"
    SerializeWorkbookSnapshot = b
End Function

Public Function IsLikelyValidJson(ByVal t As String) As Boolean
    t = Trim$(t)
    IsLikelyValidJson = (Left$(t, 1) = "{" And Right$(t, 1) = "}")
End Function

Public Function LoadWorkbookSnapshot(ByVal path As String) As Object
    Err.Raise vbObjectError + 6301, , "JSON parse for compare is not yet implemented in this MVP scaffolding."
End Function
