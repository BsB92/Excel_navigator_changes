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
    Dim i As Long
    Dim j As Long
    Dim ch As Integer

    If Len(t) = 0 Then Exit Function

    i = 1
    Do While i <= Len(t)
        ch = AscW(Mid$(t, i, 1))
        If ch > 32 Then Exit Do
        i = i + 1
    Loop

    j = Len(t)
    Do While j >= i
        ch = AscW(Mid$(t, j, 1))
        If ch > 32 Then Exit Do
        j = j - 1
    Loop

    If j < i Then Exit Function

    IsLikelyValidJson = (Mid$(t, i, 1) = "{" And Mid$(t, j, 1) = "}")
End Function

Public Function LoadWorkbookSnapshot(ByVal path As String) As Object
    Dim t As String
    Dim s As Object
    Dim pWs As Long, pCells As Long

    t = ReadAllText(path)
    If Not IsLikelyValidJson(t) Then
        Err.Raise vbObjectError + 6302, , "Invalid snapshot JSON."
    End If

    Set s = CreateObject("Scripting.Dictionary")
    s("SnapshotSchemaVersion") = ExtractJsonString(t, "snapshotSchemaVersion")
    s("WorkbookName") = ExtractJsonString(t, "workbookName")
    s("WorkbookPath") = ExtractJsonString(t, "workbookPath")
    s("CreatedAtUtc") = ExtractJsonString(t, "createdAtUtc")
    s("CreatedBy") = ExtractJsonString(t, "createdBy")
    s("MachineName") = ExtractJsonString(t, "machineName")
    s("ExcelVersion") = ExtractJsonString(t, "excelVersion")

    pWs = InStr(1, t, """worksheets"":{", vbTextCompare)
    pCells = InStr(1, t, """cells"":{", vbTextCompare)
    If pWs = 0 Or pCells = 0 Or pCells <= pWs Then
        Err.Raise vbObjectError + 6303, , "Snapshot JSON does not contain expected worksheets/cells sections."
    End If

    Set s("Worksheets") = ParseJsonFlatObject(Mid$(t, pWs + Len("""worksheets"":{"), pCells - (pWs + Len("""worksheets"":{")) - 2))
    Set s("Cells") = ParseJsonFlatObject(ExtractObjectBodyAfterKey(t, "cells"))

    If Not s.Exists("NamedRanges") Then Set s("NamedRanges") = CreateObject("Scripting.Dictionary")
    If Not s.Exists("ExternalLinks") Then Set s("ExternalLinks") = CreateObject("Scripting.Dictionary")
    If Not s.Exists("VbaModules") Then Set s("VbaModules") = CreateObject("Scripting.Dictionary")

    Set LoadWorkbookSnapshot = s
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim ff As Integer
    ff = FreeFile
    Open path For Input As #ff
    ReadAllText = Input$(LOF(ff), ff)
    Close #ff
End Function

Private Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q1 As Long, q2 As Long
    Dim token As String
    token = """" & key & """:"""
    p = InStr(1, json, token, vbTextCompare)
    If p = 0 Then Exit Function
    q1 = p + Len(token)
    q2 = FindClosingQuote(json, q1)
    If q2 = 0 Then Exit Function
    ExtractJsonString = JsonUnescape(Mid$(json, q1, q2 - q1))
End Function

Private Function FindClosingQuote(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long
    For i = startPos To Len(s)
        If Mid$(s, i, 1) = """" Then
            If i = startPos Or Mid$(s, i - 1, 1) <> "\" Then
                FindClosingQuote = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ParseJsonFlatObject(ByVal body As String) As Object
    Dim d As Object
    Dim i As Long
    Dim k As String, v As String

    Set d = CreateObject("Scripting.Dictionary")
    body = Trim$(body)
    If Len(body) = 0 Then
        Set ParseJsonFlatObject = d
        Exit Function
    End If

    i = 1
    Do While i <= Len(body)
        k = ReadJsonToken(body, i)
        If Len(k) = 0 Then Exit Do
        SkipUntil body, i, ":"
        If Mid$(body, i, 1) = ":" Then i = i + 1
        v = ReadJsonToken(body, i)
        d(JsonUnescape(k)) = JsonUnescape(v)
        SkipUntil body, i, ","
        If i <= Len(body) And Mid$(body, i, 1) = "," Then i = i + 1
    Loop

    Set ParseJsonFlatObject = d
End Function

Private Function ReadJsonToken(ByVal s As String, ByRef i As Long) As String
    Dim q1 As Long, q2 As Long
    SkipWhitespace s, i
    q1 = InStr(i, s, """")
    If q1 = 0 Then Exit Function
    q2 = FindClosingQuote(s, q1 + 1)
    If q2 = 0 Then Exit Function
    ReadJsonToken = Mid$(s, q1 + 1, q2 - q1 - 1)
    i = q2 + 1
End Function

Private Sub SkipWhitespace(ByVal s As String, ByRef i As Long)
    Do While i <= Len(s) And AscW(Mid$(s, i, 1)) <= 32
        i = i + 1
    Loop
End Sub

Private Sub SkipUntil(ByVal s As String, ByRef i As Long, ByVal ch As String)
    Do While i <= Len(s) And Mid$(s, i, 1) <> ch
        i = i + 1
    Loop
End Sub

Private Function JsonUnescape(ByVal s As String) As String
    s = Replace$(s, "\n", vbLf)
    s = Replace$(s, "\"" , """")
    s = Replace$(s, "\\", "\")
    JsonUnescape = s
End Function

Private Function ExtractObjectBodyAfterKey(ByVal json As String, ByVal key As String) As String
    Dim p As Long, i As Long, depth As Long
    Dim token As String
    token = """" & key & """:{"
    p = InStr(1, json, token, vbTextCompare)
    If p = 0 Then Exit Function
    i = p + Len(token)
    depth = 1
    Do While i <= Len(json)
        Select Case Mid$(json, i, 1)
            Case "{": depth = depth + 1
            Case "}": depth = depth - 1: If depth = 0 Then Exit Do
        End Select
        i = i + 1
    Loop
    If i > p + Len(token) Then
        ExtractObjectBodyAfterKey = Mid$(json, p + Len(token), i - (p + Len(token)))
    End If
End Function
