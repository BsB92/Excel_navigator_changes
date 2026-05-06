Attribute VB_Name = "modSnapshotStorage"
Option Explicit

Public Function SaveSnapshotTransactional(ByVal wb As Workbook, ByVal jsonText As String, ByVal userName As String) As String
    Dim finalDir As String, fallbackDir As String, baseName As String, ts As String
    Dim tmpFile As String, finalFile As String
    finalDir = BuildPrimaryFolder(wb)
    If Not EnsureFolderPath(finalDir) Then
        fallbackDir = Environ$("APPDATA") & "\ExcelNavigator\Snapshots\"
        finalDir = fallbackDir & CleanFileName(wb.Name)
        If Not EnsureFolderPath(finalDir) Then Err.Raise vbObjectError + 6201, , "Cannot create snapshot storage directory."
    End If
    ts = Format$(Now, "yyyy-mm-dd_hhnnss")
    baseName = ts & "_" & CleanFileName(LCase$(userName)) & ".json"
    finalFile = finalDir & "\" & baseName
    tmpFile = finalFile & ".tmp"
    WriteTextFile tmpFile, jsonText
    If Not modSnapshotJson.IsLikelyValidJson(ReadTextFile(tmpFile)) Then
        On Error Resume Next: Kill tmpFile: On Error GoTo 0
        Err.Raise vbObjectError + 6202, , "Snapshot JSON validation failed."
    End If
    Name tmpFile As finalFile
    SaveSnapshotTransactional = finalFile
End Function

Private Function BuildPrimaryFolder(ByVal wb As Workbook) As String
    BuildPrimaryFolder = wb.Path & "\.snapshots\" & CleanFileName(Replace(wb.Name, ".xlsx", ""))
End Function
Public Function EnsureFolderPath(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    Dim parts() As String, i As Long, cur As String
    parts = Split(folderPath, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
    EnsureFolderPath = True
    Exit Function
EH:
End Function
Private Sub WriteTextFile(ByVal path As String, ByVal txt As String)
    Dim ff As Integer: ff = FreeFile
    Open path For Output As #ff
    Print #ff, txt
    Close #ff
End Sub
Private Function ReadTextFile(ByVal path As String) As String
    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff
    ReadTextFile = Input$(LOF(ff), ff)
    Close #ff
End Function
Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\\", "/", ":", "*", "?", """", "<", ">", "|", " ")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i
    CleanFileName = s
End Function
