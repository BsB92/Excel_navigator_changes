Attribute VB_Name = "modSnapshotCompare"
Option Explicit

Public Function CompareSnapshots(ByVal oldSnap As Object, ByVal newSnap As Object, ByVal options As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("FormulaChanges") = New Collection
    Set d("ValueChanges") = New Collection
    Dim k As Variant, oldItem As String, newItem As String
    For Each k In oldSnap("Cells").Keys
        If newSnap("Cells").Exists(k) Then
            oldItem = CStr(oldSnap("Cells")(k))
            newItem = CStr(newSnap("Cells")(k))
            If oldItem <> newItem Then
                Dim o() As String, n() As String
                o = Split(oldItem, ChrW(31)): n = Split(newItem, ChrW(31))
                If o(0) <> n(0) Then d("FormulaChanges").Add Array(Split(CStr(k), "!")(0), Split(CStr(k), "!")(1), o(0), n(0))
                If o(1) <> n(1) Then d("ValueChanges").Add Array(Split(CStr(k), "!")(0), Split(CStr(k), "!")(1), o(1), n(1))
            End If
        Else
            d("ValueChanges").Add Array(Split(CStr(k), "!")(0), Split(CStr(k), "!")(1), oldSnap("Cells")(k), "<missing>")
        End If
    Next k
    Set CompareSnapshots = d
End Function
