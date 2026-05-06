VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsSnapshotDiffResult"
Option Explicit

Public FormulaChanges As Collection
Public ValueChanges As Collection
Public SheetChanges As Collection
Public NameChanges As Collection
Public VbaChanges As Collection
Public LinkChanges As Collection

Private Sub Class_Initialize()
    Set FormulaChanges = New Collection
    Set ValueChanges = New Collection
    Set SheetChanges = New Collection
    Set NameChanges = New Collection
    Set VbaChanges = New Collection
    Set LinkChanges = New Collection
End Sub
