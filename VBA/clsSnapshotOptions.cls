VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsSnapshotOptions"
Option Explicit

Public IgnoreBlankChanges As Boolean
Public IgnoreFormulaResultChanges As Boolean
Public IgnoreVolatileFormulas As Boolean

Private Sub Class_Initialize()
    IgnoreBlankChanges = True
    IgnoreFormulaResultChanges = False
    IgnoreVolatileFormulas = False
End Sub
