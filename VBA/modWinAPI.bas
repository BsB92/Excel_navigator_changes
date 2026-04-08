Attribute VB_Name = "modWinAPI"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const HWND_TOPMOST As LongPtr = -1
Public Const HWND_NOTOPMOST As LongPtr = -2
Public Const VK_ESCAPE As Long = &H1B

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_NOACTIVATE As Long = &H10

Public Const SW_RESTORE As Long = 9
Public Const SW_SHOW As Long = 5

Public Function FindFormWindow(ByVal formCaption As String) As LongPtr
    Dim h As LongPtr

    h = FindWindow("ThunderDFrame", formCaption)
    If h = 0 Then h = FindWindow("ThunderXFrame", formCaption)

    FindFormWindow = h
End Function

Public Sub SetTopMostState(ByVal formCaption As String, ByVal makeTopMost As Boolean)
    Dim h As LongPtr

    h = FindFormWindow(formCaption)
    If h = 0 Then Exit Sub

    If makeTopMost Then
        SetWindowPos h, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Else
        SetWindowPos h, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    End If
End Sub

Public Sub BringFormToFront(ByVal formCaption As String)
    Dim h As LongPtr

    h = FindFormWindow(formCaption)
    If h = 0 Then Exit Sub

    If IsIconic(h) <> 0 Then
        ShowWindow h, SW_RESTORE
    Else
        ShowWindow h, SW_SHOW
    End If

    SetWindowPos h, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    SetForegroundWindow h
End Sub
