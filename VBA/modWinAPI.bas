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
    Public Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwndLock As LongPtr) As Long
#Else
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
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

Public Type RECT
    Left As Long
    TOP As Long
    Right As Long
    Bottom As Long
End Type

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function EnumDisplayMonitorsAPI Lib "user32" Alias "EnumDisplayMonitors" (ByVal hdc As LongPtr, ByVal lprcClip As LongPtr, ByVal lpfnEnum As LongPtr, ByVal dwData As LongPtr) As Long
    Private Declare PtrSafe Function GetMonitorInfoAPI Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, lpmi As MONITORINFO) As Long
#Else
    Private Declare Function EnumDisplayMonitorsAPI Lib "user32" Alias "EnumDisplayMonitors" (ByVal hdc As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
    Private Declare Function GetMonitorInfoAPI Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, lpmi As MONITORINFO) As Long
#End If

Private mTargetMonitorIndex As Long
Private mFoundMonitor As Boolean
Private mFoundWork As RECT
Private mEnumMonitorIndex As Long

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

Public Sub SuspendExcelWindowUpdates()
    On Error Resume Next
    LockWindowUpdate Application.hWnd
    On Error GoTo 0
End Sub

Public Sub ResumeExcelWindowUpdates()
    On Error Resume Next
    LockWindowUpdate 0
    On Error GoTo 0
End Sub

Public Function TryGetMonitorWorkArea(ByVal monitorIndex As Long, _
                                      ByRef workLeft As Long, _
                                      ByRef workTop As Long, _
                                      ByRef workWidth As Long, _
                                      ByRef workHeight As Long) As Boolean
    If monitorIndex < 1 Then Exit Function

    mTargetMonitorIndex = monitorIndex
    mFoundMonitor = False
    mEnumMonitorIndex = 0

#If VBA7 Then
    EnumDisplayMonitorsAPI 0, 0, AddressOf EnumMonitorsProc, 0
#Else
    EnumDisplayMonitorsAPI 0, 0, AddressOf EnumMonitorsProc, 0
#End If

    If Not mFoundMonitor Then Exit Function

    workLeft = mFoundWork.Left
    workTop = mFoundWork.TOP
    workWidth = mFoundWork.Right - mFoundWork.Left
    workHeight = mFoundWork.Bottom - mFoundWork.TOP
    TryGetMonitorWorkArea = (workWidth > 0 And workHeight > 0)
End Function

#If VBA7 Then
Public Function EnumMonitorsProc(ByVal hMonitor As LongPtr, ByVal hdcMonitor As LongPtr, ByRef lprcMonitor As RECT, ByVal dwData As LongPtr) As Long
#Else
Public Function EnumMonitorsProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByVal dwData As Long) As Long
#End If
    Dim mi As MONITORINFO

    mEnumMonitorIndex = mEnumMonitorIndex + 1

    If mEnumMonitorIndex = mTargetMonitorIndex Then
        mi.cbSize = Len(mi)
        If GetMonitorInfoAPI(hMonitor, mi) <> 0 Then
            mFoundWork = mi.rcWork
            mFoundMonitor = True
        End If
        mEnumMonitorIndex = 0
        EnumMonitorsProc = 0
        Exit Function
    End If

    EnumMonitorsProc = 1
End Function
