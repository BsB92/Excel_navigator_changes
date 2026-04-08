Attribute VB_Name = "modFormSizeHook"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, _
         ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
        
        
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
        (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr

    Private Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
         ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
         ByVal uFlags As Long) As Long
        
#End If

Private Const GWL_WNDPROC As Long = -4
Private Const WM_GETMINMAXINFO As Long = &H24

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As LongPtr = &H40000
Private Const WS_MAXIMIZEBOX As LongPtr = &H10000
Private Const WS_MINIMIZEBOX As LongPtr = &H20000

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private mPrevWndProc As LongPtr
Private mMinW As Long, mMinH As Long, mMaxW As Long, mMaxH As Long
Private mActiveCaption As String
Public Sub EnableFormResize(ByVal formCaption As String)
    Dim hWnd As LongPtr
    Dim style As LongPtr
    Dim newStyle As LongPtr

    hWnd = FindWindow("ThunderDFrame", formCaption)
    If hWnd = 0 Then hWnd = FindWindow("ThunderXFrame", formCaption)
    If hWnd = 0 Then Exit Sub

    style = GetWindowLongPtr(hWnd, GWL_STYLE)
    newStyle = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX

    If newStyle <> style Then
        SetWindowLongPtr hWnd, GWL_STYLE, newStyle
        DoEvents
        SetWindowPos hWnd, 0, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If
End Sub

Public Sub HookMinMax(ByVal formCaption As String, ByVal minW As Long, ByVal minH As Long, ByVal maxW As Long, ByVal maxH As Long)
    Dim hWnd As LongPtr
    If mPrevWndProc <> 0 Then Exit Sub

    mActiveCaption = formCaption
    mMinW = minW: mMinH = minH
    mMaxW = maxW: mMaxH = maxH

    hWnd = FindWindow("ThunderDFrame", formCaption)
    If hWnd = 0 Then hWnd = FindWindow("ThunderXFrame", formCaption)
    If hWnd = 0 Then Exit Sub


    mPrevWndProc = SetWindowLongPtr(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnhookMinMax(ByVal formCaption As String)

    Dim hWnd As LongPtr
    If mPrevWndProc = 0 Then GoTo CLEANUP

    hWnd = FindWindow("ThunderDFrame", formCaption)
    If hWnd = 0 Then hWnd = FindWindow("ThunderXFrame", formCaption)
    If hWnd = 0 Then GoTo CLEANUP

    If mPrevWndProc <> 0 Then
        SetWindowLongPtr hWnd, GWL_WNDPROC, mPrevWndProc
    End If

CLEANUP:
    mPrevWndProc = 0
    mActiveCaption = vbNullString

End Sub


Private Function WindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                            ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    If uMsg = WM_GETMINMAXINFO Then
        Dim mmi As MINMAXINFO
        CopyMemory mmi, ByVal lParam, Len(mmi)

        mmi.ptMinTrackSize.X = mMinW
        mmi.ptMinTrackSize.Y = mMinH

        mmi.ptMaxTrackSize.X = mMaxW
        mmi.ptMaxTrackSize.Y = mMaxH

        CopyMemory ByVal lParam, mmi, Len(mmi)
        WindowProc = 0
        Exit Function
    End If

    WindowProc = CallWindowProc(mPrevWndProc, hWnd, uMsg, wParam, lParam)
End Function


