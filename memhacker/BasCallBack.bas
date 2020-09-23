Attribute VB_Name = "BasCallBack"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WHEEL_DELTA = 120

Public lPrevProc As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Dim b() As Byte
Dim OM As oh_message

Dim I As Integer
Dim I2 As Integer

Select Case Msg
    Case OH_MSG_HOOK
            'get the structure from memory
            ReDim b(Len(OM))
            ReadProcessMemory SelectedProcessHandle, wParam, b(0), Len(OM), 0
            CopyMemory OM, b(0), Len(OM)


            TRACE "name: " & OM.oName
            For I = 0 To lParam
                TRACE "param " & I & OM.oParams(I)
            Next I
    Case Else
    
End Select
    WindowProc& = CallWindowProc(lPrevProc, hwnd&, Msg&, wParam&, lParam&)
    Exit Function
ErrHandle:
    Err.Clear
End Function

