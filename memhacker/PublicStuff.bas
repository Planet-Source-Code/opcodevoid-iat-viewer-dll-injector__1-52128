Attribute VB_Name = "PublicStuff"
'By : Opcodevoid
'
'Desc: this is where all public variables go
'
'Website: www.crackingislife.com
'



Public tProcessList() As PROCESSENTRY32
Public SelectedProcessID As Long
Public SelectedProcessHandle As Long
Public ImageOptNT As IMAGE_OPTIONAL_HEADER_NT
Public InJectCode As typInjectCode
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type oh_message
  oName As String * 255
  oParams(50) As Long
End Type
Public Const OH_MSG_HOOK As Long = &H401

Public Sub SafeExit()
SendMessage TheHookHWND, WM_QUIT, 0, 0

End Sub
