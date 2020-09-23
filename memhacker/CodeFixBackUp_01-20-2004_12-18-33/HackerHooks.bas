Attribute VB_Name = "HackerHooks"
'Visual basic 6.0 Doing Api hooking , yes i have lost my mind
'
'
'By: Opcodevoid
'Website: www.eliteproxy.com
'
'
'
'

'To make sure this work WndProc must be declare a function and return a long value
'not a variant , a long!!!
'also you must return the value returned from DefWindowProc when nessarcy
'
'

Option Explicit
Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type


Public Const NM_FIRST As Long = 0
Public Const NM_DBLCLK As Long = (NM_FIRST - 3)
Public Const WM_NOTIFY As Long = &H4E


Public StatusBar As Long
Public ListView As Long
Public Label As Long
Public SeekButton As Long
Public OpenButton As Long
Public Window As Long
Public hFont As Long
Public Menu As Long
' End hWnd Variables Block

' Holds the structures of items and colums for the listview, so we can use it for manipulation after
Dim SBPartsWidths(1) As Long
Dim LVC As LVCOLUMN
Public LVI As LVITEM
Public ControlHeader As NMHDR
' End listview structures

' Same idea as above, but for the Open File Dialog
Dim OFN As OPENFILENAME
' End Open File Dialog Variable

' Same idea as above, but for the Menu
Dim MII As MENUITEMINFO
' End Menu Variable

' All the constants we use when calling the APIs...I won't comment them all, most are clear to understand
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WM_USER = &H400
Public Const WM_DESTROY As Long = &H2
Public Const SB_SETPARTS = (WM_USER + 4)
Public Const SB_SETTEXTA = (WM_USER + 1)
Public Const DEFAULT_GUI_FONT As Long = 17
Public Const WM_SETFONT As Long = &H30
Public Const WM_SETTEXT As Long = &HC
Public Const LVM_FIRST As Long = &H1000
Public Const LVCF_TEXT As Long = &H4
Public Const LVCF_WIDTH As Long = &H2
Public Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Public Const LVIF_TEXT As Long = &H1
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Public Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Public Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Public Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Public Const LVM_DELETECOLUMN = LVM_FIRST + 28
Public Const LVS_REPORT As Long = &H1
Public Const WS_BORDER As Long = &H800000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_CAPTION As Long = &HC00000
Public Const WM_COMMAND As Long = &H111
Public Const ICC_BAR_CLASSES As Long = &H4
Public Const ICC_LISTVIEW_CLASSES As Long = &H1
Public Const MIIM_STRING As Long = &H40
Public Const MIIM_ID As Long = &H2
Public Const TPM_RETURNCMD As Long = &H100&
Public Const WM_CONTEXTMENU As Long = &H7B
Public Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Public Const LVNI_SELECTED As Long = &H2
Public Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
' End Constants

' Listview Item Structure
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
' End Listview Item Structure

' Listview Column Structure
Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type
' End Listview Item Structure

' Window Structure
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
' End Window Structure

' Mouse location structure (called by MSG)
Public Type POINTAPI
    X As Long
    Y As Long
End Type
' End Mouse location structure

' Window Message structure
Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
' Window Message structure

' Common Control Initialisation Structure
Public Type INITCOMMONCONTROLSEX
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type
' End Common Control Initialisation Structure

' Common Dialog OpenFile Structure
Public Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type
' End Common Dialog OpenFile Structure

' Menu Structure
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
' End Menu Structure

' APIs used to Create the form, controls and perform message manipulation
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hmenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function TranslateMessage Lib "user32.dll" (lpMsg As MSG) As Long

Public Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)
Public Declare Function INITCOMMONCONTROLSEX Lib "comctl32.dll" Alias "InitCommonControlsEx" (ByRef TLPINITCOMMONCONTROLSEX As INITCOMMONCONTROLSEX) As Long
Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hmenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hmenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long

Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByRef lphModule As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long

Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByVal hModule As Long, _
                                                        ByVal ModuleName As String, _
                                                        ByVal nSize As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT    As Long = &H1000
Private Const MEM_DECOMMIT  As Long = &H4000

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public TheHookHWND As Long
Public IsRunning As Boolean


Public Function AttachDllToProcess(DLLPath As String, Optional ProcessHandle As Long = -1)
Dim PHandle As Long
Dim LoadAddr As Long
Dim SA As SECURITY_ATTRIBUTES
Dim Ret As Long
Dim TID As Long
Dim b() As Byte
Dim RemoteDllPath As Long
PHandle = SelectedProcessHandle

Dim StrDll As String * 255
StrDll = DLLPath
LoadAddr = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
If LoadAddr = 0 Then DoError "Can't find LoadLibrary from kernel32.dll": Exit Function
TRACE "ProcessHandle: " & PHandle & " LoadAddr: " & Hex(LoadAddr) & " DllPath: " & StrDll
'The parameter must be stored remotely
RemoteDllPath = VirtualAllocEx(SelectedProcessHandle, ByVal 0, Len(DLLPath) + 1, MEM_COMMIT, ByVal &H4)

If RemoteDllPath = 0 Then DoError "Failed to allocate virtual memory for process -> " & GetLastError: Exit Function
Ret = WriteProcessMemory(SelectedProcessHandle, RemoteDllPath, StrDll, Len(DLLPath) + 1, Ret)
If Ret = 0 Then DoError "Failed to write dll to process memory -> " & GetLastError




Ret = CreateRemoteThread(PHandle, ByVal 0, 0, LoadAddr, RemoteDllPath, 0, TID)
If Ret = 0 Then DoError "Failed to CreateRemoteThead-> " & GetLastError: Exit Function
TRACE "Thread Handle: " & Ret & " Thread ID: " & TID & " Remote Dll Address: " & RemoteDllPath
AttachDllToProcess = 1



End Function




Public Function WndProc(ByVal hwnd As Long, ByVal UINT As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case UINT
    Case WM_DESTROY
        MsgBox "I might crash "
    Case Else

End Select


WndProc = DefWindowProc(hwnd, UINT, wParam, lParam)
       

End Function
Private Function GetAddress(P As Long)
GetAddress = P
End Function

Sub Main()
Dim aMsg As MSG                                                 ' Needed for our MessageLoop
Dim icc As INITCOMMONCONTROLSEX                                 ' Initialize the common controls we use
icc.dwSize = Len(icc)                                           ' Initialize the common controls we use
icc.dwICC = ICC_LISTVIEW_CLASSES Or ICC_BAR_CLASSES             ' Initialize the common controls we use
INITCOMMONCONTROLSEX icc                                        ' Initialize the common controls we use
CreateForm                                                      ' Create the form

IsRunning = True
Do While IsRunning = True
    If PeekMessage(aMsg, TheHookHWND, 0, 0, 1) <> 0 Then
          TranslateMessage aMsg
          DispatchMessage aMsg
    End If
    DoEvents
  
Loop                                                            ' End Message loop

UnregisterClass "NTFSClass", App.hInstance
End Sub
Public Sub CreateForm()
Dim wc As WNDCLASS
With wc
    .lpfnwndproc = GetAddress(AddressOf WndProc)  ' I don't know how to subclass in a remote thread yet, so I'm telling windows to use the default subclasser
    .hbrBackground = 5 ' Default color for a window
    .lpszClassName = "NTFSClass" ' Name of our class
End With
RegisterClass wc ' Register it
TheHookHWND = CreateWindowEx(0&, "NTFSClass", "The Hidden Hook", WS_CAPTION Or WS_SYSMENU, 300, 300, 448, 298, 0, 0, App.hInstance, ByVal 0&)

'we keep the window hidden because there is no reason to see it
'I took learn how to do the window from a NTFS found of pscode


'ShowWindow TheHookHWND, 1
'UpdateWindow TheHookHWND
'SetFocus TheHookHWND

End Sub
Public Function GetFilename(lngProcessID As Long) As String
  Dim cb As Long
  Dim cbNeeded As Long
  Dim NumElements As Long
  Dim ProcessIDs() As Long
  Dim cbNeeded2 As Long
  Dim NumElements2 As Long
  Dim Modules(1 To 200) As Long
  Dim lRet As Long
  Dim ModuleName As String
  Dim nSize As Long
  Dim hProcess As Long
  Dim I As Long
  Dim Y As Long
  Dim q As Long
  Dim Huh As Boolean

 'Get a handle to the Process
         hProcess = OpenProcess(&H1F0FFF, 0, lngProcessID)
      'Got a Process handle
         If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
             lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
             
           'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(260)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                                     
                                     
                GetFilename = Left(ModuleName, lRet)
                
                End If
            End If
    
               
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
End Function

