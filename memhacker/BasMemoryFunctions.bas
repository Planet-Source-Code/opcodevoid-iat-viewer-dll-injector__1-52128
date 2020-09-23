Attribute VB_Name = "BasMemoryFunctions"
'
'www.crackingislife.com
'
'

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT    As Long = &H1000
Private Const MEM_DECOMMIT  As Long = &H4000

Public Type typInjectCode
    IAddr As Long
    ISize As Long
End Type



Public Sub ReadStructFromMemory(MemoryAddress As Long, StructAddr As Long, Length As Long)
'==============I can't get this function to work
'One day it may work :)
'
'
'Visual basic is a real pain, so we must yet create another helper fuction hacker thingy
'type something....
'


Dim Buffer As Byte
Dim ProcessHandle As Long
Dim Str As String
Dim b() As Byte
ReDim b(Length)
    
Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress, b(0), Length, 0&)
CopyMemory ByVal StructAddr, ByVal b(0), Length     'copy it to the structure and bail out

End Sub

Public Function ReadStringFromMemory(MemoryAddress As Long, Length As Long)
'Visual basic can't read strings from other proccess, well it can
'but visual basic strings are in some werid gay(no offense if your gay) format
'
'so we must read a byte at a time

'This function automaticly stops at null bytes
'
'
'
Dim Buffer As Byte
Dim ProcessHandle As Long
Dim Str As String
Dim I As Long

Do While I < Length
    Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress + I, Buffer, 1, 0&)
    If Buffer = 0 Then Exit Do
    Str = Str & Chr(Buffer)
    I = I + 1
    DoEvents
Loop

ReadStringFromMemory = Str

End Function


Public Function ReadLongFromMemory(MemoryAddress As Long) As Long
Dim Buffer As Long
Dim ProcessHandle As Long
Dim Str As String

Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer, 4, 0&)
ReadLongFromMemory = Buffer

End Function


Public Function ReadLongArrayFromMemory(MemoryAddress As Long, NumberOfLongs As Long) As Long()

Dim Buffer() As Long
Dim Ret As Long
ReDim Buffer(NumberOfLongs)
Dim Old As Long

Ret = VirtualProtectEx(SelectedProcessHandle, MemoryAddress, NumberOfLongs * 4, &H40, Old)
If Ret = 0 Then TRACE "VirtualProtectEx suffers from error " & GetLastError



Ret = ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer(0), NumberOfLongs * 4, 0&)
If Ret = 0 Then DoError "Failed to read Longs from memory " & GetLastError: TRACE "ReadLongArrayFromMemory Suffers from error " & GetLastError()
ReadLongArrayFromMemory = Buffer

End Function

Public Function ReadIntFromMemory(MemoryAddress As Long) As Integer
'Please note, all real programmers know that on a Intel 32 platform
'Integers are usually 4 bytes, but microsoft had to go ruin that
'
'

Dim Buffer As Integer
Dim ProcessHandle As Long
Dim Str As String

Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer, 2, 0&)
ReadIntFromMemory = Buffer

End Function

Public Function InjectCodeintoProcess(pFile As String) As typInjectCode
On Error GoTo Dope
Dim Buffer As String
Dim FreeIndex
Dim IC As typInjectCode


FreeIndex = FreeFile
Buffer = String(FileLen(pFile), "Z")

Open pFile For Binary Access Read As #FreeIndex
 Get #FreeIndex, , Buffer
Close #FreeIndex
Exit Function
IC.ISize = FileLen(pFile)
IC.IAddr = VirtualAllocEx(SelectedProcessHandle, 0, FileLen(pFile), MEM_COMMIT, &H40) '&h40 = Page_execute_readWrite
If IC.IAddr = 0 Then TRACE "Failed to allocate memory for code injection ": Exit Function
TRACE "Write process memory = " & WriteProcessMemory(SelectedProcessHandle, IC.IAddr, Buffer, FileLen(pFile), 0)
InjectCodeintoProcess = IC
Exit Function
Dope:
IC.IAddr = 0
IC.ISize = 0 'test is ISize 0 for failure
InjectCodeintoProcess = IC
End Function

Public Sub UnInjectCode(IC As typInjectCode)
VirtualFree SelectedProcessHandle, IC.IAddr, IC.ISize, MEM_DECOMMIT
End Sub
