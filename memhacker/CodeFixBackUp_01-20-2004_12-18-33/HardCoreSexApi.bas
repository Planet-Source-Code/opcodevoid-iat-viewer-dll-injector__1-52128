Attribute VB_Name = "BasHardCoreSex"
'By: Opcodevoid
'Desc: Process manipultion, and IAT maniplution
'this stuff is hard core,
'
'
'Website: www.eliteproxy.com
'
'

Option Explicit


Private Const MAXPATH As Integer = 260

Private Const TH32CS_SNAPPROCESS As Integer = 2
Private Const STANDARD_HEADER = &H400000

Public Const IMAGE_DOS_SIGNATURE = &H5A4D        ''\\ MZ
Public Const IMAGE_OS2_SIGNATURE = &H454E        ''\\ NE
Public Const IMAGE_OS2_SIGNATURE_LE = &H454C     ''\\ LE
Public Const IMAGE_VXD_SIGNATURE = &H454C        ''\\ LE
Public Const IMAGE_NT_SIGNATURE = &H4550         ''\\ PE00

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessId As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessId As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAXPATH
End Type

Private Type IMAGE_DOS_HEADER
    e_magic As String * 2  ''\\ Magic number
    e_cblp As Integer    ''\\ Bytes on last page of file
    e_cp As Integer      ''\\ Pages in file
    e_crlc As Integer    ''\\ Relocations
    e_cparhdr As Integer ''\\ Size of header in paragraphs
    e_minalloc As Integer ''\\ Minimum extra paragraphs needed
    e_maxalloc As Integer ''\\ Maximum extra paragraphs needed
    e_ss As Integer    ''\\ Initial (relative) SS value
    e_sp As Integer    ''\\ Initial SP value
    e_csum As Integer  ''\\ Checksum
    e_ip As Integer  ''\\ Initial IP value
    e_cs As Integer  ''\\ Initial (relative) CS value
    e_lfarlc As Integer ''\\ File address of relocation table
    e_ovno As Integer ''\\ Overlay number
    e_res(0 To 3) As Integer ''\\ Reserved words
    e_oemid As Integer ''\\ OEM identifier (for e_oeminfo)
    e_oeminfo As Integer ''\\ OEM information; e_oemid specific
    e_res2(0 To 9) As Integer ''\\ Reserved words
    e_lfanew As Long ''\\ File address of new exe header
End Type

Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
End Type
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Public Type IMAGE_IMPORT_DESCRIPTOR
    lpImportByName As Long ''\\ The names
    TimeDateStamp As Long  ''\\ 0 if not bound,
                           ''\\ -1 if bound, and real date\time stamp
                           ''\\ in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                           ''\\ O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain As Long ''\\ -1 if no forwarders
    lpName As Long
    lpFirstThunk As Long ''\\ The actual addresses
End Type
Public Type IMAGE_OPTIONAL_HEADER_NT
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Private Const IMAGE_SIZEOF_SHORT_NAME = 8

Public Type IMAGE_SECTION_HEADER
    sName As String * IMAGE_SIZEOF_SHORT_NAME
    PhysicalAddress As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type


    
    


Public Type IMAGE_IMPORT_NAME 'custom structure i made, has nothing to do with PE file
    Ordinal As Integer
    tName As String
    TAddress As Long
    TAddressOfImport As Long
    
End Type



Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const PAGE_EXECUTE_READWRITE = &H40
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_WRITE = &H20

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'Error log: one of my hugest errors was in the way WriteProcessMemory was declared
'simply chaing a few byval to byrefs, and vesa versa , bam it works
'
'

Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub CopyMemoryByval Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

Public ImgSections() As IMAGE_SECTION_HEADER 'section for the current process
Public IID() As IMAGE_IMPORT_DESCRIPTOR
Public IIDNames() As String



Public Function GetProcessHandle(PId As Long) As Long
GetProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PId)
End Function

Public Function GetAllProcesses() As PROCESSENTRY32()
Dim P() As PROCESSENTRY32
Dim T As PROCESSENTRY32
Dim Shot As Long
Dim pCount As Integer
Dim R As Long

ReDim P(0)
pCount = 1
Shot = CreateToolhelpSnapshot(PROCESS_ALL_ACCESS, 0)
T.dwSize = Len(T)
ProcessFirst Shot, T
R = 1

P(0) = T



Do While R
    ReDim Preserve P(pCount)
    R = ProcessNext(Shot, T)
    If R = 0 Then Exit Do
    P(pCount) = T
    If P(pCount).th32ProcessId = 0 Then MsgBox "Found a null id at " & P(pCount).szExeFile
    
    pCount = pCount + 1
    DoEvents
Loop
CloseHandle Shot
GetAllProcesses = P
End Function

Public Function GetAllProcessDLLS() As IMAGE_IMPORT_DESCRIPTOR()
Dim DosHeader As IMAGE_DOS_HEADER
Dim ImageHeader As IMAGE_FILE_HEADER
Dim ImageOpt As IMAGE_OPTIONAL_HEADER
Dim ImportDesc As IMAGE_IMPORT_DESCRIPTOR ' the hard core sex stuff
Dim ImportDescArray() As IMAGE_IMPORT_DESCRIPTOR ' the hard core sex array
Dim ImportDescArrayCount As Long 'the hard core sex counter
Dim ImportDescAddr As Long
Dim TempSection As IMAGE_SECTION_HEADER


Dim lRead As Long
Dim Result As Long
Dim Str As String
Dim Old As Long
Dim I As Integer


Dim b() As Byte
ReDim b(Len(DosHeader))
ReDim ImportDescArray(0)
ImportDescArray(0).lpName = 0


    
Result = ReadProcessMemory(SelectedProcessHandle, GetHeaderAddress, b(0), Len(DosHeader), 0&)
CopyMemory DosHeader, b(0), Len(DosHeader)     'copy it to the structure and bail out

If Result = 0 Then 'attempt to remove protection
   Result = VirtualProtectEx(SelectedProcessHandle, GetHeaderAddress, Len(DosHeader), PAGE_EXECUTE_READWRITE, Old)
   If Result = 0 Then DoError "failed to protect pages " & GetLastError
   
    Result = ReadProcessMemory(SelectedProcessHandle, GetHeaderAddress, b(0), Len(DosHeader), 0&)
    CopyMemory DosHeader, b(0), Len(DosHeader)     'copy it to the structure and bail out
End If

If DosHeader.e_magic <> "MZ" Then DoError "Failed to get Dos Header <" & Result & ":->" & GetLastError & ">": Exit Function
TRACE "PE offset at " & DosHeader.e_lfanew

DosHeader.e_lfanew = DosHeader.e_lfanew + GetHeaderAddress

ReDim b(4) ' get pe\0\0

Call ReadProcessMemory(SelectedProcessHandle, DosHeader.e_lfanew, b(0), 4, 0)
If ByteArrayToString(b) <> "PE" Then DoError "Failed to get Pe header ": Exit Function
TRACE "Found the pe header hot dog"

ReDim b(Len(ImageHeader))
Call ReadProcessMemory(SelectedProcessHandle, DosHeader.e_lfanew + 4, b(0), Len(ImageHeader), 0)
'read the image header the dosheader.e_lfanew + 4 (the +4 is there because we skip over the pe\0\0)
'FYI
CopyMemory ImageHeader, b(0), Len(ImageHeader) 'reading memory from a process is a choir really
TRACE "Number of section: " & ImageHeader.NumberOfSections

'now right under this header is the optional header
'its no where near optional, but you know microsoft when it
'comes to names , i mean what on earth is a xp

'update the pointer so it points to the optional header
'why use dosheader.e_lfanew, well because i'm to lazy to declare variables


DosHeader.e_lfanew = DosHeader.e_lfanew + 4 + Len(ImageHeader)


ReDim b(Len(ImageOpt))

Call ReadProcessMemory(SelectedProcessHandle, DosHeader.e_lfanew, b(0), Len(ImageOpt), 0)
CopyMemory ImageOpt, b(0), Len(ImageOpt)
TRACE "Entry point:  " & ImageOpt.AddressOfEntryPoint
TRACE "Base of code: " & ImageOpt.BaseOfCode

'now lets finaly read the freaking good stuff

ReDim b(Len(ImageOptNT))
'lets point to the header first ImageOpt
DosHeader.e_lfanew = DosHeader.e_lfanew + Len(ImageOpt)

'why do i make to reads like a freaking idiot for the same header
'because the stupid idiot i rip these definition from did

Call ReadProcessMemory(SelectedProcessHandle, DosHeader.e_lfanew, b(0), Len(ImageOptNT), 0)
CopyMemory ImageOptNT, b(0), Len(ImageOptNT)

TRACE "Offset To Import(cess poll) table: " & Hex(ImageOptNT.DataDirectory(1).VirtualAddress)
TRACE "Image base(RVA helper): " & Hex(ImageOptNT.ImageBase)

ImageOptNT.DataDirectory(1).VirtualAddress = ImageOptNT.DataDirectory(1).VirtualAddress + ImageOptNT.ImageBase
ImportDescAddr = ImageOptNT.DataDirectory(1).VirtualAddress
'do we really want to be typing imageopt.datadirectory(1).virtualaddress every freaking
'time


'we must keep reading until we read a null descripter
ReDim ImportDescArray(0)
ReDim b(Len(ImportDesc)) 'setup the structure reader byte

ImportDesc.lpImportByName = 1  'so we don't exit the loop as soon as we start

'side effects: we covert the lpImportByName field from a RVA to a VA
'side effects: we do the same for lpFirstThunk, and lpName

Do While ImportDesc.lpImportByName <> 0
    Result = ReadProcessMemory(SelectedProcessHandle, ImportDescAddr, b(0), Len(ImportDesc), 0)
    If Result = 0 Then DoError "Failed to read import table, application may be using protection "
    CopyMemory ImportDesc, b(0), Len(ImportDesc)
    
    If ImportDesc.lpImportByName <> 0 Then  'its a valid descriptor
      
      ImportDesc.lpImportByName = ImportDesc.lpImportByName + ImageOptNT.ImageBase
      ImportDesc.lpFirstThunk = ImportDesc.lpFirstThunk + ImageOptNT.ImageBase
      ImportDesc.lpName = ImportDesc.lpName + ImageOptNT.ImageBase
      
      CopyMemory ImportDescArray(ImportDescArrayCount), ImportDesc, Len(ImportDesc)
      ImportDescArrayCount = ImportDescArrayCount + 1
      ReDim Preserve ImportDescArray(ImportDescArrayCount) 'if you want the size use a ubound
      TRACE "Name: " & Hex(ImportDesc.lpImportByName)
    End If
    'increase pointer to next descriptor
    ImportDescAddr = ImportDescAddr + Len(ImportDesc)
Loop

'lets read sections

'1.make sure i have sections not ever one does

DosHeader.e_lfanew = DosHeader.e_lfanew + Len(ImageOptNT)

ReDim ImgSections(ImageHeader.NumberOfSections - 1)
ReDim b(Len(TempSection))


For I = 0 To ImageHeader.NumberOfSections - 1
   Call ReadProcessMemory(SelectedProcessHandle, DosHeader.e_lfanew, b(0), Len(TempSection), 0)
   CopyMemory TempSection, b(0), Len(TempSection)
   ImgSections(I) = TempSection
   DosHeader.e_lfanew = DosHeader.e_lfanew + Len(TempSection)
   DoEvents 'incase we have alot of sections
Next I

GetAllProcessDLLS = ImportDescArray

End Function
Public Function ResolveImportNames(IID() As IMAGE_IMPORT_DESCRIPTOR) As String()
Dim I As Integer
Dim S() As String
Dim SCount As Long
Dim b() As Byte

For I = 0 To UBound(IID) - 1 'the last import desciptor is null
    ReDim Preserve S(SCount)
    S(SCount) = ReadStringFromMemory(IID(I).lpName, 32000)  'maxium string size(almost)
    TRACE "Dll import name is " & S(SCount)
    SCount = SCount + 1
    DoEvents 'incase its alot of names
Next I
ResolveImportNames = S
End Function

Public Function GetDllImportNamesAndAddress(IID As IMAGE_IMPORT_DESCRIPTOR) As IMAGE_IMPORT_NAME()
'Microsoft was sure to make this very difficault and odd to read
'
'
'Seems to have a problem reading wsock.dll, and doesn't seem to read correctly
'if there is just 1 import.
'

Dim I As Integer
Dim NameAddr As Long
Dim NameAddrPos As Long
Dim IIN As IMAGE_IMPORT_NAME
Dim Reading As Boolean
Dim TheRealIIN() As IMAGE_IMPORT_NAME
Dim INNCount As Long

ReDim TheRealIIN(0)
TheRealIIN(0).tName = "*"

Reading = True

Do While Reading
    NameAddr = ReadLongFromMemory(IID.lpImportByName + NameAddrPos)
  
    If NameAddr <= 0 Then GoTo Done 'speghitaii programming
    NameAddr = NameAddr + ImageOptNT.ImageBase 'convert the RVA->VA(if i'm not mistaken)
    
    'we can't just read the string from memory because microsoft saw it wise to put
    'ordinals there(crap......) and we can't just read a struct from memory
    'because we don't know how long the string is
    'whats does microsoft tell us.... Its are code just
    'make it work(stupid microsoft slogo...)
    
       IIN.Ordinal = ReadIntFromMemory(NameAddr)
       TRACE "My ordinal is " & IIN.Ordinal
       IIN.tName = ReadStringFromMemory(NameAddr + 2, 32000) 'simply skip over the ordinal
       TRACE "My name is " & IIN.tName
       
       'the address are stored in a seperate array
       NameAddr = ReadLongFromMemory(IID.lpFirstThunk + NameAddrPos)
              
       TRACE "Address of function is at " & Hex(NameAddr)
       IIN.TAddress = NameAddr
       ReDim Preserve TheRealIIN(INNCount)
       With TheRealIIN(INNCount)
         .Ordinal = IIN.Ordinal
         .tName = IIN.tName
         .TAddressOfImport = IID.lpFirstThunk + NameAddrPos
         .TAddress = IIN.TAddress
       End With
       
       INNCount = INNCount + 1
       
       NameAddrPos = NameAddrPos + 4 'update the pointer
       
Loop
Done:
GetDllImportNamesAndAddress = TheRealIIN

End Function


Private Function GetHeaderAddress() As Long
'Old Comment:one day we won't just guess we will detect it
'thus becomming strong and leet
'
'New Comment: We know longer guess, we are strong and leet now
'
'
'

Dim b(MAXPATH) As Byte
Dim Str As String
Dim FreeIndex
Dim DosHeader As IMAGE_DOS_HEADER
Dim ImageHeader As IMAGE_FILE_HEADER
Dim ImageOpt As IMAGE_OPTIONAL_HEADER
Dim ImageOptNT As IMAGE_OPTIONAL_HEADER_NT
FreeIndex = FreeFile

Str = HackerHooks.GetFilename(SelectedProcessID)


Open Str For Binary Access Read As FreeIndex
    Get #FreeIndex, , DosHeader
    If DosHeader.e_magic <> "MZ" Then TRACE "Failed to get dos header in GetHeaderAddress"
    'we use +5 because 0 is not a acceptable positon to read a file
    'in memory it is, but in a file, 1 is the beginning
    
    Get #FreeIndex, DosHeader.e_lfanew + 5 + Len(ImageHeader) + Len(ImageOpt), ImageOptNT
    TRACE "Base of code is at : " & Hex(ImageOptNT.MajorOperatingSystemVersion)
    Close FreeIndex




GetHeaderAddress = ImageOptNT.ImageBase



End Function

Private Function ByteArrayToString(b() As Byte) As String
'This is what you call a quicky
'
'
'
Dim I As Integer
Dim S As String
For I = 0 To UBound(b)
  If b(I) = 0 Then ByteArrayToString = S: Exit Function
  S = S & Chr(b(I))
Next I

ByteArrayToString = S
End Function
