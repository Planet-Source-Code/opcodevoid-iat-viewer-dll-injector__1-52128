VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search Memory"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Manipulater"
      Height          =   1215
      Left            =   0
      TabIndex        =   21
      Top             =   6240
      Width           =   6615
      Begin VB.CommandButton CmdRead 
         Caption         =   "Read memory"
         Height          =   375
         Left            =   5160
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdWrite 
         Caption         =   "Write memory"
         Height          =   375
         Left            =   5160
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxTMValue 
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxTMAddress 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Value"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Found Addresses"
      Height          =   2655
      Left            =   2280
      TabIndex        =   18
      Top             =   3480
      Width           =   4335
      Begin VB.ListBox LstAddress 
         Height          =   2595
         ItemData        =   "frmSearch.frx":0000
         Left            =   120
         List            =   "frmSearch.frx":0002
         TabIndex        =   19
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Toolbar"
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   2295
      Begin VB.CommandButton CmdChange 
         Caption         =   "Value has changed"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton CmdDec 
         Caption         =   "Value has decrease"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton CmdInc 
         Caption         =   "Value has increase"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   6615
      Begin VB.TextBox TxTValue 
         Height          =   285
         Left            =   3600
         TabIndex        =   29
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox LstTypes 
         Height          =   315
         ItemData        =   "frmSearch.frx":0004
         Left            =   1440
         List            =   "frmSearch.frx":0011
         TabIndex        =   20
         Text            =   "Types"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "start new search"
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxTEnd 
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxTStart 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Search value(leave blank for unknown value)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Byte Width"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "End address"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Start address"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sections"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      Begin VB.ComboBox LstSections 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "Sections"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Section Address"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Size of Section"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Choose Section"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TO DO:
'
'Make code really generic
'

Option Explicit

Dim LongArray() As Long
Dim INTArray() As Integer
Dim ByteArray() As Byte

Dim TempLongArray() As Long
Dim TempIntArray() As Integer
Dim TempByteArray() As Byte

Dim FoundAddress() As Long
Dim FoundAddressCount As Long
Dim FoundArray As New Collection



Dim ArrayType As Byte 'one day i hope to genericlize this
Dim ArrayLength As Long


Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Const AT_LONG As Byte = 1

Private Sub CmdChange_Click()
'search for all values that changed
    
'the values are now stale, so we have to read the current memory values
Dim I As Long
LstAddress.Clear

If ArrayType = AT_LONG Then
    TempLongArray = ReadLongArrayFromMemory(CLng(TxTStart.Text), ArrayLength)
    For I = 0 To UBound(TempLongArray)
        If TempLongArray(I) <> GetValue(I) Then
          'i can stay
        Else
            FoundArray.Remove I ' i must go
        End If
        
      
    Next I
End If
DumpCollectionToList LstAddress, FoundArray 'update values
End Sub

Private Sub CmdDec_Click()
'search for all values that decreased
    
'the values are now stale, so we have to read the current memory values

Dim I As Long
LstAddress.Clear

If ArrayType = AT_LONG Then
    TempLongArray = ReadLongArrayFromMemory(CLng(TxTStart.Text), ArrayLength)
    For I = 1 To UBound(TempLongArray)
        If TempLongArray(I) < GetValue(I) Then
            'i can stay
        Else
            FoundArray.Remove I 'i must go
        End If
        Me.Caption = I & " out of " & UBound(TempLongArray)
        DoEvents
    Next I
End If
DumpCollectionToList Me.LstAddress, FoundArray

End Sub

Private Sub cmdINc_Click()
'search for all values that increased
    
'the values are now stale, so we have to read the current memory values
Dim I As Long
LstAddress.Clear

If ArrayType = AT_LONG Then
    TempLongArray = ReadLongArrayFromMemory(CLng(TxTStart.Text), ArrayLength)
    For I = 0 To UBound(TempLongArray)
        If TempLongArray(I) > GetValue(I) Then
            'i can stay
        Else
            FoundArray.Remove I 'i must go
        End If
    Next I
End If
DumpCollectionToList Me.LstAddress, FoundArray

End Sub

Private Sub CmdRead_Click()
Dim L As Long
If ArrayType = AT_LONG Then
     ReadProcessMemory SelectedProcessHandle, GetAbsValue(TxTMAddress.Text), L, 4, 0
    TxTMValue.Text = L
End If

End Sub

'0x00321140 Malloc start (debug)
'0x00321140 new start (debug)

'0x00321030 malloc start (release)
'0x00321030 new start (release)

'0x00143248 Global Alloc (release)
'0x00143240 Global Alloc (debug)


Private Sub CmdSearch_Click()

Erase FoundAddress
Erase TempLongArray
Erase TempIntArray
Erase TempByteArray
Erase LongArray
Erase INTArray
Erase ByteArray

Set FoundArray = Nothing
Set FoundArray = New Collection



If LstTypes.List(LstTypes.ListIndex) = "long(4)" Then
    ArrayLength = TxTEnd.Text - TxTStart.Text
    If ArrayLength <= 0 Then DoError "Start can not be greater than end ", ERRLEVEL_WARNING: Exit Sub
    'TO DO: check can length divide by 4
    Me.Caption = "Reading " & ArrayLength / 4 & " Longs from memory "
    ArrayLength = ArrayLength / 4
    If ArrayLength <= 0 Then DoError "There is not enough range to read a long from memory", ERRLEVEL_WARNING
    LongArray = ReadLongArrayFromMemory(CLng(TxTStart.Text), ArrayLength)
    'dump longarray into foundarray
    DumpArrayToCollection LongArray
    ArrayType = AT_LONG
    DumpCollectionToList LstAddress, FoundArray
End If

End Sub

Private Sub CmdWrite_Click()
Dim L As Long
Dim Old As Long
Dim Ret As Long
If ArrayType = AT_LONG Then

'Error log: VirtualProtectEx fails if the last parameter is null or points a
'a invalid location

     L = CLng(TxTMValue.Text)
  
     Ret = WriteProcessMemory(SelectedProcessHandle, CLng(TxTMAddress.Text), L, 4, Old)
     If Ret = 0 Then DoError "Failed to write memory because of [" & GetLastError() & "]"
     
     Me.Caption = " Wrote " & GetAbsValue(TxTMAddress.Text) & " To memory address " & TxTMAddress.Text
End If
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 0 To UBound(ImgSections)
    With ImgSections(I)
            Me.LstSections.AddItem ImgSections(I).sName
    End With
Next I

        
End Sub

Private Sub LstAddress_Click()
'
'
'681A20
'682C30
TxTMAddress.Text = GetAddress(LstAddress.ListIndex)

End Sub

Private Sub LstSections_Click()
Me.lblAddress = ImgSections(LstSections.ListIndex).PointerToRawData + ImageOptNT.ImageBase
Me.lblSize = ImgSections(LstSections.ListIndex).SizeOfRawData
TxTStart.Text = Me.lblAddress
TxTEnd.Text = Val(Me.lblAddress) + Val(Me.lblSize)

End Sub

Sub AddNewValue(pValue As Long)
FoundArray.Add pValue
End Sub

Sub DumpArrayToCollection(L() As Long)
Dim I As Integer
For I = 0 To UBound(L)
    FoundArray.Add L(I) & ":" & CLng(TxTStart.Text) + CLng(CLng(I) * 4)
Next I
End Sub

Sub DumpCollectionToList(L As ListBox, C As Collection)
Dim I As Integer
For I = 1 To C.Count
    L.AddItem C.Item(I)
Next I
End Sub

Function GetValue(Pindex As Long) As Long
GetValue = Mid$(FoundArray.Item(Pindex), 1, InStr(FoundArray.Item(Pindex), ":") - 1)

End Function

Function GetAddress(Pindex As Long) As Long
GetAddress = CLng(Mid(FoundArray.Item(Pindex), InStr(FoundArray.Item(Pindex), ":") + 1))
End Function

Function GetAbsValue(P As String) As Long
'converts from hex if nessarcy
'
'
Dim R As String

If Mid(P, 1, 2) = "0x" Then
    R = Mid(P, 3, Len(P))
    GetAbsValue = CLng(Val("&h" & R))
Else
    GetAbsValue = CLng(P)
End If

End Function

