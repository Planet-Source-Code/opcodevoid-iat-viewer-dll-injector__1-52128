VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIAT 
   Caption         =   "Read the Import table"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step 2. Select a address"
      Height          =   3495
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid MsFlex 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5530
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin VB.ListBox LstAddress 
         Height          =   2400
         ItemData        =   "frmIAT.frx":0000
         Left            =   120
         List            =   "frmIAT.frx":0002
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 1. Select a dll"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ListBox LstDll 
         Height          =   2985
         ItemData        =   "frmIAT.frx":0004
         Left            =   120
         List            =   "frmIAT.frx":0006
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Process"
      Begin VB.Menu mnuInject 
         Caption         =   "Inject dll into process"
      End
   End
End
Attribute VB_Name = "frmIAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim I As Integer
'If UBound(IID) Is Nothing Then Unload Me: Exit Sub

If IID(0).lpName = 0 Then MsgBox "can't read dll's from this module ": Exit Sub
For I = 0 To UBound(IIDNames)
    Me.LstDll.AddItem IIDNames(I)
Next I
Me.LstDll.Refresh


MsFlex.FixedCols = 0 ' added adrey 2003-11-11
MsFlex.Cols = 4 ' changed from 3 to 4 adrey 2003-11-11
MsFlex.Col = 0
MsFlex.Row = 0
MsFlex.Text = "function name"
MsFlex.Col = 1
MsFlex.Text = "Ordinal"
MsFlex.Col = 2
MsFlex.Text = "Address"
MsFlex.Col = 3
MsFlex.Text = "Pointer Address"

MakeColsNiceSize 1800





End Sub

Private Sub LstDll_Click()
Dim IIN() As IMAGE_IMPORT_NAME
Dim I As Integer
If LstDll.ListIndex = -1 Then Exit Sub

RemoveAllRows

IIN = BasHardCoreSex.GetDllImportNamesAndAddress(IID(LstDll.ListIndex))
LstAddress.Clear
If IIN(0).tName = "*" Then Exit Sub

For I = 0 To UBound(IIN)
    'LstAddress.AddItem "Import name: " & IIN(I).tName & " Ordinal " & IIN(I).Ordinal
    AddNewEntryToCol IIN(I).tName, IIN(I).Ordinal, Hex(IIN(I).TAddress), Hex(IIN(I).TAddressOfImport)
    DoEvents
Next I

End Sub

Sub MakeColsNiceSize(Optional Size As Integer = 2000)
Dim I As Integer
For I = 0 To MsFlex.Cols - 1
  MsFlex.ColWidth(I) = Size
Next I

End Sub

Sub AddNewEntryToCol(tName As String, tOrdinal As Integer, TAddress As String, PAddress As String)

MsFlex.Rows = MsFlex.Rows + 1
MsFlex.Row = MsFlex.Rows - 1


MsFlex.Col = 0
MsFlex.Text = tName

MsFlex.Col = 1
MsFlex.Text = tOrdinal

MsFlex.Col = 2
MsFlex.Text = TAddress

MsFlex.Col = 3
MsFlex.Text = PAddress

End Sub

Sub RemoveAllRows()
Dim I As Integer
For I = MsFlex.Rows To 3 Step -1
   MsFlex.RemoveItem I
Next I

End Sub

Private Sub mnuInject_Click()
CD.DialogTitle = "Select a DLL"
CD.ShowOpen

Dim DLLPath As String: DLLPath = CD.FileName
Dim Ret As Long
Ret = AttachDllToProcess(DLLPath)
If Ret <> 1 Then TRACE "fail to attach dll to process "


End Sub
