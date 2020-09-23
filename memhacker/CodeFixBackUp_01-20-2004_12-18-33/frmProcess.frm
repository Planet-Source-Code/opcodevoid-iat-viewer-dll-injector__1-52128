VERSION 5.00
Begin VB.Form frmProcess 
   Caption         =   "Task List"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstLIst 
      Height          =   5130
      ItemData        =   "frmProcess.frx":0000
      Left            =   0
      List            =   "frmProcess.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Simply double click a process"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim I As Integer

tProcessList = BasHardCoreSex.GetAllProcesses
For I = 0 To UBound(tProcessList)
    LstLIst.AddItem tProcessList(I).szExeFile
Next I
LstLIst.Refresh

End Sub


Private Sub LstLIst_DblClick()
'reset all allocated memory
Erase IID
Erase IIDNames
Erase ImgSections

If LstLIst.ListIndex = -1 Then Exit Sub
SelectedProcessID = tProcessList(LstLIst.ListIndex).th32ProcessId
SelectedProcessHandle = GetProcessHandle(SelectedProcessID)

IID = GetAllProcessDLLS
IIDNames = ResolveImportNames(IID)
frmMemHacker.frmToolBar.Enabled = True
Unload Me
End Sub
