VERSION 5.00
Begin VB.Form frmMemHacker 
   Caption         =   "_-/void opcode();\\-_"
   ClientHeight    =   4020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmToolBar 
      Caption         =   "Tool bar"
      Enabled         =   0   'False
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton CmdIAT 
         Caption         =   "Memory Hook"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMemHacker.frx":0000
      Height          =   1335
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Choose setup than choose a process than click Memory hook"
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuSelect 
         Caption         =   "Select a application"
      End
   End
End
Attribute VB_Name = "frmMemHacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'www.crackingislife.com
'


Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)


Private Sub CmdIAT_Click()
frmIAT.Show
End Sub

Private Sub CmdSearch_Click()
frmSearch.Show
End Sub



Private Sub Form_Load()
SaveSetting "ohacker", "datahacker", "path", App.Path & "\"
lPrevProc = SetWindowLong(Me.hwnd&, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'3 if any errors in the callback function it crashes, always start with full compile
'2.Window proc must call CallWindowProc
'1.You can't call a end
'MsgBox lPrevProc
'
Call SetWindowLong(Me.hwnd&, GWL_WNDPROC, lPrevProc)
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuSelect_Click()
frmProcess.Show
End Sub

