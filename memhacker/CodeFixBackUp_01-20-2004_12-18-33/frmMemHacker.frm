VERSION 5.00
Begin VB.Form frmMemHacker 
   Caption         =   "Opcodevoid Memory Hacker"
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
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Memory Search"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CmdIAT 
         Caption         =   "Memory Hook"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
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
Private Sub CmdIAT_Click()
frmIAT.Show
End Sub

Private Sub CmdSearch_Click()
frmSearch.Show
End Sub



Private Sub Form_Load()
SaveSetting "ohacker", "datahacker", "path", App.Path & "\"

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuSelect_Click()
frmProcess.Show
End Sub

