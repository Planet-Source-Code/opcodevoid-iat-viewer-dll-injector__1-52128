VERSION 5.00
Begin VB.Form frmFake 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'since we can stop this window from being the top most
'we will hack to hide it
Dim P() As Long

If VarType(P) = vbError Then MsgBox "done"

frmMemHacker.Show
'Main
Me.Hide






End Sub
