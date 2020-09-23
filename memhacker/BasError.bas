Attribute VB_Name = "BasError"
'By : Opcodevoid
'
'Desc: Some Basic Error handling routines
'
'
'Website: www.crackingislife.com
'
Public Const ERRLEVEL_FATAL As Integer = 1
Public Const ERRLEVEL_WARNING As Integer = 0

Public Sub DoError(ErrDesc As String, Optional ErrLevel As Long = ERRLEVEL_WARNING)
MsgBox "Error occured " & ErrDesc
If ErrLevel = ERRLEVEL_FATAL Then End

End Sub


Public Sub TRACE(Msg As String)
Debug.Print Msg & vbCrLf
End Sub
