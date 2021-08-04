Attribute VB_Name = "basMain"
Option Explicit

Public Const INI_PATH As String = "STATUS.INI"
Public Const INI_SEC_CONNECT As String = "CONNECT"
Public Const INI_KEY_SVR As String = "SERVER"
Public Const INI_KEY_DB As String = "DATABASE"
Public Const INI_KEY_UID As String = "USERNAME"
Public Const INI_SEC_GREP As String = "GREP"
Public Const INI_KEY_GREP As String = "CHAR"

Public Type MsSqlInfo
    Server As String
    Database As String
    Username As String
    Password As String
End Type
Public ptypMsSqlInfo As MsSqlInfo

Public Const QUERY_TIMEOUT As Long = 3000

Sub Main()
    frmMain.Show
End Sub
