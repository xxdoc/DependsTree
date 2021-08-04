VERSION 5.00
Begin VB.Form frmObjectGrep 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "オブジェクト検索"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox lstGrep 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton cmdGrep 
      Caption         =   "検索"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtChar 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmObjectGrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintReturnDialog As Integer
Private mstrObjectName As String

Public Property Get ReturnDialog() As Integer
    ReturnDialog = mintReturnDialog
End Property
Private Property Let ReturnDialog(ByVal vintRet As Integer)
    mintReturnDialog = vintRet
End Property

Public Property Get ObjectName() As String
    ObjectName = mstrObjectName
End Property
Private Property Let ObjectName(ByVal vstrName As String)
    mstrObjectName = vstrName
End Property

Private Sub Form_Load()
    ReturnDialog = vbCancel
    ObjectName = ""

    Call ReadGrepChar
End Sub

Private Sub cmdGrep_Click()
    If CheckValue() = False Then
        Exit Sub
    End If

    lstGrep.Clear

    If GrepObject() = False Then
        Exit Sub
    End If
End Sub

Private Sub cmdOK_Click()
    Dim blnSelItem As Boolean
    blnSelItem = False

    Dim i As Long
    For i = 0 To lstGrep.ListCount - 1
        If lstGrep.Selected(i) Then
            blnSelItem = True
            Exit For
        End If
    Next i

    If blnSelItem = False Then
        MsgBox ERROR_NOT_SELECT, vbOKOnly + vbExclamation, App.Title
        Exit Sub
    End If

    ObjectName = lstGrep.Text

    Call WriteGrepChar

    ReturnDialog = vbOK

    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Function CheckValue() As Boolean
    CheckValue = False

    If txtChar.Text = "" Then
        MsgBox ERROR_NOT_INPUT, vbOKOnly + vbExclamation, App.Title
        txtChar.SetFocus
        Exit Function
    End If

    CheckValue = True
End Function

Private Function GrepObject() As Boolean
    GrepObject = False

    Dim objCn As New clsMsSqlConnect

    Dim blnRet As Boolean
    blnRet = objCn.Connect( _
        ptypMsSqlInfo.Server, _
        ptypMsSqlInfo.Database, _
        ptypMsSqlInfo.Username, _
        ptypMsSqlInfo.Password _
    )
    If blnRet = False Then
        MsgBox objCn.Message, vbOKOnly + vbCritical, App.Title
        Set objCn = Nothing
        Exit Function
    End If

    objCn.QueryTimeOut = QUERY_TIMEOUT
'    objCn.SetParam 0, txtChar.Text  'TODO:

    If objCn.ExecQuery(GetGrepSql) = False Then
        MsgBox objCn.Message, vbOKOnly + vbCritical, App.Title
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    If objCn.Recordset.EOF Then
        MsgBox ERROR_NOT_FOUND_OBJECT, vbOKOnly + vbExclamation, App.Title
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    Call SetList(objCn.Recordset)

    objCn.CloseRs
    objCn.Disconnect
    Set objCn = Nothing

    GrepObject = True
End Function

'TODO:
'Private Function GetGrepSql() As String
'    Dim strSql As String
'    strSql = ""
'
'    strSql = strSql & "SELECT "
'    strSql = strSql & "    name "
'    strSql = strSql & "FROM "
'    strSql = strSql & "    sysobjects "
'    strSql = strSql & "WHERE "
'    strSql = strSql & "    id IN ( "
'    strSql = strSql & "        SELECT "
'    strSql = strSql & "            id "
'    strSql = strSql & "        FROM "
'    strSql = strSql & "            syscomments "
'    strSql = strSql & "        WHERE "
'    strSql = strSql & "            CHARINDEX(?, text) > 0 "
'    strSql = strSql & "    ) "
'    strSql = strSql & "AND "
'    strSql = strSql & "    type = 'P' "
'    strSql = strSql & "ORDER BY "
'    strSql = strSql & "    name "
'
'    GetGrepSql = strSql
'End Function

Private Function GetGrepSql() As String
    Dim strSql As String
    strSql = ""

    Dim strChar As String
    strChar = txtChar.Text
    strChar = Replace(strChar, "--", "")
    strChar = Replace(strChar, "/*", "")
    strChar = Replace(strChar, "'", "")

    strSql = strSql & "SELECT "
    strSql = strSql & "    name "
    strSql = strSql & "FROM "
    strSql = strSql & "    sysobjects "
    strSql = strSql & "WHERE "
    strSql = strSql & "    id IN ( "
    strSql = strSql & "        SELECT "
    strSql = strSql & "            id "
    strSql = strSql & "        FROM "
    strSql = strSql & "            syscomments "
    strSql = strSql & "        WHERE "
    strSql = strSql & "            CHARINDEX('" & strChar & "', text) > 0 "
    strSql = strSql & "    ) "
    strSql = strSql & "ORDER BY "
    strSql = strSql & "    name "

    GetGrepSql = strSql
End Function

Private Sub SetList(ByVal vobjRs As Object)
    Do
        If vobjRs.EOF Then
            Exit Do
        End If
        lstGrep.AddItem vobjRs(0).Value
        vobjRs.MoveNext
    Loop
End Sub

Private Sub ReadGrepChar()
    txtChar.Text = GetIniValue(INI_KEY_GREP, INI_SEC_GREP, INI_PATH)
End Sub

Private Sub WriteGrepChar()
    Call SetIniValue(txtChar.Text, INI_KEY_GREP, INI_SEC_GREP, INI_PATH)
End Sub
