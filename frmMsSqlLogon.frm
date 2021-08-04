VERSION 5.00
Begin VB.Form frmMsSqlLogon 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "SQL Server接続"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
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
   ScaleHeight     =   2340
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "キャンセル"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "接続"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtUid 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtDb 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtSvr 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      Caption         =   "パスワード:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "ユーザー:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "データベース:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblSvr 
      Alignment       =   1  '右揃え
      Caption         =   "サーバー:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMsSqlLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintReturnDialog As Integer

Public Property Get ReturnDialog() As Integer
    ReturnDialog = mintReturnDialog
End Property
Private Property Let ReturnDialog(ByVal vintRet As Integer)
    mintReturnDialog = vintRet
End Property

Private Sub Form_Load()
    ReturnDialog = vbCancel

    Call ReadConnectInfo
End Sub

Private Sub Form_Activate()
    Call SetDefaultFocus
End Sub

Private Sub cmdConnect_Click()
    If ConnectTest() = False Then
        Exit Sub
    End If

    Call WriteConnectInfo

    ReturnDialog = vbOK

    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Function ConnectTest() As Boolean
    Dim objCn As New clsMsSqlConnect

    If objCn.Connect(txtSvr.Text, txtDb.Text, txtUid.Text, txtPwd.Text) = False Then
        MsgBox objCn.Message, vbOKOnly + vbExclamation
        ConnectTest = False
    Else
        objCn.Disconnect
        ConnectTest = True
    End If

    Set objCn = Nothing
End Function

Private Sub ReadConnectInfo()
    txtSvr.Text = GetIniValue(INI_KEY_SVR, INI_SEC_CONNECT, INI_PATH)
    txtDb.Text = GetIniValue(INI_KEY_DB, INI_SEC_CONNECT, INI_PATH)
    txtUid.Text = GetIniValue(INI_KEY_UID, INI_SEC_CONNECT, INI_PATH)
End Sub

Private Sub WriteConnectInfo()
    Call SetIniValue(txtSvr.Text, INI_KEY_SVR, INI_SEC_CONNECT, INI_PATH)
    Call SetIniValue(txtDb.Text, INI_KEY_DB, INI_SEC_CONNECT, INI_PATH)
    Call SetIniValue(txtUid.Text, INI_KEY_UID, INI_SEC_CONNECT, INI_PATH)
End Sub

Private Sub SetDefaultFocus()
    If txtSvr.Text = "" Then
        txtSvr.SetFocus
    Else
        txtPwd.SetFocus
    End If
End Sub
