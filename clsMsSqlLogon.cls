VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsSqlLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrServer As String
Private mstrDatabase As String
Private mstrUsername As String
Private mstrPassword As String

Public Property Get Server() As String
    Server = mstrServer
End Property
Private Property Let Server(ByVal vstrSvr As String)
    mstrServer = vstrSvr
End Property

Public Property Get Database() As String
    Database = mstrDatabase
End Property
Private Property Let Database(ByVal vstrDb As String)
    mstrDatabase = vstrDb
End Property

Public Property Get Username() As String
    Username = mstrUsername
End Property
Private Property Let Username(ByVal vstrUid As String)
    mstrUsername = vstrUid
End Property

Public Property Get Password() As String
    Password = mstrPassword
End Property
Private Property Let Password(ByVal vstrPwd As String)
    mstrPassword = vstrPwd
End Property

Private Sub Class_Initialize()
    Server = ""
    Database = ""
    Username = ""
    Password = ""
End Sub

Public Function ShowForm() As Boolean
    ShowForm = False

    Dim objForm As New frmMsSqlLogon

    objForm.Show vbModal

    If objForm.ReturnDialog <> vbOK Then
        GoTo Exit_Label
    End If

    Server = objForm.txtSvr.Text
    Database = objForm.txtDb.Text
    Username = objForm.txtUid.Text
    Password = objForm.txtPwd.Text

    ShowForm = True

Exit_Label:
    Unload objForm
    Set objForm = Nothing
End Function
