VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectGrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrObjectName As String

Public Property Get ObjectName() As String
    ObjectName = mstrObjectName
End Property
Private Property Let ObjectName(ByVal vstrName As String)
    mstrObjectName = vstrName
End Property

Private Sub Class_Initialize()
    mstrObjectName = ""
End Sub

Public Function ShowForm() As Boolean
    ShowForm = False

    Dim objForm As New frmObjectGrep

    objForm.Show vbModal

    If objForm.ReturnDialog <> vbOK Then
        GoTo Exit_Label
    End If

    ObjectName = objForm.ObjectName

    ShowForm = True

Exit_Label:
    Unload objForm
    Set objForm = Nothing
End Function
