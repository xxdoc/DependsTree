VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsSqlConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrMessage As String
Private mobjConnection As Object
Private mobjRecordset As Object
Private mlngTimeout As Long
Private mintParamCounter As Integer
Private mvarCommandParam() As Variant

Public Property Get Message() As String
    Message = mstrMessage
End Property
Private Property Let Message(ByVal vstrMsg As String)
    mstrMessage = vstrMsg
End Property

Private Property Get Connection() As Object
    Set Connection = mobjConnection
End Property
Private Property Let Connection(ByVal vobjCn As Object)
    Set mobjConnection = vobjCn
End Property

Public Property Get Recordset() As Object
    Set Recordset = mobjRecordset
End Property
Public Property Let Recordset(ByVal vobjRs As Object)
    Set mobjRecordset = vobjRs
End Property

Public Property Get QueryTimeOut() As Long
    QueryTimeOut = mlngTimeout
End Property
Public Property Let QueryTimeOut(ByVal vlngTimeOut As Long)
    mlngTimeout = vlngTimeOut
End Property

Private Property Get ParamCounter() As Integer
    ParamCounter = mintParamCounter
End Property
Private Property Let ParamCounter(ByVal vintCnt As Integer)
    mintParamCounter = vintCnt
End Property

Private Sub Class_Initialize()
    Message = ""

    Set mobjConnection = CreateObject("ADODB.Connection")
    Set mobjRecordset = CreateObject("ADODB.Recordset")

    QueryTimeOut = QUERY_TIMEOUT
    ParamCounter = 0
End Sub

Private Sub Class_Terminate()
    Set mobjRecordset = Nothing
    Set mobjConnection = Nothing
End Sub

Public Function Connect(ByVal vstrSvr As String, _
                        ByVal vstrDb As String, _
                        ByVal vstrUid As String, _
                        ByVal vstrPwd As String) As Boolean
    Connect = False

    On Error GoTo Exception

    Dim strCn
    strCn = "driver=SQL Server" & ";" _
          & "server=" & vstrSvr & ";" _
          & "database=" & vstrDb & ";" _
          & "uid=" & vstrUid & ";" _
          & "pwd=" & vstrPwd & ";"

    Connection.ConnectionString = strCn
    Connection.Open
    Connection.CursorLocation = 3

    On Error GoTo 0

    Connect = True

    Exit Function

Exception:
    Message = CStr(Err.Number) & ":" & Err.Description
End Function

Public Sub Disconnect()
    Connection.Close
End Sub

Public Sub CloseRs()
    Recordset.Close
End Sub

Public Function ExecQuery(ByVal vstrSql As String) As Boolean
    ExecQuery = False

    On Error GoTo Exception

    Dim objCmd As Object
    Set objCmd = CreateObject("ADODB.Command")

    objCmd.ActiveConnection = Connection
    objCmd.CommandTimeOut = QueryTimeOut
    objCmd.CommandText = vstrSql

    If ParamCounter > 0 Then
        Dim i As Integer
        For i = 0 To UBound(mvarCommandParam())
            objCmd.Parameters(i) = mvarCommandParam(i)
        Next i
    End If

    Recordset.Open objCmd

    On Error GoTo 0

    ExecQuery = True

    Exit Function

Exception:
    Message = CStr(Err.Number) & ":" & Err.Description
End Function

Public Sub SetParam(ByVal vintIndex As Integer, ByVal vvarValue As Variant)
    ReDim Preserve mvarCommandParam(vintIndex)
    mvarCommandParam(vintIndex) = vvarValue

    ParamCounter = ParamCounter + 1
End Sub
