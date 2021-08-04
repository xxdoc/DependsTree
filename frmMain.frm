VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   3990
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows の既定値
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  '下揃え
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   3660
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMain 
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTreeviewIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4560
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Map Network Drive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AEE
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C48
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  '上揃え
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CONNECT"
            Object.ToolTipText     =   "データベース接続"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "GREP"
            Object.ToolTipText     =   "オブジェクト検索"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TREE"
            Object.ToolTipText     =   "ツリー表示"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "オブジェクトを開く"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   0
         Width           =   3135
      End
   End
   Begin MSComctlLib.ImageList imlTreeviewIcons 
      Left            =   4560
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D5A
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11AC
            Key             =   "D"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1306
            Key             =   "FN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1758
            Key             =   "K"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B2
            Key             =   "P"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D04
            Key             =   "S"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E5E
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FB8
            Key             =   "V"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル(&F)"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "データベース接続(&C)..."
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEnd 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "表示(&V)"
      Begin VB.Menu mnuViewGrep 
         Caption         =   "オブジェクト検索(&G)..."
      End
      Begin VB.Menu mnuViewBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTree 
         Caption         =   "ツリー表示(&T)"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "オブジェクトを開く(&O)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ヘルプ(&H)"
      Begin VB.Menu mnuHelpVersion 
         Caption         =   "バージョン(&V)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Private Const SW_NORMAL = 1

Private Type ObjectDepends
    ParentID As Long
    ParantName As String
    ChildID As Long
    ChildName As String
    ChildType As String
End Type

Private Const KEY_SEP As String = "_"

Private Const NO_SEL_ITEM As Integer = -1

Private Const TYPE_NOT_OPEN As String = "D,K,S,U"   'Default, Primary, SystemTable, UserTable

Private Const TLB_CONNECT As String = "CONNECT"
Private Const TLB_GREP As String = "GREP"
Private Const TLB_TREE As String = "TREE"
Private Const TLB_OPEN As String = "OPEN"

Private mblnLoad As Boolean

Private Sub Form_Load()
    Me.Caption = App.Title
    mblnLoad = True

    Call Initialize
End Sub

Private Sub Form_Activate()
    If mblnLoad Then
        mblnLoad = False
        Call mnuFileConnect_Click
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    trvMain.Width = Me.ScaleWidth - (trvMain.Left * 2)
    trvMain.Height = Me.ScaleHeight - (stbMain.Height + trvMain.Top)
End Sub

Private Sub mnuFileConnect_Click()
    Call Initialize

    Dim objForm As New clsMsSqlLogon

    If objForm.ShowForm Then
        ptypMsSqlInfo.Server = objForm.Server
        ptypMsSqlInfo.Database = objForm.Database
        ptypMsSqlInfo.Username = objForm.Username
        ptypMsSqlInfo.Password = objForm.Password

        Call ShowStatus(SUCCESS_CONNECT)
    End If

    Set objForm = Nothing
End Sub

Private Sub mnuFileEnd_Click()
    Unload Me
End Sub

Private Sub mnuViewTree_Click()
    Call ClearTree

    If CheckConnect() = False Then
        Call ShowStatus(ERROR_NOT_CONNECT)
        Exit Sub
    End If

    Dim lngID As Long
    Dim strType As String
    If SelType(lngID, strType) = False Then
        Exit Sub
    End If

    Dim strKey As String
    strKey = strType & _
        KEY_SEP & _
        "" & _
        KEY_SEP & _
        CStr(lngID)

    If IsCanTreeNodeAdd(strKey) Then
        trvMain.Nodes.Add _
            , _
            , _
            strKey, _
            txtName.Text, _
            strType
    End If

    If SelRecDeps(lngID, strKey) = False Then
        Exit Sub
    End If

    Call ShowStatus(SUCCESS_OPEN_TREE)
End Sub

Private Sub mnuViewOpen_Click()
    If CheckConnect() = False Then
        Call ShowStatus(ERROR_NOT_CONNECT)
        Exit Sub
    End If

    Dim lngID As Long
    lngID = GetID()

    If lngID = NO_SEL_ITEM Then
        Exit Sub
    End If

    If OpenText(lngID) = False Then
        Exit Sub
    End If

    Call ShowStatus(SUCCESS_OPEN_TEXT)
End Sub

Private Sub mnuViewGrep_Click()
    Call Initialize

    If CheckConnect() = False Then
        Call ShowStatus(ERROR_NOT_CONNECT)
        Exit Sub
    End If

    Dim objForm As New clsObjectGrep

    If objForm.ShowForm Then
        txtName.Text = objForm.ObjectName
        Call mnuViewTree_Click
    End If

    Set objForm = Nothing
End Sub

Private Sub mnuHelpVersion_Click()
    MsgBox "Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case TLB_CONNECT
            Call mnuFileConnect_Click
        Case TLB_GREP
            Call mnuViewGrep_Click
        Case TLB_TREE
            Call mnuViewTree_Click
        Case TLB_OPEN
            Call mnuViewOpen_Click
    End Select
End Sub

Private Sub trvMain_Click()
    If trvMain.SelectedItem Is Nothing Then
        mnuViewOpen.Enabled = False
        tbToolBar.Buttons(TLB_OPEN).Enabled = False
        Exit Sub
    End If

    Dim arrKeyParts As Variant
    arrKeyParts = Split(trvMain.SelectedItem.Key, "_")

    Dim strType As String
    strType = arrKeyParts(0)

    Dim arrNotOpenType As Variant
    arrNotOpenType = Split(TYPE_NOT_OPEN, ",")

    Dim blnExists As Boolean
    blnExists = False

    Dim i As Integer
    For i = 0 To UBound(arrNotOpenType)
        If strType = arrNotOpenType(i) Then
            blnExists = True
            Exit For
        End If
    Next i

    mnuViewOpen.Enabled = Not (blnExists)
    tbToolBar.Buttons(TLB_OPEN).Enabled = Not (blnExists)
End Sub

Private Sub txtName_Change()
    mnuViewTree.Enabled = (txtName.Text <> "")
    tbToolBar.Buttons(TLB_TREE).Enabled = (txtName.Text <> "")
End Sub

Private Sub Initialize()
    Call ClearTree

    txtName.Text = ""

    mnuViewTree.Enabled = False
    tbToolBar.Buttons(TLB_TREE).Enabled = False
End Sub

Private Sub ClearTree()
    trvMain.Nodes.Clear

    Call ShowStatus

    mnuViewOpen.Enabled = False
    tbToolBar.Buttons(TLB_OPEN).Enabled = False
End Sub

Private Sub ShowStatus(Optional ByVal vstrStatus As String = "")
    Const PANEL_STA As Integer = 1
    Const PANEL_SVR As Integer = 2
    Const PANEL_DB As Integer = 3
    Const PANEL_UID As Integer = 4

    stbMain.Panels(PANEL_STA).Text = vstrStatus

    If ptypMsSqlInfo.Server = "" Then
        stbMain.Panels(PANEL_SVR).Text = ERROR_NOT_CONNECT
        stbMain.Panels(PANEL_DB).Text = ""
        stbMain.Panels(PANEL_UID).Text = ""
    Else
        stbMain.Panels(PANEL_SVR).Text = ptypMsSqlInfo.Server
        stbMain.Panels(PANEL_DB).Text = ptypMsSqlInfo.Database
        stbMain.Panels(PANEL_UID).Text = ptypMsSqlInfo.Username
    End If
End Sub

Private Function CheckConnect() As Boolean
    CheckConnect = (ptypMsSqlInfo.Server <> "")
End Function

Private Function SelType(ByRef rlngID As Long, ByRef rstrType As String) As Boolean
    SelType = False

    Dim objCn As New clsMsSqlConnect

    Dim blnRet As Boolean
    blnRet = objCn.Connect( _
        ptypMsSqlInfo.Server, _
        ptypMsSqlInfo.Database, _
        ptypMsSqlInfo.Username, _
        ptypMsSqlInfo.Password _
    )
    If blnRet = False Then
        Call ShowStatus(objCn.Message)
        Set objCn = Nothing
        Exit Function
    End If

    objCn.QueryTimeOut = QUERY_TIMEOUT
    objCn.SetParam 0, txtName.Text

    If objCn.ExecQuery(GetSqlSelType) = False Then
        Call ShowStatus(objCn.Message)
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    If objCn.Recordset.EOF Then
        Call ShowStatus(ERROR_NOT_FOUND_OBJECT)
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    rlngID = objCn.Recordset(0).Value
    rstrType = Trim(objCn.Recordset(1).Value)

    objCn.CloseRs
    objCn.Disconnect
    Set objCn = Nothing

    SelType = True
End Function

Private Function GetSqlSelType() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "    id, "
    strSql = strSql & "    type "
    strSql = strSql & "FROM "
    strSql = strSql & "    sysobjects "
    strSql = strSql & "WHERE "
    strSql = strSql & "    type IN ('C', 'FN', 'P', 'V') "
    strSql = strSql & "AND "
    strSql = strSql & "    name = ? "

    GetSqlSelType = strSql
End Function

Private Function SelDepends(ByVal vlngID As Long, _
                            ByRef rlngCnt As Long, _
                            ByRef rtypDepends() As ObjectDepends) As Boolean
    SelDepends = False

    Dim objCn As New clsMsSqlConnect

    Dim blnRet As Boolean
    blnRet = objCn.Connect( _
        ptypMsSqlInfo.Server, _
        ptypMsSqlInfo.Database, _
        ptypMsSqlInfo.Username, _
        ptypMsSqlInfo.Password _
    )
    If blnRet = False Then
        Call ShowStatus(objCn.Message)
        Set objCn = Nothing
        Exit Function
    End If

    objCn.QueryTimeOut = QUERY_TIMEOUT
    objCn.SetParam 0, vlngID

    If objCn.ExecQuery(GetSqlSelDepends) = False Then
        Call ShowStatus(objCn.Message)
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    rlngCnt = objCn.Recordset.RecordCount

    If StructuredDepends(rtypDepends(), objCn.Recordset) = False Then
        objCn.CloseRs
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    objCn.CloseRs
    objCn.Disconnect
    Set objCn = Nothing

    SelDepends = True
End Function

Private Function GetSqlSelDepends() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT DISTINCT "
    strSql = strSql & "    op.id AS [ParentObjectID], "
    strSql = strSql & "    op.name AS [ParentObjectName], "
    strSql = strSql & "    oc.id AS [ChildObjectID], "
    strSql = strSql & "    oc.name AS [ChildObjectName], "
    strSql = strSql & "    oc.type AS [ChildObjectType] "
    strSql = strSql & "FROM "
    strSql = strSql & "    sysdepends AS d "
    strSql = strSql & "        INNER JOIN sysobjects AS op "
    strSql = strSql & "            ON d.id = op.id "
    strSql = strSql & "        INNER JOIN sysobjects AS oc "
    strSql = strSql & "            ON d.depid = oc.id "
    strSql = strSql & "WHERE "
    strSql = strSql & "    d.id = ? "
    strSql = strSql & "ORDER BY "
    strSql = strSql & "    op.name, "
    strSql = strSql & "    oc.name "

    GetSqlSelDepends = strSql
End Function

Private Function StructuredDepends( _
    ByRef rtypDepends() As ObjectDepends, _
    ByVal vobjRs As Object _
) As Boolean
    StructuredDepends = False
    On Error GoTo Exception

    Dim i As Long
    i = 0
    Do
        If vobjRs.EOF Then
            Exit Do
        End If

        ReDim Preserve rtypDepends(i)
        rtypDepends(i).ParentID = vobjRs(0).Value
        rtypDepends(i).ParantName = Trim(vobjRs(1).Value)
        rtypDepends(i).ChildID = vobjRs(2).Value
        rtypDepends(i).ChildName = Trim(vobjRs(3).Value)
        rtypDepends(i).ChildType = Trim(vobjRs(4).Value)

        i = i + 1

        vobjRs.MoveNext
    Loop

    StructuredDepends = True

    Exit Function

Exception:
    Call ShowStatus(CStr(Err.Number) & ":" & Err.Description)
End Function

Private Function SelRecDeps(ByVal vlngID As Long, ByVal vstrPKey As String) As Boolean
    SelRecDeps = False

    Dim lngCnt As Long
    Dim typDepends() As ObjectDepends

    If SelDepends(vlngID, lngCnt, typDepends()) = False Then
        Exit Function
    End If

    If lngCnt < 1 Then
        SelRecDeps = True
        Exit Function
    End If

    Dim i As Integer
    For i = 0 To UBound(typDepends())
        Dim strKey As String
        strKey = _
            typDepends(i).ChildType & _
            KEY_SEP & _
            CStr(typDepends(i).ParentID) & _
            KEY_SEP & _
            CStr(typDepends(i).ChildID)

        If IsCanTreeNodeAdd(strKey) Then
            trvMain.Nodes.Add _
                vstrPKey, _
                tvwChild, _
                strKey, _
                typDepends(i).ChildName, _
                typDepends(i).ChildType

            Dim lngChildID As Long
            lngChildID = typDepends(i).ChildID

            If SelRecDeps(lngChildID, strKey) = False Then
                Exit For
            End If
        End If
    Next i

    SelRecDeps = True
End Function

Private Function GetID() As Long
    GetID = NO_SEL_ITEM

    If trvMain.Nodes.Count = 0 Then
        Exit Function
    End If

    Dim blnSel As Boolean
    blnSel = False

    Dim i As Long
    For i = 1 To trvMain.Nodes.Count
        If trvMain.Nodes(i).Selected Then
            blnSel = True
            Exit For
        End If
    Next i

    If blnSel = False Then
        Exit Function
    End If

    Dim strKey As String
    strKey = trvMain.Nodes(i).Key

    Dim intSepPos As Integer
    intSepPos = InStr(1, strKey, KEY_SEP)

    Dim strType As String
    strType = Left(strKey, intSepPos - 1)

    Dim arrNotOpenType As Variant
    arrNotOpenType = Split(TYPE_NOT_OPEN, ",")

    Dim j As Integer
    For j = 0 To UBound(arrNotOpenType)
        If strType = arrNotOpenType(j) Then
            Exit Function
        End If
    Next j

    intSepPos = InStr(intSepPos + 1, strKey, KEY_SEP)

    Dim strID As String
    strID = Right(strKey, Len(strKey) - intSepPos)

    GetID = CLng(strID)
End Function

Private Function GetSelText() As String
    GetSelText = ""

    If trvMain.Nodes.Count = 0 Then
        Exit Function
    End If

    Dim blnSel As Boolean
    blnSel = False

    Dim i As Long
    For i = 1 To trvMain.Nodes.Count
        If trvMain.Nodes(i).Selected Then
            blnSel = True
            Exit For
        End If
    Next i

    If blnSel = False Then
        Exit Function
    End If

    GetSelText = trvMain.Nodes(i).Text
End Function

Private Function OpenText(ByVal vlngID As Long) As Boolean
    OpenText = False

    Dim strText As String
    If SelText(vlngID, strText) = False Then
        Exit Function
    End If

    Dim strPath As String
    If ExportText(strText, strPath) = False Then
        Exit Function
    End If

    If OpenFile(strPath) = False Then
        Exit Function
    End If

    OpenText = True
End Function

Private Function SelText(ByVal vlngID As Long, ByRef rstrText As String) As Boolean
    SelText = False

    rstrText = ""

    Dim objCn As New clsMsSqlConnect

    Dim blnRet As Boolean
    blnRet = objCn.Connect( _
        ptypMsSqlInfo.Server, _
        ptypMsSqlInfo.Database, _
        ptypMsSqlInfo.Username, _
        ptypMsSqlInfo.Password _
    )
    If blnRet = False Then
        Call ShowStatus(objCn.Message)
        Set objCn = Nothing
        Exit Function
    End If

    objCn.QueryTimeOut = QUERY_TIMEOUT
    objCn.SetParam 0, vlngID

    If objCn.ExecQuery(GetSqlSelText) = False Then
        Call ShowStatus(objCn.Message)
        objCn.Disconnect
        Set objCn = Nothing
        Exit Function
    End If

    If objCn.Recordset.EOF = False Then
        rstrText = objCn.Recordset(0).Value
    End If

    objCn.CloseRs
    objCn.Disconnect
    Set objCn = Nothing

    SelText = True
End Function

Private Function GetSqlSelText() As String
    GetSqlSelText = ""

    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "    text "
    strSql = strSql & "FROM "
    strSql = strSql & "    syscomments "
    strSql = strSql & "WHERE "
    strSql = strSql & "    id = ? "

    GetSqlSelText = strSql
End Function

Private Function ExportText(ByVal vstrText As String, ByRef rstrPath As String) As Boolean
    ExportText = False

    On Error GoTo Exception

    Dim strTmpDir As String
    strTmpDir = Environ("TEMP")

    Dim strFileNm As String
    strFileNm = GetSelText & ".sql"

    rstrPath = strTmpDir & "\" & strFileNm

    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")

    Dim objFile As Object
    Set objFile = objFso.OpenTextFile(rstrPath, 2, True)
    objFile.WriteLine vstrText
    Set objFile = Nothing

    ExportText = True

    Exit Function

Exception:
    Call ShowStatus(CStr(Err.Number) & ":" & Err.Description)
End Function

Private Function OpenFile(ByVal vstrPath As String) As Boolean
    OpenFile = False

    Dim lngRet As Long
    lngRet = ShellExecute(Me.hwnd, "open", vstrPath, "", "", SW_NORMAL)

    If lngRet < 33 Then
        Exit Function
    End If

    OpenFile = True
End Function

Private Function IsCanTreeNodeAdd(ByVal vstrKey As String) As Boolean
    IsCanTreeNodeAdd = False

    Dim i As Integer
    For i = 1 To trvMain.Nodes.Count
        If vstrKey = trvMain.Nodes.Item(i).Key Then
            Exit Function
        End If
    Next i

    IsCanTreeNodeAdd = True
End Function
