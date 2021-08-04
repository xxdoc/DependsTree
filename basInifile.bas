Attribute VB_Name = "basInifile"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long

Public Function GetIniValue(ByVal vstrKey As String, _
                            ByVal vstrSection As String, _
                            ByVal vstrPath As String) As String
    Dim strValue As String * 255
    Call GetPrivateProfileString(vstrSection, vstrKey, "", strValue, Len(strValue), vstrPath)

    GetIniValue = Left(strValue, InStr(1, strValue, vbNullChar) - 1)
End Function

Public Function SetIniValue(ByVal vstrValue As String, _
                            ByVal vstrKey As String, _
                            ByVal vstrSection As String, _
                            ByVal vstrPath As String) As Boolean
    Dim lngRet As Long
    lngRet = WritePrivateProfileString(vstrSection, vstrKey, vstrValue, vstrPath)

    SetIniValue = CBool(lngRet)
End Function
