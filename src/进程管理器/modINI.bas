Attribute VB_Name = "modINI"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SaveINIS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long


Sub Saveini(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, Optional ByVal lplFileName As String = "")
    If lplFileName = "" Then lplFileName = Setini
    SaveINIS lpApplicationName, lpKeyName, lpString, lplFileName
End Sub

'获取INI中的值
Function GetINI(ByVal AppName As String, ByVal KeyName As String, ByVal FileName As String) As String
    Dim RetStr As String
    RetStr = String(255, Chr(0))
    GetINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName))
    GetINI = CheckStr(GetINI)
End Function

'获取INI中的值，但返回Long类型
Function GetLongINI(AppName As String, KeyName As String, Optional ByVal ReturnVal As Long = 0, Optional ByVal FileName As String = "") As Long '获取INI中整数值
On Error GoTo aaaa
    Dim RetStr As String, Str As String
    If FileName = "" Then FileName = Setini
    RetStr = String(255, Chr(0))
    Str = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName))
    If Str = "" Then
        GetLongINI = ReturnVal
    Else
        GetLongINI = CLng(Str)
    End If
Exit Function
aaaa:
    GetLongINI = 0
End Function
