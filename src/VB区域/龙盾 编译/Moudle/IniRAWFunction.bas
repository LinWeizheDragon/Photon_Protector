Attribute VB_Name = "IniRAWFunction"
' 定义API函数
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long


' 为方便使用，把两个API函数封装起来
Function ReadIniFile(filename As String, AppName As String, KeyName As String) As String
Dim temp As String * 100
Dim n As Long
n = GetPrivateProfileString(AppName, KeyName, "", temp, Len(temp), filename)
ReadIniFile = Mid(temp, 1, n)
End Function

Function WriteIniFile(filename As String, AppName As String, KeyName As String, NewKeyName As String)
Dim n As Long
n = WritePrivateProfileString(AppName, KeyName, NewKeyName, filename)
End Function


