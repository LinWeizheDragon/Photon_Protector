Attribute VB_Name = "IniRAWFunction"
' 定义API函数
'写入到配置文件中去
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'获取配置文件中的值
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



'读ini文件
 Public Function ReadString(ByVal Caption As String, ByVal item As String, ByVal Path As String) As String
    On Error Resume Next
    Dim sBuffer As String
    sBuffer = Space(32767)
    GetPrivateProfileString Caption, item, vbNullString, sBuffer, 32766, Path
    ReadString = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
 End Function

'写ini文件
 Public Function WriteString(ByVal Caption As String, ByVal item As String, ByVal ItemValue As String, ByVal Path As String) As Long
    Dim sBuffer As String
    sBuffer = Space(32766)
    sBuffer = ItemValue & vbNullChar
    WriteString = WritePrivateProfileString(Caption, item, sBuffer, Path)
 End Function
