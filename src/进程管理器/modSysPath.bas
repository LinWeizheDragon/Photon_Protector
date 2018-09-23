Attribute VB_Name = "modSysPath"
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'获得系统system32目录
Public Function GetSysDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetSystemDirectory(temp, Len(temp))
    GetSysDir = Left$(temp, x)
End Function

'获得Win目录
Public Function GetWinDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetWindowsDirectory(temp, Len(temp))
    GetWinDir = Left$(temp, x)
End Function

