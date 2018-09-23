

Attribute VB_Name = "modIni"
Option Explicit
'''''''''''''''''''''''''
'读写INI文件模块
'''''''''''''''''''''''''
Private Declare Function GetPrivateProfileSection Lib "KERNEL32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, lpString As Any, ByVal lpFileName As String) As Long
Public strIniFilePath As String '设置文件路径

'读取指定节点下对应名称的值
Public Function GetiniValue(ByVal lpKeyName As String, ByVal strName As String, ByVal strIniFile As String) As String
    Dim strTmp As String * 32767
    Call GetPrivateProfileString(lpKeyName, strName, "", strTmp, Len(strTmp), strIniFile)
    GetiniValue = Left$(strTmp, InStr(strTmp, vbNullChar) - 1)
End Function

'给指定节点下对名称赋值
Public Function WriteIniStr(ByVal strSection As String, ByVal strKey As String, ByVal strData As String, ByVal strIniFile As String) As Boolean
    On Error GoTo WriteIniStrErr
    WriteIniStr = True
    If strData = "0" Then
        WritePrivateProfileString strSection, strKey, ByVal 0, strIniFile
    Else
        WritePrivateProfileString strSection, strKey, ByVal strData, strIniFile
    End If
    Exit Function
WriteIniStrErr:
    err.Clear
    WriteIniStr = False
End Function

'获取指定节电下的最大索引
Public Function GetMaxIndex(ByVal strSection As String, strIniFile As String) As String
    Dim strReturn As String * 32767
    Dim strTmp As String
    Dim lngReturn As Integer, i As Integer, strTmpArray() As String, sum As Integer
    lngReturn = GetPrivateProfileSection(strSection, strReturn, Len(strReturn), strIniFile)
    strTmp = Left(strReturn, lngReturn)
    strTmpArray = Split(strTmp, Chr(0))
    For i = 0 To UBound(strTmpArray)
        If strTmpArray(i) <> "" And strTmpArray(i) <> Chr(0) Then
            strTmp = Left(strTmpArray(i), InStr(strTmpArray(i), "=") - 1)
            If Val(strTmp) > sum Then sum = Val(strTmp)
        End If
    Next
    GetMaxIndex = sum + 1
End Function

'判断数据是否已经添加过了
Public Function IsIniDataExist(ByVal strSection As String, ByVal strData As String, ByVal strIniFile As String) As String
    Dim strReturn As String * 32767
    Dim strTmp As String
    Dim lngReturn As Integer, i As Integer, strTmpArray() As String, sum As Integer
    lngReturn = GetPrivateProfileSection(strSection, strReturn, Len(strReturn), strIniFile)
    strTmp = Left(strReturn, lngReturn)
    strTmpArray = Split(strTmp, Chr(0))
    For i = 0 To UBound(strTmpArray)
        If strTmpArray(i) <> "" And strTmpArray(i) <> Chr(0) Then
            strTmp = Trim(Mid(strTmpArray(i), InStr(strTmpArray(i), "=") + 1, Len(strTmpArray(i)) - InStr(strTmpArray(i), "=")))
            If strTmp <> "" Then
                If LCase(strTmp) = LCase(strData) Then
                    IsIniDataExist = Left(strTmpArray(i), InStr(strTmpArray(i), "=") - 1)
                    Exit Function
                End If
            End If
        End If
    Next
End Function
