Attribute VB_Name = "开机运行"
''''开机运行
  Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
  Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long                                                                   '   Note   that   if   you   declare   the   lpData   parameter   as   String,   you   must   pass   it   By   Value.
  Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
  Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
  Public Const REG_SZ = 1
  Public Const REG_DWORD = 4
  Public Function AddToStarup(DesName As String, exePath As String) As Boolean
  Dim SubKey     As String
  Dim Hkey     As Long
  On Error GoTo acd
  AddToStarup = False
    '' "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\GKWebTool", dirwin
  SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
'HKEY_CURRENT_USER
  RegCreateKey HKEY_LOCAL_MACHINE, SubKey, Hkey

  RegSetValueEx Hkey, DesName, 0, REG_SZ, ByVal exePath, LenB(StrConv(exePath, vbFromUnicode)) + 1

  RegCloseKey Hkey

  AddToStarup = True
  Exit Function
acd:
  AddToStarup = False
  End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long

r = RegOpenKey(Hkey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
r = RegCloseKey(keyhand)
End Function

Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function
Public Sub Savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub
