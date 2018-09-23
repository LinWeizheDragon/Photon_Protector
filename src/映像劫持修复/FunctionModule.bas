Attribute VB_Name = "FunctionModule"
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
lpcbData As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
"RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long _
' Note that if you declare the lpData parameter as String, _
you must pass it By Value.
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long _
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_DWORD = 4
Public Const REG_SIZE As Long = 1024
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public YW_Regedit_Hkey As Long
Public YW_Regedit_Focus2 As String
Public Sub YW_Regedit_Scanning()
On Error GoTo YW_Regedit_ScanningRegedit_Error
'声明变量
Dim YW_Regedit_ScanningRegedit1 As Long
Dim YW_Regedit_ScanningRegedit3 As Long
Dim YW_Regedit_ScanningRegedit4 As String
Dim YW_Regedit_ScanningRegedit5 As Long
Dim YW_Regedit_ScanningRegedit6 As String
Dim YW_Regedit_ScanningRegedit7 As Long
Dim YW_Regedit_ScanningRegedit8 As Byte
Dim YW_Regedit_Scanning_Hkey2 As Long
'清空ListBox
MainForm.List1.Clear
'扫描镜像劫持:
YW_Regedit_ScanningRegedit4 = Space(REG_SIZE) '分配变量空间
YW_Regedit_ScanningRegedit6 = Space(REG_SIZE) '分配变量空间
'=====扫描项代码开始=====
While RegEnumKey(YW_Regedit_Hkey, YW_Regedit_ScanningRegedit3, _
YW_Regedit_ScanningRegedit4, REG_SIZE) = 0
YW_Regedit_ScanningRegedit3 = YW_Regedit_ScanningRegedit3 + 1
YW_Regedit_ScanningRegedit1 = RegOpenKey(YW_Regedit_Hkey, _
YW_Regedit_ScanningRegedit4, YW_Regedit_Scanning_Hkey2)
'-----扫描键值代码开始-----
While RegEnumValue(YW_Regedit_Scanning_Hkey2, YW_Regedit_ScanningRegedit5, _
YW_Regedit_ScanningRegedit6, REG_SIZE, ByVal 0&, YW_Regedit_ScanningRegedit7, _
ByVal YW_Regedit_ScanningRegedit8, REG_SIZE) = 0
YW_Regedit_ScanningRegedit5 = YW_Regedit_ScanningRegedit5 + 1
'V  YW_Regedit_ScanningRegedit4与YW_Regedit_ScanningRegedit6变量无法直接判断
MainForm.Label1.Caption = YW_Regedit_ScanningRegedit6
If MainForm.Label1.Caption = "Debugger" Then
MainForm.Label1.Caption = YW_Regedit_ScanningRegedit4
If MainForm.Label1.Caption <> "Your Image File Name Here without a path" Then
MainForm.List1.AddItem YW_Regedit_ScanningRegedit4
End If
MainForm.Label1.Caption = ""
End If
MainForm.Label1.Caption = ""
Wend
'-----扫描键值代码结束-----
YW_Regedit_ScanningRegedit5 = 0 '<-鄙视这个
YW_Regedit_ScanningRegedit1 = RegCloseKey(YW_Regedit_Scanning_Hkey2)
Wend
'=====扫描项代码结束=====
'镜像劫持扫描结束
Exit Sub
'错误处理
YW_Regedit_ScanningRegedit_Error:
MsgBox ("扫描注册表时发生错误!")
End Sub
