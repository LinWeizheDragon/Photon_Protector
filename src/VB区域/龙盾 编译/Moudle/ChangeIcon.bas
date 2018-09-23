Attribute VB_Name = "ChangeIcon"
Option Explicit
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
'（以上函数是一些注册表的常量，用来定义 hKey）

Enum ValueType
REG_NONE = 0
REG_SZ = 1
REG_EXPAND_SZ = 2
REG_BINARY = 3
REG_DWORD = 4
REG_DWORD_BIG_ENDIAN = 5
REG_MULTI_SZ = 7
End Enum

'（这个枚举是用来定义 dwType）

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long   '这个函数是用来创建注册表的主键
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long   '这个函数用来关闭打开的注册表
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long '这个函数用来改写注册表的键值





Public Function ChangeFile(ByVal Name As String, ByVal ShellPath As String, ByVal IconPath As String)
Dim ret As Long, hKey As Long, ExePath As String
ret = RegCreateKey(HKEY_CLASSES_ROOT, "." & Name, hKey) '定义 .abc文件
ret = RegSetValue(HKEY_CLASSES_ROOT, "." & Name, REG_SZ, Name & "file", Len(Name & "file") + 1) '定义文件的类型,注意最后一个数字，它是 "userfile"的字节数 + 1
ret = RegCreateKey(HKEY_CLASSES_ROOT, Name & "file", hKey)          '定义"userfile"
ret = RegCreateKey(HKEY_CLASSES_ROOT, Name & "file\shell", hKey)      '定义它的操作
ret = RegCreateKey(HKEY_CLASSES_ROOT, Name & "file\shell\open", hKey) '具体定义操作的名称
ret = RegCreateKey(HKEY_CLASSES_ROOT, Name & "file\shell\open\command", hKey) '定义操作的动作
If ShellPath <> "" Then
ExePath = ShellPath        '获得VB程序名称
ret = RegSetValue(HKEY_CLASSES_ROOT, Name & "file\shell\open\command", REG_SZ, ExePath, LenB(StrConv(ExePath, vbFromUnicode)) + 1)
'最关键的一步！将 "userfile" 的打开（open)操作和我们的程序关联起来
          
End If
Dim iconfile
ret = RegCreateKey(HKEY_CLASSES_ROOT, Name & "file\DefaultIcon", hKey)
iconfile = IconPath
ret = RegSetValue(HKEY_CLASSES_ROOT, Name & "file\DefaultIcon", REG_SZ, iconfile, LenB(StrConv(iconfile, vbFromUnicode)) + 1)

RegCloseKey hKey
End Function
