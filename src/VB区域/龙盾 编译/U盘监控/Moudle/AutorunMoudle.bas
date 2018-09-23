Attribute VB_Name = "AutorunMoudle"
Option Explicit
Public pathini As String
Public mark As Integer '控制是否自启动的标志变量（1自启动，0不自启动）

'添加删除自启动项目的API函数声明

Public Const HKEY_CLASSES_ROOT = &H80000000

Public Const HKEY_CURRENT_USER = &H80000001

Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const HKEY_USERS = &H80000003

Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_NONE = 0

Public Const REG_SZ = 1

Public Const REG_EXPAND_SZ = 2

Public Const REG_BINARY = 3

Public Const REG_DWORD = 4

Public Const REG_DWORD_BIG_ENDIAN = 5

Public Const REG_MULTI_SZ = 7

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'在注册表中添加删除自启动项目的模块

