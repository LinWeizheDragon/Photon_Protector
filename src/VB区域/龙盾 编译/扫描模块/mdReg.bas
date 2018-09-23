Attribute VB_Name = "mdReg"
Option Explicit
'---------------------------------------------------------------
'- 注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0        ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

' 有关导入/导出的常量
Const REG_FORCE_RESTORE As Long = 8&
Const TOKEN_QUERY As Long = &H8&
Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Const SE_PRIVILEGE_ENABLED As Long = &H2
Const SE_RESTORE_NAME = "SeRestorePrivilege"
Const SE_BACKUP_NAME = "SeBackupPrivilege"

'---------------------------------------------------------------
'- 注册表类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

'---------------------------------------------------------------
'- 自定义枚举类型...
'---------------------------------------------------------------
' 注册表数据类型...
Public Enum ValueType
    REG_SZ = 1                         ' 字符串值
    REG_EXPAND_SZ = 2                  ' 可扩充字符串值
    REG_BINARY = 3                     ' 二进制值
    REG_DWORD = 4                      ' DWORD值
    REG_MULTI_SZ = 7                   ' 多字符串值
End Enum

' 注册表关键字根类型...
Public Enum keyRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public strstring As String
Private hKey As Long                   ' 注册表打开项的句柄
Private i As Long, j As Long           ' 循环变量
Private Success As Long                ' API函数的返回值, 判断函数调用是否成功

'-------------------------------------------------------------------------------------------------------------
'- 新建注册表关键字并设置注册表关键字的值...
'- 如果 ValueName 和 Value 都缺省, 则只新建 KeyName 空项, 无子键...
'- 如果只缺省 ValueName 则将设置指定 KeyName 的默认值
'- 参数说明: KeyRoot--根类型, KeyName--子项名称, ValueName--值项名称, Value--值项数据, ValueType--值项类型
'-------------------------------------------------------------------------------------------------------------
Public Function SetKeyValue(keyRoot As keyRoot, KeyName As String, Optional ValueName As String, Optional Value As Variant = "", Optional ValueType As ValueType = REG_SZ) As Boolean
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...
    
    ' 新建注册表关键字...
    Success = RegCreateKeyEx(keyRoot, KeyName, 0, ValueType, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, 0)
    If Success <> ERROR_SUCCESS Then SetKeyValue = False: RegCloseKey hKey: Exit Function
    
    ' 设置注册表关键字的值...
    If IsMissing(ValueName) = False Then
        Select Case ValueType
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                Success = RegSetValueEx(hKey, ValueName, 0, ValueType, ByVal CStr(Value), LenB(StrConv(Value, vbFromUnicode)) + 1)
            Case REG_DWORD
                If CDbl(Value) <= 4294967295# And CDbl(Value) >= 0 Then
                    Dim sValue As String
                    sValue = DoubleToHex(Value)
                    Dim dValue(3) As Byte
                    dValue(0) = Format("&h" & Mid(sValue, 7, 2))
                    dValue(1) = Format("&h" & Mid(sValue, 5, 2))
                    dValue(2) = Format("&h" & Mid(sValue, 3, 2))
                    dValue(3) = Format("&h" & Mid(sValue, 1, 2))
                    Success = RegSetValueEx(hKey, ValueName, 0, ValueType, dValue(0), 4)
                Else
                    Success = ERROR_BADKEY
                End If
            Case REG_BINARY
                On Error Resume Next
                Success = 1                             ' 假设调用API不成功(成功返回0)
                ReDim tmpValue(UBound(Value)) As Byte
                For i = 0 To UBound(tmpValue)
                    tmpValue(i) = Value(i)
                Next i
                Success = RegSetValueEx(hKey, ValueName, 0, ValueType, tmpValue(0), UBound(Value) + 1)
        End Select
    End If
    If Success <> ERROR_SUCCESS Then SetKeyValue = False: RegCloseKey hKey: Exit Function
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
    SetKeyValue = True                                       ' 返回函数值
End Function

'-------------------------------------------------------------------------------------------------------------
'- 获得已存在的注册表关键字的值...
'- 如果 ValueName="" 则返回 KeyName 项的默认值...
'- 如果指定的注册表关键字不存在, 则返回空串...
'- 参数说明: KeyRoot--根类型, KeyName--子项名称, ValueName--值项名称, ValueType--值项类型
'-------------------------------------------------------------------------------------------------------------
Public Function GetKeyValue(ByVal keyRoot As keyRoot, ByVal KeyName As String, ByVal ValueName As String, Optional ByVal ValueType As Long) As String
    Dim TempValue As String                             ' 注册表关键字的临时值
    Dim Value As String                                 ' 注册表关键字的值
    Dim ValueSize As Long                               ' 注册表关键字的值的实际长度
    TempValue = Space(1024)                             ' 存储注册表关键字的临时值的缓冲区
    ValueSize = 1024                                    ' 设置注册表关键字的值的默认长度
    ' 打开一个已存在的注册表关键字...
    RegOpenKeyEx keyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey
    If hKey = 0 Then
        GetKeyValue = "^_*_*_^"
        Exit Function
    End If
    Dim x As Integer
    x = RegQueryValueEx(hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize)
    ' 获得已打开的注册表关键字的值...
    If x <> 0 Then
        If x = 2 And ValueSize = 1024 Then
            GetKeyValue = "^_*_*_^"
            Exit Function
        End If
    End If
    ' 返回注册表关键字的的值...
    Select Case ValueType                                                        ' 通过判断关键字的类型, 进行处理
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            If ValueSize > 0 Then TempValue = Left$(TempValue, ValueSize - 1)                       ' 去掉TempValue尾部空格
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(3) As Byte
            RegQueryValueEx hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' 生成长度为8的十六进制字符串
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' 将十六进制的 Value 转换为十进制
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1) As Byte                                     ' 存储 REG_BINARY 值的临时数组
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' 将数组转换成字符串
                Next i
            End If
    End Select
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
    Value = Trim(Value)
    If InStr(Value, Chr(0)) Then
        GetKeyValue = Left(Value, InStr(Value, Chr(0)) - 1)                                       ' 返回函数值
    Else
        GetKeyValue = Value
    End If
End Function
Public Function RegDeleteKeyName(mhKey As keyRoot, SubKey As String, hKeyName As String) As Boolean
    '删除子键数据
    'mhKey是指主键的名称，SubKey是指路径，hKeyName是指键名
    Dim hKey As Long, ret As Long
    ret = RegOpenKey(mhKey, SubKey, hKey)
    RegDeleteKeyName = False
    If ret = 0 Then
        If RegDeleteValue(hKey, hKeyName) = 0 Then RegDeleteKeyName = True
    End If
    RegCloseKey hKey '删除打开的键值，释放内存
End Function
'-------------------------------------------------------------------------------------------------------------
'- 将 Double 型( 限制在 0--2^32-1 )的数字转换为十六进制并在前面补零
'- 参数说明: Number--要转换的 Double 型数字
'-------------------------------------------------------------------------------------------------------------
Private Function DoubleToHex(ByVal Number As Double) As String
    Dim strHex As String
    strHex = Space(8)
    For i = 1 To 8
        Select Case Number - Int(Number / 16) * 16
            Case 10
                Mid(strHex, 9 - i, 1) = "A"
            Case 11
                Mid(strHex, 9 - i, 1) = "B"
            Case 12
                Mid(strHex, 9 - i, 1) = "C"
            Case 13
                Mid(strHex, 9 - i, 1) = "D"
            Case 14
                Mid(strHex, 9 - i, 1) = "E"
            Case 15
                Mid(strHex, 9 - i, 1) = "F"
            Case Else
                Mid(strHex, 9 - i, 1) = CStr(Number - Int(Number / 16) * 16)
        End Select
        Number = Int(Number / 16)
    Next i
    DoubleToHex = strHex
End Function

Public Function GetKeyValueType(ByVal keyRoot As keyRoot, ByVal KeyName As String, ByVal checkValueName As String) As ValueType
    Dim f As FILETIME, CountKey As Long, CountValue As Long, MaxLenKey As Long, MaxLenValue As Long
    Dim l As Long, s As String, strTmp As String, intTmp As Long, ValueName() As String, ValueType() As ValueType
    
    ' 打开一个已存在的注册表关键字...
    Success = RegOpenKeyEx(keyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If Success <> ERROR_SUCCESS Then GetKeyValueType = 0: RegCloseKey hKey: Exit Function
    
    ' 获得一个已打开的注册表关键字的信息...
    Success = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, ByVal 0&, CountKey, MaxLenKey, ByVal 0&, CountValue, MaxLenValue, ByVal 0&, ByVal 0&, f)
    
    If Success <> ERROR_SUCCESS Then GetKeyValueType = 0: RegCloseKey hKey: Exit Function
    If CountValue <> 0 Then
        ReDim ValueName(CountValue - 1) As String           ' 重新定义数组, 使用数组大小与注册表关键字的子键数量匹配
        ReDim ValueType(CountValue - 1) 'As Long             ' 重新定义数组, 使用数组大小与注册表关键字的子键数量匹配
        For i = 0 To CountValue - 1
            strTmp = String(255, vbNullChar) 'Space(255)
            l = 255
            RegEnumValue hKey, i, ByVal strTmp, l, 0, intTmp, ByVal 0&, ByVal 0&
            ValueType(i) = intTmp
            ValueName(i) = Left(strTmp, l)
            If InStr(ValueName(i), vbNullChar) - 1 <> -1 Then
                ValueName(i) = Left$(ValueName(i), InStr(ValueName(i), vbNullChar) - 1)
            End If
            If ValueName(i) = checkValueName Then
                GetKeyValueType = ValueType(i)
                Exit Function
            End If
        Next i
    End If
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
End Function
Public Function GetKeyInfo(keyRoot As keyRoot, KeyName As String, SubKeyName() As String, ValueName() As String, ValueType() As ValueType, Optional CountKey As Long, Optional CountValue As Long, Optional MaxLenKey As Long, Optional MaxLenValue As Long) As Boolean
    Dim f As FILETIME
    Dim l As Long, s As String, strTmp As String, intTmp As Long
    
    ' 打开一个已存在的注册表关键字...
    Success = RegOpenKeyEx(keyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If Success <> ERROR_SUCCESS Then GetKeyInfo = False: RegCloseKey hKey: Exit Function
    
    ' 获得一个已打开的注册表关键字的信息...
    Success = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, ByVal 0&, CountKey, MaxLenKey, ByVal 0&, CountValue, MaxLenValue, ByVal 0&, ByVal 0&, f)
    
    If Success <> ERROR_SUCCESS Then GetKeyInfo = False: RegCloseKey hKey: Exit Function
    
    If CountKey <> 0 Then
        ReDim SubKeyName(CountKey - 1) As String            ' 重新定义数组, 使用数组大小与注册表关键字的子项数量匹配
        For i = 0 To CountKey - 1
            strTmp = String(255, vbNullChar) 'Space(255)
            l = 255
            RegEnumKeyEx hKey, i, ByVal strTmp, l, 0, vbNullString, ByVal 0&, f
            SubKeyName(i) = Left(strTmp, l)
            If InStr(SubKeyName(i), vbNullChar) - 1 <> -1 Then
                SubKeyName(i) = Left$(SubKeyName(i), InStr(SubKeyName(i), vbNullChar) - 1)
            End If
        Next i
        
        ' 下面的二重循环对字符串数组进行冒泡排序
        For i = 0 To UBound(SubKeyName)
            For j = i + 1 To UBound(SubKeyName)
                If SubKeyName(i) > SubKeyName(j) Then
                    s = SubKeyName(i)
                    SubKeyName(i) = SubKeyName(j)
                    SubKeyName(j) = s
                End If
            Next j
        Next i
    End If
    If CountValue <> 0 Then
        ReDim ValueName(CountValue - 1) As String           ' 重新定义数组, 使用数组大小与注册表关键字的子键数量匹配
        ReDim ValueType(CountValue - 1) 'As Long             ' 重新定义数组, 使用数组大小与注册表关键字的子键数量匹配
        For i = 0 To CountValue - 1
            strTmp = String(255, vbNullChar) 'Space(255)
            
            l = 255
            RegEnumValue hKey, i, ByVal strTmp, l, 0, intTmp, ByVal 0&, ByVal 0&
            ValueType(i) = intTmp
            ValueName(i) = Left(strTmp, l)
            If InStr(ValueName(i), vbNullChar) - 1 <> -1 Then
                ValueName(i) = Left$(ValueName(i), InStr(ValueName(i), vbNullChar) - 1)
            End If
        Next i
        
        ' 下面的二重循环对字符串数组进行冒泡排序
        For i = 0 To UBound(ValueName)
            For j = i + 1 To UBound(ValueName)
                If ValueName(i) > ValueName(j) Then
                    s = ValueName(i)
                    ValueName(i) = ValueName(j)
                    ValueName(j) = s
                End If
            Next j
        Next i
    End If
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
    GetKeyInfo = True                                   ' 返回函数值
End Function
Public Function RegDeleteSubkey(hKey As keyRoot, SubKey As String) As Boolean
    '删除目录
    'mhKey是指主键的名称，SubKey是指路径
    Dim ret As Long, Index As Long, hName As String
    Dim hSubkey As Long
    ret = RegOpenKey(hKey, SubKey, hSubkey)
    If ret <> 0 Then
        RegDeleteSubkey = False
        Exit Function
    End If
    ret = RegDeleteKey(hSubkey, "")
    If ret <> 0 Then '如果删除失败则认为是NT则用递归方法删除目录
        hName = String(256, Chr(0))
        While RegEnumKey(hSubkey, 0, hName, Len(hName)) = 0 And _
              RegDeleteSubkey(hSubkey, hName)
        Wend
        ret = RegDeleteKey(hSubkey, "")
    End If
    RegDeleteSubkey = (ret = 0)
    RegCloseKey hSubkey '删除打开的键值，释放内存
End Function
Public Sub GetRegRootPath(ByVal RegPath As String, regRoot As keyRoot)
    If InStr(UCase(RegPath), "HKEY_CLASSES_ROOT") > 0 Then
        regRoot = HKEY_CLASSES_ROOT
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_CONFIG") > 0 Then
        regRoot = HKEY_CURRENT_CONFIG
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_USER") > 0 Then
        regRoot = HKEY_CURRENT_USER
    ElseIf InStr(UCase(RegPath), "HKEY_DYN_DATA") > 0 Then
        regRoot = HKEY_DYN_DATA
    ElseIf InStr(UCase(RegPath), "HKEY_LOCAL_MACHINE") > 0 Then
        regRoot = HKEY_LOCAL_MACHINE
    ElseIf InStr(UCase(RegPath), "HKEY_PERFORMANCE_DATA") > 0 Then
        regRoot = HKEY_PERFORMANCE_DATA
    Else
        regRoot = HKEY_USERS
    End If
End Sub
Public Function GetRegSubPath(ByVal RegPath As String) As String
    If InStr(UCase(RegPath), "HKEY_CLASSES_ROOT") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_CLASSES_ROOT") + 2, Len(RegPath) - Len("HKEY_CLASSES_ROOT") + 1)
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_CONFIG") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_CURRENT_CONFIG") + 2, Len(RegPath) - Len("HKEY_CURRENT_CONFIG") + 1)
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_USER") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_CURRENT_USER") + 2, Len(RegPath) - Len("HKEY_CURRENT_USER") + 1)
    ElseIf InStr(UCase(RegPath), "HKEY_DYN_DATA") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_DYN_DATA") + 2, Len(RegPath) - Len("HKEY_DYN_DATA") + 1)
    ElseIf InStr(UCase(RegPath), "HKEY_LOCAL_MACHINE") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_LOCAL_MACHINE") + 2, Len(RegPath) - Len("HKEY_LOCAL_MACHINE") + 1)
    ElseIf InStr(UCase(RegPath), "HKEY_PERFORMANCE_DATA") > 0 Then
        GetRegSubPath = Mid(RegPath, Len("HKEY_PERFORMANCE_DATA") + 2, Len(RegPath) - Len("HKEY_PERFORMANCE_DATA") + 1)
    Else
        GetRegSubPath = Mid(RegPath, Len("HKEY_USERS") + 2, Len(RegPath) - Len("HKEY_USERS") + 1)
    End If
End Function
'Public Sub GetRegType(ByVal RegType As String, valueTypes As ValueType)
'    Select Case RegType
'        Case "1"
'            valueTypes = REG_SZ
'        Case "2"
'            valueTypes = REG_EXPAND_SZ
'        Case "3"
'            valueTypes = REG_BINARY
'        Case "4"
'            valueTypes = REG_DWORD
'        Case "7"
'            valueTypes = REG_MULTI_SZ
'        Case Else
'            valueTypes = REG_SZ
'    End Select
'End Sub
Public Function RegRootPathIsTrue(ByVal RegPath As String) As Boolean
    If InStr(UCase(RegPath), "HKEY_CLASSES_ROOT") > 0 Then
        RegRootPathIsTrue = True
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_CONFIG") > 0 Then
        RegRootPathIsTrue = True
    ElseIf InStr(UCase(RegPath), "HKEY_CURRENT_USER") > 0 Then
        RegRootPathIsTrue = True
    ElseIf InStr(UCase(RegPath), "HKEY_DYN_DATA") > 0 Then
        RegRootPathIsTrue = True
    ElseIf InStr(UCase(RegPath), "HKEY_LOCAL_MACHINE") > 0 Then
        RegRootPathIsTrue = True
    ElseIf InStr(UCase(RegPath), "HKEY_PERFORMANCE_DATA") > 0 Then
        RegRootPathIsTrue = True
    Else
        RegRootPathIsTrue = False
    End If
End Function
Public Function GetRegRoot(regRoot As keyRoot) As String
    Select Case regRoot
        Case &H80000000
            GetRegRoot = "HKEY_CLASSES_ROOT"
        Case &H80000001
            GetRegRoot = "HKEY_CURRENT_USER"
        Case &H80000002
            GetRegRoot = "HKEY_LOCAL_MACHINE"
        Case &H80000003
            GetRegRoot = "HKEY_USERS"
        Case &H80000004
            GetRegRoot = "HKEY_PERFORMANCE_DATA"
        Case &H80000005
            GetRegRoot = "HKEY_CURRENT_CONFIG"
        Case &H80000006
            GetRegRoot = "HKEY_DYN_DATA"
    End Select
        
End Function


