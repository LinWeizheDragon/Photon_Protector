Attribute VB_Name = "modControls"


Option Explicit
'获取注册表子路径
Public Function GetRegistrySubPath(ByVal strRegPath As String) As String
    Dim strTmp As String, blnIsMachine As Boolean, intStart As Integer
    If InStr(strRegPath, "\REGISTRY\MACHINE") > 0 Then blnIsMachine = True
    intStart = InStr(strRegPath, "*value:")
    If intStart > 0 Then
        If blnIsMachine Then
            strTmp = Mid(strRegPath, Len("\REGISTRY\MACHINE") + 2, intStart - Len("\REGISTRY\MACHINE") - 1)
        Else
            strTmp = Mid(strRegPath, Len("\REGISTRY\USER") + 2, intStart - Len("\REGISTRY\USER") - 1)
        End If
        strTmp = GetPath(strTmp)
        GetRegistrySubPath = Left(strTmp, Len(strTmp) - 1)
        Exit Function
    Else
        intStart = InStr(strRegPath, "**")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = Mid(strRegPath, Len("\REGISTRY\MACHINE") + 2, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = Mid(strRegPath, Len("\REGISTRY\USER") + 2, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetPath(strTmp)
            GetRegistrySubPath = Left(strTmp, Len(strTmp) - 1)
            Exit Function
        End If
        intStart = InStr(strRegPath, "^^")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = Mid(strRegPath, Len("\REGISTRY\MACHINE") + 2, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = Mid(strRegPath, Len("\REGISTRY\USER") + 2, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetPath(strTmp)
            GetRegistrySubPath = Left(strTmp, Len(strTmp) - 1)
            Exit Function
        End If
    End If
    
End Function



'获取keyRoot对应的字符串
Public Function GetRootString(ByVal strRegPath As String) As String
    If InStr(UCase(strRegPath), "\REGISTRY\MACHINE") > 0 Then
        GetRootString = "HKEY_LOCAL_MACHINE"
    ElseIf InStr(UCase(strRegPath), "\REGISTRY\USER") > 0 Then
        GetRootString = "HKEY_USERS"
    End If
End Function

'获取注册表路径，因为从DLL传来的是以REGISTRY开始的
Public Function GetRegistryPath(ByVal strRegPath As String) As String
    Dim strTmp As String, blnIsMachine As Boolean, intStart As Integer
    strTmp = GetRootString(strRegPath)
    If InStr(strRegPath, "\REGISTRY\MACHINE") > 0 Then blnIsMachine = True
    intStart = InStr(strRegPath, "*value:")
    If intStart > 0 Then
        If blnIsMachine Then
            strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
        Else
            strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
        End If
        strTmp = GetPath(strTmp)
        GetRegistryPath = Left(strTmp, Len(strTmp) - 1)
        Exit Function
    Else
        intStart = InStr(strRegPath, "**")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetPath(strTmp)
            GetRegistryPath = Left(strTmp, Len(strTmp) - 1)
            Exit Function
        End If
        intStart = InStr(strRegPath, "^^")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetPath(strTmp)
            GetRegistryPath = Left(strTmp, Len(strTmp) - 1)
            Exit Function
        End If
    End If
End Function

'获取DLL传来的完整信息
Public Function GetFullPath(ByVal strPath As String)
    Dim strTmp As String, intStart As Integer
    intStart = InStr(strPath, ":")
    If intStart > 0 Then
        strTmp = Mid(strPath, intStart + 1, Len(strPath) - intStart)
    End If
    GetFullPath = strTmp
End Function

'获取注册表键名
Public Function GetRegValueName(ByVal strRegPath As String) As String
    Dim strTmp As String, blnIsMachine As Boolean, intStart As Integer
    strTmp = GetRootString(strRegPath)
    If InStr(strRegPath, "\REGISTRY\MACHINE") > 0 Then blnIsMachine = True
    intStart = InStr(strRegPath, "*value:")
    If intStart > 0 Then
        If blnIsMachine Then
            strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
        Else
            strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
        End If
        strTmp = GetFileName(strTmp)
        GetRegValueName = strTmp
        Exit Function
    Else
        intStart = InStr(strRegPath, "**")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetFileName(strTmp)
            GetRegValueName = strTmp
            Exit Function
        End If
        intStart = InStr(strRegPath, "^^")
        If intStart > 0 Then
            If blnIsMachine Then
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\MACHINE") + 1, intStart - Len("\REGISTRY\MACHINE") - 1)
            Else
                strTmp = strTmp & Mid(strRegPath, Len("\REGISTRY\USER") + 1, intStart - Len("\REGISTRY\USER") - 1)
            End If
            strTmp = GetFileName(strTmp)
            GetRegValueName = strTmp
            Exit Function
        End If
    End If
End Function

'获取注册表键值
Public Function GetRegValue(ByVal strRegPath As String) As String
    Dim strTmp As String, intStart As Integer, intStart1 As Integer
    intStart = InStr(strRegPath, "*value:")
    If intStart > 0 Then
        intStart1 = InStr(strRegPath, "**")
        If intStart1 > 0 Then
            strTmp = Mid(strRegPath, intStart + Len("*value:"), intStart1 - intStart - Len("*value:"))
            GetRegValue = strTmp
        Else
            intStart1 = InStr(strRegPath, "^^")
            If intStart1 > 0 Then
                strTmp = Mid(strRegPath, intStart + Len("*value:"), intStart1 - intStart - Len("*value:"))
                GetRegValue = strTmp
            Else
                GetRegValue = ""
            End If
        End If
    Else
        GetRegValue = ""
    End If
End Function

'获取操作类型
Public Function GetRegistryType(ByVal strRegPath As String) As String
    Dim strTmp As String, intStart As Integer, intStart1 As Integer
    intStart = InStr(strRegPath, "**")
    If intStart > 0 Then
        intStart1 = InStr(strRegPath, "^^")
        If intStart1 > 0 Then
            strTmp = Mid(strRegPath, intStart + Len("**"), intStart1 - intStart - Len("**"))
            GetRegistryType = strTmp
        Else
            GetRegistryType = ""
        End If
    Else
        GetRegistryType = ""
    End If
    
End Function



'

'获取指定注册表类型对应的类型
Public Function GetRegTypeString(ByVal strRegType As String) As String
    Select Case strRegType
        Case "1"
            GetRegTypeString = "REG_SZ"
        Case "2"
            GetRegTypeString = "REG_EXPAND_SZ"
        Case "3"
            GetRegTypeString = "REG_BINARY"
        Case "4"
            GetRegTypeString = "REG_DWORD"
        Case "7"
            GetRegTypeString = "REG_MULTI_SZ"
        Case Else
            GetRegTypeString = "REG_SZ"
    End Select
End Function

'获取进程路径信息包括没分离的PID信息
Public Function GetRegProcessPath(ByVal strRegPath As String) As String
    Dim strTmp As String, intStart As Integer
    intStart = InStr(strRegPath, "^^")
    If intStart > 0 Then
        strTmp = Mid(strRegPath, intStart + 2, Len(strRegPath) - intStart)
    End If
    GetRegProcessPath = strTmp
End Function

'获取进程路径信息
Public Function GetRegProcessPathEx(ByVal strRegPath As String) As String
    Dim strTmp As String, intStart As Integer
    intStart = InStr(strRegPath, "^^")
    If intStart > 0 Then
        strTmp = Mid(strRegPath, intStart + 2, InStr(strRegPath, "进程ID<") - 2 - intStart)
    End If
    GetRegProcessPathEx = strTmp
End Function

'此函数从字符串中分离出路径
Public Function GetPath(ByVal strPathIn As String) As String
    Dim i As Integer
    For i = Len(strPathIn) To 1 Step -1
        If InStr(":\", Mid$(strPathIn, i, 1)) Then Exit For
    Next
    GetPath = Left$(strPathIn, i)
End Function

'此函数从字符串中分离出文件名
Public Function GetFileName(ByVal strFileIn As String) As String
    Dim i As Integer
    For i = Len(strFileIn) To 1 Step -1
        If InStr("\", Mid$(strFileIn, i, 1)) Then Exit For
    Next
    GetFileName = Mid$(strFileIn, i + 1, Len(strFileIn) - i)
End Function
