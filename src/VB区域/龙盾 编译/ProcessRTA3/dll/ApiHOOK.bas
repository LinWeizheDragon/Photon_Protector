Attribute VB_Name = "ApiHOOK"
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'Download by http://www.codefans.net
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function CreateIExprSrvObj Lib "msvbvm60.dll" (ByVal p1_0 As Long, ByVal p2_4 As Long, ByVal p3_0 As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxW" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Public Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long



Public Const CREATE_SUSPENDED = &H4
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type


Public CreateProcessAddr As Long

Public CreateProcessOldcode1 As Long
Public CreateProcessOldcode2 As Long

Public Sub KillvbaSetSystemError()
        'msvbvm60.__vbaSetSystemError
        Dim code As Long
        Dim DllFunAddr As Long
        Dim hModule As Long, dwJmpAddr As Long
        hModule = LoadLibrary("msvbvm60")
        
        If hModule = 0 Then
                Exit Sub
        End If
        
        DllFunAddr = GetProcAddress(hModule, "__vbaSetSystemError")
        If DllFunAddr = 0 Then
                Exit Sub
        End If
        
        code = &HC3C3C3C3
        WriteProcessMemory -1, ByVal DllFunAddr, code, 4, ByVal 0
End Sub

Sub HookOff(ByVal dwFunaddr As Long)
        '8B FF 55 8B EC
        Dim code As Long
        code = &H8B55FF8B '高字
        WriteProcessMemory -1, ByVal dwFunaddr, CreateProcessOldcode1, 4, 0
        code = &HEC8B55FF
        WriteProcessMemory -1, ByVal dwFunaddr + 1, CreateProcessOldcode2, 4, 0
End Sub

Function HookOn(ByVal strDllName As String, ByVal strFunName As String, ByVal lngFunAddr As Long, ByVal JmpBackFunc As Long) As Long

        Dim DllFunAddr As Long
        Dim JmpBackOffset As Long
        Dim code As Long
        Dim tmp As Long

        Dim hModule As Long, dwJmpAddr As Long

        hModule = LoadLibrary(strDllName)

        If hModule = 0 Then
                Exit Function
        End If
        
        DllFunAddr = GetProcAddress(hModule, strFunName)
        If DllFunAddr = 0 Then
                Exit Function
        End If
        
        JmpBackOffset = DllFunAddr - JmpBackFunc - 5
        
        '先处理空壳函数
        ReadProcessMemory -1, ByVal DllFunAddr, CreateProcessOldcode1, 4, 0
        ReadProcessMemory -1, ByVal DllFunAddr + 1, CreateProcessOldcode2, 4, 0
        
        If (CreateProcessOldcode1 And &HE9) = &HE9 Then
            tmp = CreateProcessOldcode2 + DllFunAddr + 5
            tmp = tmp - JmpBackFunc - 5
            WriteProcessMemory -1, ByVal JmpBackFunc, CreateProcessOldcode1, 4, 0
            WriteProcessMemory -1, ByVal JmpBackFunc + 1, tmp, 4, 0
        Else
            HookOff JmpBackFunc
        End If
        
        
        code = &HE9
        WriteProcessMemory -1, ByVal JmpBackFunc + 5, code, 4, 0
        code = JmpBackOffset
        WriteProcessMemory -1, ByVal JmpBackFunc + 6, code, 4, 0
        HookOn = DllFunAddr
        
        
        'HOOK API
        code = &HE9

        WriteProcessMemory -1, ByVal DllFunAddr, code, 4, 0
        code = lngFunAddr - DllFunAddr - 5
        WriteProcessMemory -1, ByVal DllFunAddr + 1, code, 4, 0
        HookOn = DllFunAddr
End Function

Public Function CreateProcessCallBack(ByVal a As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, ByVal lpStartupInfo As Long, ByVal lpProcessInformation As Long, ByVal b As Long) As Long

       '  MessageBox 0, lpApplicationName, lpCommandLine, 16
        Dim tmpp As Long
        Dim tmpProcessinf As PROCESS_INFORMATION
    
        GetData ShareMem

        If IsWindow(ShareMem.AntiHwnd) <> 0 Then
                tmpp = pShareMem + 12
                Call wcscpy(ByVal tmpp, ByVal lpApplicationName)
                tmpp = pShareMem + 12 + 1000
                Call wcscpy(ByVal tmpp, ByVal lpCommandLine)
                tmpp = SendMessage(ShareMem.AntiHwnd, &H400, 0, ByVal 0)

                If tmpp = 0 Then
                        CreateProcessCallBack = 0
                    
                Else
                        'dwCreationFlags = dwCreationFlags Or CREATE_SUSPENDED '打上挂起标记
                        
                        CreateProcessCallBack = CreateProcessJmpBack(a, lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDriectory, lpStartupInfo, lpProcessInformation, b)
                        
                        CopyMemory tmpProcessinf, ByVal lpProcessInformation, LenB(tmpProcessinf)
                        InjectMyself tmpProcessinf.hProcess
                        'ResumeThread tmpProcessinf.hThread
                End If

        Else
                'MessageBoxA 0, "no", "lpCommandLine", 16
                CreateProcessCallBack = CreateProcessJmpBack(a, lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDriectory, lpStartupInfo, lpProcessInformation, b)
        End If

        'MessageBox 0, pShareMem + 8, pShareMem + 8 + 1000, 16

End Function
 
 Public Function CreateProcessJmpBack(ByVal a As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, ByVal lpStartupInfo As Long, ByVal lpProcessInformation As Long, ByVal b As Long) As Long
        Dim tmp As Long
        tmp = GetLastError()
        tmp = GetLastError()
        tmp = GetLastError()
        tmp = GetLastError()
        tmp = GetLastError()
 End Function
Public Function GetFunAddr(lngFunAddr As Long) As Long
    GetFunAddr = lngFunAddr
End Function

