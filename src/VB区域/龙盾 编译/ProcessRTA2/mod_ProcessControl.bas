Attribute VB_Name = "mod_ProcessControl"
Option Explicit
'*************************************
'* 保存为标准模块.bas
'* 作者:jpkb@qq.com
'* 开源,转载请保留作者信息。
'* 2008.11.7
'*************************************
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function NtQueryInformationProcess Lib "ntdll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByRef ProcessInformation As Any, ByVal lProcessInformationLength As Long, ByRef lReturnLength As Long) As Long
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 500
End Type
Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type
Public Const PROCESS_TERMINATE = &H1
Public Const TH32CS_SNAPPROCESS = 2
Public Const PROCESS_VM_READ = 16
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SYNCHRONIZE As Long = &H100000
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const TH32CS_SNAPHEAPLIST = 1
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000
Public Const MAX_PATH As Integer = 260
Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_SHOW = 5

Public Function GetProcessCmdLine(ByVal PID As Long) As String
    '返回程序命令行
    Dim strBuffer As String
    Dim hProcess As Long
    Dim offset1 As Long
    Dim offset2 As Long
    Dim Dummy As Long
    Dim info As PROCESS_BASIC_INFORMATION
    Const STATUS_SUCCESS As Long = 0
    offset1 = 1
    offset2 = 0
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    If (hProcess = 0) Then
        Exit Function
    End If
    If (NtQueryInformationProcess(hProcess, 0, info, Len(info), ByVal 0&) <> STATUS_SUCCESS) Then
        CloseHandle hProcess
        Exit Function
    End If
    If (ReadProcessMemory(hProcess, (info.PebBaseAddress + &H10), offset1, 4, Dummy) = STATUS_SUCCESS) Then
        CloseHandle hProcess
        Exit Function
    End If
    If (ReadProcessMemory(hProcess, (offset1 + &H44), offset2, 4, Dummy) = STATUS_SUCCESS) Then
        CloseHandle hProcess
        Exit Function
    End If
    strBuffer = String(512, " ")
    If (ReadProcessMemory(hProcess, offset2, ByVal strBuffer, 512, Dummy) = STATUS_SUCCESS) Then
        CloseHandle hProcess
        Exit Function
    End If
    CloseHandle hProcess
    strBuffer = Left$(strBuffer, InStr(strBuffer, Chr(0) & Chr(0)))
    GetProcessCmdLine = StrConv(strBuffer, vbFromUnicode)
End Function

Public Function GetProcessPath(ByVal PID As Long) As String
    '返回程路径。
    On Error GoTo Z
    Dim cbNeeded As Long
    Dim szBuf(1 To 250) As Long
    Dim Ret As Long
    Dim szPathName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameEx(hProcess, szBuf(1), szPathName, nSize)
            GetProcessPath = Left(szPathName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    If GetProcessPath = "" Then
        GetProcessPath = "SYSTEM"
    End If
    Exit Function
Z:
    GetProcessPath = "ERROR"
End Function

Public Function GetProcessPID(sEXEName As String, Optional ByVal ID As Long = 1) As Long
    'ID为进程列表中第ID个sEXEName。返回PID值
    Dim hPS As Long, xx As Long
    Dim pe32 As PROCESSENTRY32
    Dim buffer As String * 255, loaded As Boolean
    Dim hand As Long, EXEName As String
    Dim theloop As Long, myID As Long
    hPS = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hPS = -1 Then Exit Function
    pe32.dwSize = Len(pe32)
    theloop = ProcessFirst(hPS, pe32)
    While theloop <> 0 '
        hand = OpenProcess(PROCESS_TERMINATE, True, CLng(pe32.th32ProcessID))
        EXEName = pe32.szExeFile
        myID = pe32.th32ProcessID
        If UCase(Left(EXEName, Len(sEXEName))) = UCase(sEXEName) Then
            xx = xx + 1
            If ID = xx Then
                GetProcessPID = myID
                CloseHandle hPS
                Exit Function
            End If
        End If
        theloop = ProcessNext(hPS, pe32)
    Wend
    CloseHandle hPS
End Function


