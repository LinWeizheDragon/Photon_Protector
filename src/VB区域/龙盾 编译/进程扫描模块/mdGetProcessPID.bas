Attribute VB_Name = "mdGetProcessPID"
Option Explicit
'=================================================
'通过PID获得进程路径所用
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'==========================================================
'通过进程名获得PID所用 上面声明过的这里暂且注释掉
'私有的CreateToolhelp32Snapshot
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
'私有的TerminateProcess
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS = &H2&

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Const PROCESS_TERMINATE = 1

'==============================================================
'通过进程PID获取应用程序的完整路径
Public Function GetProcessPathByProcessID(pid As Long) As String
    On Error GoTo ErrLine
    Dim cbNeeded As Long
    Dim szBuf(1 To 250) As Long
    Dim Ret As Long
    Dim szPathName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, pid)
    If hProcess <> 0 Then
       Ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
       If Ret <> 0 Then
         szPathName = Space(260)
         nSize = 500
         Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
         GetProcessPathByProcessID = Left(szPathName, Ret)
       End If
    End If
    Ret = CloseHandle(hProcess)
    If GetProcessPathByProcessID = "" Then
       GetProcessPathByProcessID = "SYSTEM"
    End If
ErrLine:
End Function


'通过进程名获得进程PID
Public Function GetProcessPID(sProcess As String) As Long
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    Dim lProcess As Long
    Dim lExitCode As Long
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                'Dim lProcess As Long
                'Dim lExitCode As Long
                GetProcessPID = tPE.th32ProcessID
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Function

'通过进程名获得应用程序的路径(通过上面两个函数)
Public Function GetProcessPath(sProcess As String) As String
    Dim aa As Long
    aa = GetProcessPID(sProcess)
    GetProcessPath = GetProcessPathByProcessID(aa)
End Function
 


