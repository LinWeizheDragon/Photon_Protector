Attribute VB_Name = "mdHook"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Public c As Long
Option Explicit
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Function GetProcessNameByProcessId(ByVal Pid As Long) As String
     Dim szBuf(1 To 250) As Long
     Dim Ret As Long
     Dim szPathName As String
     Dim nSize As Long
     Dim hProcess As Long
     hProcess = OpenProcess(&H400 Or &H10, 0, Pid)
     If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, Pid)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
            GetProcessNameByProcessId = szPathName
        End If
     End If
     Ret = CloseHandle(hProcess)
End Function


Public Function Wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim s As String
MsgBox ".."
Dim cds As COPYDATASTRUCT
    If Msg = &H4A Then
    CopyMemory cds, ByVal lParam, Len(cds)
    s = Space(cds.cbData)
    CopyMemory ByVal s, ByVal cds.lpData, cds.cbData
    s = StrConv(s, vbFromUnicode)
    s = Left(s, InStr(1, s, Chr(0)) - 1)
    Dim ProcessPath As String
    ProcessPath = s
    AddInfo ProcessPath
    AddInfo s
Dim MyFSO As New FileSystemObject
Dim Result As String
'If MyFSO.GetDrive(Split(ProcessPath, ":\")(0) & ":\").DriveType = Removable Then
'   MsgBox ProcessPath & vbCrLf & ProcessName & vbCrLf & "已经从移动设备" & Split(ProcessPath, ":\")(0) & ":\" & "运行"
'         Result = ProcessScan(ProcessPath)
'         AddInfo Result
'         If Result <> "SAFE" And Result <> "Error" Then '检测这是个病毒
'            DesString = DesString & "病毒名:" & Split(Result, "|")(0) & "|" & "描述:" & Split(Result, "|")(1) & "|"
'         End If
'         If Result = "SAFE" Then
'            DesString = DesString & "检测结果：安全" & "|"
'         ElseIf Result = "Error" Then
'            DesString = DesString & "扫描文件出错：未知" & "|"
'         End If
'
'         DesString = DesString & "创建者进程:" & GetProName(GetProcessNameByProcessId(wParam)) & "|" & "创建者路径" & GetProcessNameByProcessId(wParam) & "|" & "进程名:" & GetProName(ProcessPath) & "|" & "进程路径:" & ProcessPath
'         If ShowTextTip("检测到进程正在从移动设备创建", DesString, ProcessPath) = 1 Then
'           Wndproc = 1234
'         Else
'           Wndproc = 0
'         End If
'      End If
'   End If
'
'
'Else '没有在移动设备的
'
'         Result = ProcessScan(ProcessPath)
'         AddInfo Result
'         If Result <> "SAFE" And Result <> "Error" Then '检测这是个病毒
'            DesString = DesString & "病毒名:" & Split(Result, "|")(0) & "|" & "描述:" & Split(Result, "|")(1) & "|"
'         End If
'         If Result = "SAFE" Then
'            DesString = DesString & "检测结果：安全" & "|"
'            GoTo Out:
'         ElseIf Result = "Error" Then
'            DesString = DesString & "扫描文件出错：未知" & "|"
'            GoTo Out:
'         End If
'
'         DesString = DesString & "创建者进程:" & GetProName(GetProcessNameByProcessId(wParam)) & "|" & "创建者路径" & GetProcessNameByProcessId(wParam) & "|" & "进程名:" & GetProName(ProcessPath) & "|" & "进程路径:" & ProcessPath
'         If ShowTextTip("检测到木马进程正在创建", DesString, ProcessPath) = 1 Then
'           Wndproc = 1234
'         Else
'           Wndproc = 0
'         End If
'      End If
'   End If
'
'
'End If
        If MessageBoxA(0, s, "", 4) = 6 Then
            Wndproc = 1234
        Else
            Wndproc = 0
        End If
        Exit Function

End If
Wndproc = CallWindowProc(c, hwnd, Msg, wParam, lParam)
End Function

Private Function GetProName(ByVal Path As String) As String
Dim i As Integer
i = InStrRev("c:\windows\text.txt", "\")
GetProName = Right("c:\windows\text.txt", Len(("c:\windows\text.txt") - i))
End Function
