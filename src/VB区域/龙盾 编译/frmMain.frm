VERSION 5.00
Begin VB.Form frmWatch 
   Caption         =   "龙盾-实时监控"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   1920
   ScaleWidth      =   4350
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "关闭"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "隐藏"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "状态"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "检测中......"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Working As Boolean
'引用Microsoft WMI Scripting V1.2 Library
Private objSWbemServices As SWbemServices
Private WithEvents CreateProcessEvent As SWbemSink
Attribute CreateProcessEvent.VB_VarHelpID = -1
Private WithEvents DeleteProcessEvent As SWbemSink
Attribute DeleteProcessEvent.VB_VarHelpID = -1
Private WithEvents ModificationProcessEvent As SWbemSink
Attribute ModificationProcessEvent.VB_VarHelpID = -1
Private Declare Function OpenProcess Lib "kernel32.dll " (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'打开进程API
Private Declare Function EnumProcessModules Lib "psapi.dll " (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'进程模块API
Private Declare Function GetModuleFileNameExA Lib "psapi.dll " (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
'获得进程EXE执行文件模块API
Private Declare Function CloseHandle Lib "kernel32.dll " (ByVal hObject As Long) As Long
'关闭句柄API
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private hProcess As Long

Function GetProcessPathByProcessID(PID As Long) As String
        On Error GoTo Z
        Dim cbNeeded     As Long
        Dim szBuf(1 To 250)         As Long
        Dim ret     As Long
        Dim szPathName     As String
        Dim nSize     As Long
        Dim hProcess     As Long
        hProcess = OpenProcess(&H400 Or &H10, 0, PID)
        If hProcess <> 0 Then
                ret = EnumProcessModules(hProcess, szBuf(1), 250, cbNeeded)
                If ret <> 0 Then
                        szPathName = Space(260)
                        nSize = 500
                        ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
                        GetProcessPathByProcessID = Left(szPathName, ret)
                End If
        End If
        ret = CloseHandle(hProcess)
        If GetProcessPathByProcessID = " " Then
              GetProcessPathByProcessID = "SYSTEM "
        End If
        Exit Function
Z:
End Function


Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
If Working = True Then
  CreateProcessEvent.Cancel
  DeleteProcessEvent.Cancel
  Label1.Caption = "已经停止运行......"
  'addinfo "实时监控被关闭！......"
  Command2.Caption = "开启"
  Working = False
Else
  StartMonitorCreateProcessEvent
  StartMonitorDeleteProcessEvent
  Label1.Caption = "正在运行，实时监控系统所有进程的创建......"
  'addinfo "实时监控开启......"
  Command2.Caption = "关闭"
  Working = True
End If
End Sub

Private Sub Command3_Click()

End Sub

'各种操作集合
'StartMonitorCreateProcessEvent
'StartMonitorDeleteProcessEvent
'StartMonitorModificationProcessEvent
'CreateProcessEvent.Cancel
'DeleteProcessEvent.Cancel
'ModificationProcessEvent.Cancel


Private Sub Form_Load()
StartMonitorCreateProcessEvent
StartMonitorDeleteProcessEvent

Label1.Caption = "正在运行，实时监控系统所有进程的创建......"
Working = True
'StartMonitorModificationProcessEvent
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> 1 Then
Cancel = True
Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CreateProcessEvent.Cancel
DeleteProcessEvent.Cancel

Working = False
'ModificationProcessEvent.Cancel
End Sub

'进程创建事件
Private Sub CreateProcessEvent_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
On Error Resume Next
         Dim Result As String
         Dim DesString As String
Dim ProcessName As String, ProcessID As Long, ProcessPath As String, CommandLine As String, CreationDate As String
ProcessName = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("Name").Value
ProcessID = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ProcessId").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("CommandLine").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("CreationDate").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ExecutablePath").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("Handle").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("CreationDate").Value
'Debug.Print objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ProcessId").Value
ProcessPath = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ExecutablePath").Value
CommandLine = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("CommandLine").Value
If ProcessPath = "" Then Exit Sub
'addinfo "实时防护：进程启动：" & ProcessName
''addinfo "进程ID：" & ProcessId
''addinfo "命令行：" & CommandLine
''addinfo "进程路径：" & ProcessPath
'SendText "process|" & ProcessName & "|" & ProcessPath, "process"
  Dim MyFSO As New FileSystemObject
 '  MsgBox ProcessPath & vbCrLf & ProcessName & vbCrLf & "已经从移动设备" & Split(ProcessPath, ":\")(0) & ":\" & "运行"
   If IsNumeric(ProcessID) Then
      hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CLng(ProcessID))
      If hProcess <> 0 Then
         NtSuspendProcess hProcess '挂起
         'Result = ProcessScan(ProcessPath)
         DesString = DesString & "进程名:" & ProcessName & "|" & "进程ID:" & ProcessID & "|" & "命令行:" & CommandLine & "|" & "进程路径:" & ProcessPath
         If CheckProcess(ProcessID, ProcessPath) = True Then
           NtResumeProcess hProcess '继续
         Else
           TerminateProcess hProcess, 0 '终止
         End If
      End If
   End If

End Sub
Private Function CheckProcess(ByVal ProcessID As String, ByVal ProcessPath As String) As Boolean
If ProcessID = "0" Then
Exit Function
End If
On Error GoTo ERR:

Dim YesNo As Boolean

Dim ID As String, Result As String
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
If Result = "" Then '没有记录，就新增一个默认的记录
  WriteString "Rules", """" & ProcessPath & """", "2", App.Path & "\Rules.ini"
End If
'重新读取
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
 If Result = "1" Then '信任的东西
 CheckProcess = True
 Exit Function
 ElseIf Result = "0" Then '不信任的东西
 CheckProcess = False
 Exit Function
 End If
 '默认2
Dim MyForm As New frmTip
MyForm.PicIcon = GetIconFromFile(ProcessPath, 0, True)
MyForm.Text1.Text = "进程：" & ProcessPath & vbCrLf & "进程ID：" & ProcessID & vbCrLf & "正在创建"
Dim MyFSO As New FileSystemObject
Dim StrDrv As String
StrDrv = Left(ProcessPath, 3)
If Right(StrDrv, 2) <> ":\" Then '如果不是标准路径名
  MyForm.Option2.Value = True
MyForm.Tip = "可疑进程正在创建中，请不要运行来历不明的文件！如果您是打开文件夹，则此为伪装的应用程序。"
GoTo Kip:
End If
If MyFSO.GetDrive(StrDrv).DriveType <> Fixed Then '如果是不是本地出现的东西
 MyForm.Option2.Value = True
MyForm.Tip = "非本地磁盘中正在运行可疑进程，请不要运行来历不明的文件！如果您是打开文件夹，则此为伪装的应用程序。"
Else
 MyForm.Option1.Value = True
MyForm.Tip = "本地磁盘中正在运行进程，请不要运行来历不明的文件！由于在本地磁盘中，可能是系统自动运行的程序，默认30秒放行。"
End If
Kip:
MyForm.Command1.Caption = "Ｘ"
MyForm.Show vbModal

'如果选择以后也这么处理
If MyForm.ChooseNum <> 1 And MyForm.ChooseNum <> 2 Then
Dim MyForm2 As New frmTip
MyForm2.PicIcon = GetIconFromFile(ProcessPath, 0, True)
   If MyForm.ChooseNum = 3 Then
   WriteString "Rules", """" & ProcessPath & """", "1", App.Path & "\Rules.ini"
   MyForm2.Text1 = "文件：" & ProcessPath & vbCrLf & "已经添加到龙盾的信任列表，不拦截，不扫描。"
   ElseIf MyForm.ChooseNum = 4 Then
   MyForm2.Text1 = "文件：" & ProcessPath & vbCrLf & "已经添加到龙盾的黑名单列表，禁止运行，禁止操作"
   WriteString "Rules", """" & ProcessPath & """", "0", App.Path & "\Rules.ini"
   End If
MyForm2.Option1.Visible = False
MyForm2.Option2.Visible = False
MyForm2.Check1.Visible = False
MyForm2.Command2.Caption = "我知道了"
MyForm2.Label2.Visible = False
MyForm2.Label3.Caption = "添加规则"
MyForm2.Show
End If

If MyForm.ChooseNum = 1 Then
CheckProcess = True
ElseIf MyForm.ChooseNum = 2 Then
CheckProcess = False
ElseIf MyForm.ChooseNum = 3 Then
CheckProcess = True
ElseIf MyForm.ChooseNum = 4 Then
CheckProcess = False
End If

ERR:
End Function
'进程退出事件
Private Sub DeleteProcessEvent_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
On Error Resume Next
Dim ProcessName As String, ProcessID As Long, ProcessPath As String, CommandLine As String, CreationDate As String
ProcessName = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("Name").Value
ProcessID = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ProcessId").Value
ProcessPath = objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("ExecutablePath").Value
If ProcessPath = "" Then Exit Sub
'addinfo "实时防护：进程退出：" & ProcessName
End Sub

'进程属性变更事件
Private Sub ModificationProcessEvent_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
'addinfo "属性更改：" & objWbemObject.Properties_.item("TargetInstance").Value.Properties_.item("Name").Value
End Sub

  
Private Sub StartMonitorCreateProcessEvent()
Set CreateProcessEvent = New SWbemSink
Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
objSWbemServices.ExecNotificationQueryAsync CreateProcessEvent, "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
End Sub

Private Sub StartMonitorDeleteProcessEvent()
Set DeleteProcessEvent = New SWbemSink
Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
objSWbemServices.ExecNotificationQueryAsync DeleteProcessEvent, "SELECT * FROM __InstanceDeletionEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
End Sub

Private Sub StartMonitorModificationProcessEvent()
Set ModificationProcessEvent = New SWbemSink
Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")
objSWbemServices.ExecNotificationQueryAsync ModificationProcessEvent, "SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
End Sub

