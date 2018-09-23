VERSION 5.00
Begin VB.Form frmHookCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   2175
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "测试"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Timer"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Text            =   "HookNtCreateProcessEx"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unload"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmHookCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Load_Drv As New cls_Driver

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal Dst As Long, ByVal Src As Long, ByVal uLen As Long)
Private Declare Sub GetMem4 Lib "msvbvm60.dll" (ByVal Address As Long, ByVal Dst As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function ZwOpenProcess _
               Lib "ntdll.dll" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientId As CLIENT_ID) As Long


Const FILE_DEVICE_ROOTKIT As Long = &H2A7B
Const METHOD_BUFFERED     As Long = 0
Const METHOD_IN_DIRECT    As Long = 1
Const METHOD_OUT_DIRECT   As Long = 2
Const METHOD_NEITHER      As Long = 3
Const FILE_ANY_ACCESS     As Long = 0
Const FILE_READ_ACCESS    As Long = &H1     '// file & pipe
Const FILE_WRITE_ACCESS   As Long = &H2     '// file & pipe
Const FILE_READ_DATA      As Long = &H1
Const FILE_WRITE_DATA     As Long = &H2

Const TA_ALLOWCREATE      As Long = &H1
Const TA_UNALLOWCREATE    As Long = &H2
Const TA_LOOPING          As Long = &H1
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long '打开进程
Private Declare Function EnumProcessModules Lib "psapi.dll " (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, cbNeeded As Long) As Long '枚举进程模块
Private Declare Function GetModuleFileNameExA Lib "psapi.dll " (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long '获取模块文件名
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long '关闭句柄

Public Function GetProcPath(pid As Long) As String
'根据PID获取进程路径
    On Error GoTo over
    Dim tmp As Long, process As Long, Modules(255) As Long, Path As String * 512
    process = OpenProcess(&H400 Or &H10, 0, pid)                                '打开进程
    If process = 0 Then GoTo over                                               '判断进程打开是否成功
    If EnumProcessModules(process, Modules(0), 256, tmp) <> 0 Then GetProcPath = Replace(Left$(Path, GetModuleFileNameExA(process, Modules(0), Path, 256)), "\??\", "") '枚举模块成功 '取得路径'处理字串并返回
over:
    Call CloseHandle(process)                                                   '关闭句柄
End Function



Public Sub StartHookFunction()
DoEvents
    If EnablePrivilege(SE_DEBUG) = False Then
       If Not EnablePrivilege1(SE_DEBUG_PRIVILEGE, True) Then
          If MsgBox("程序初始化失败。是否退出？不退出可能造成严重后果。", vbYesNo, "错误") = vbYes Then
           Unload frmMain
           Unload Me
          End If
          
       End If
    End If
    '初始化驱动
    With Load_Drv
        .szDrvFilePath = App.Path & "\HookNtCreateProcessEx.sys"
        .szDrvLinkName = "HookNtCreateProcessEx"
        .szDrvSvcName = "HookNtCreateProcessEx"
        .szDrvDisplayName = "HookNtCreateProcessEx"
        .InstDrv
        .StartDrv
        .OpenDrv
        
    End With
    '加载驱动
    Call Load_Drv.IoControl(Load_Drv.CTL_CODE_GEN(&H805), 0, 0, 0, 0)
    frmMain.Status.Caption = "Ring0驱动加载成功，程序运行中……"
    Timer1.Enabled = True
    Call ShowTip("龙盾-高级实时防护", "驱动加载成功，实时监控所有进程的创建……", 4)

End Sub

Public Sub StopHookFunction()
DoEvents
 '卸载驱动，关闭计时器
    With Load_Drv
        .DelDrv
    End With
    Timer1.Enabled = False
    frmMain.Status.Caption = "程序暂停运行……"
    Unload Me
End Sub

Private Sub Command4_Click()

    Call Load_Drv.IoControl(Load_Drv.CTL_CODE_GEN(&H806), 0, 0, 0, 0)
    Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
'MsgBox YesOrNo("I:\WPS.19.996.exe", 1000, "C:\Windows\System32\explorer.exe")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With Load_Drv '卸载驱动
        .DelDrv
    End With
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
    Dim i As Long
    With Load_Drv
    Call .IoControl(.CTL_CODE_GEN(&H801), 0, 0, 0, 0, i)
    
    'MsgBox i
    'Debug.Print i
    If i = TA_LOOPING Then
        
        Dim allow As Long
        Dim ProcessName As String, FProcessPath As String
        Dim ProcessID As Long
        
        ProcessName = String$(260, 0)
        
        Call .IoControl(.CTL_CODE_GEN(&H802), 0, 0, VarPtr(ProcessID), 4)
        Call .IoControl(.CTL_CODE_GEN(&H803), 0, 0, StrPtr(ProcessName), 260)
        
        ProcessName = StrConv(ProcessName, vbUnicode)
        ProcessName = Left(ProcessName, InStr(1, ProcessName, Chr(0)) - 1)
        FProcessPath = GetProcPath(ProcessID)
        Debug.Print ProcessName
        Debug.Print FProcessPath
        Debug.Print ProcessID
         Timer1.Enabled = False
         
        ' If MsgBox("创建进程是否允许", vbYesNo) = vbYes Then
        If CheckProcess(FProcessPath, ProcessName) = True Then
           allow = TA_ALLOWCREATE
          ' Call ShowTip("龙盾-高级实时防护", "进程：" & ProcessName & vbCrLf & "已放行……", 4)
         Else
           allow = TA_UNALLOWCREATE
         '  Call ShowTip("龙盾-高级实时防护", "进程：" & ProcessName & vbCrLf & "已拦截……", 4)
         End If

         Call .IoControl(.CTL_CODE_GEN(&H804), VarPtr(allow), 4, 0, 0)
         
         Dim hProcess As Long
         hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
         NtResumeProcess hProcess '继续
         ZwClose hProcess
         
         SuperSleep 1
         
         hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
         NtResumeProcess hProcess '继续
         ZwClose hProcess
         Timer1.Enabled = True
    End If
    End With
    

End Sub
Public Function CheckProcess(ByVal FromProcessPath As String, ByVal ProcessPath As String) As Boolean

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
MyForm.Text1.Text = "进程：" & FromProcessPath & vbCrLf & "正在创建进程：" & ProcessPath
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
