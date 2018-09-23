VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmMain 
   Caption         =   "进程管理器"
   ClientHeight    =   7080
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   2520
   End
   Begin VB.PictureBox PicMain 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6810
      Left            =   0
      ScaleHeight     =   454
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin VB.Frame 解锁_F 
         Caption         =   " 解 锁 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton Mmqx_C 
            Caption         =   "取 消"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   6
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   480
            Width           =   2445
         End
         Begin VB.CommandButton Mmqr_C 
            Caption         =   "确 定"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   4335
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "IM1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7091
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IM1 
      Left            =   3480
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   7680
      Top             =   1680
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "系统(&F)"
      Begin VB.Menu mnuRefurbish 
         Caption         =   "立即刷新(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu 修改密码_M 
         Caption         =   "修改密码"
      End
      Begin VB.Menu nu6 
         Caption         =   "-"
      End
      Begin VB.Menu 开机运行_M 
         Caption         =   "开机运行"
      End
      Begin VB.Menu 取消开机运行_M 
         Caption         =   "取消开机运行"
      End
      Begin VB.Menu nu7 
         Caption         =   "-"
      End
      Begin VB.Menu 注销_M 
         Caption         =   "注销"
      End
      Begin VB.Menu 关机_M 
         Caption         =   "关机"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "进程(&P)"
      Begin VB.Menu mnuEndPro 
         Caption         =   "结束进程(&E)"
      End
      Begin VB.Menu mnuDelPro 
         Caption         =   "删除进程文件(&D)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "文件属性(&R)..."
      End
      Begin VB.Menu mnuFolder 
         Caption         =   "所在目录(&F)..."
      End
      Begin VB.Menu mnuLine22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetProClass 
         Caption         =   "设置优先级(&P)"
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "实时(&R)"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "高(&H)"
            Index           =   1
         End
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "高于标准(&A)"
            Index           =   2
         End
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "标准(&N)"
            Index           =   3
         End
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "低于标准(&B)"
            Index           =   4
         End
         Begin VB.Menu mnuSetProClassSub 
            Caption         =   "低(&L)"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "窗口(&W)"
      Begin VB.Menu mnuOnTop 
         Caption         =   "总在最上(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu 隐藏_M 
         Caption         =   "隐藏"
      End
   End
   Begin VB.Menu XTSD_S 
      Caption         =   "系统锁定(&S)"
      Begin VB.Menu 全部锁定_M 
         Caption         =   "全部锁定"
      End
      Begin VB.Menu 全部解锁_M 
         Caption         =   "全部解锁"
      End
      Begin VB.Menu nu5 
         Caption         =   "-"
      End
      Begin VB.Menu XTSD_S_SD 
         Caption         =   "锁定运行程序"
         Shortcut        =   {F8}
      End
      Begin VB.Menu XTSD_S_JS 
         Caption         =   "解锁运行程序"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mu1 
         Caption         =   "-"
      End
      Begin VB.Menu 关闭任务栏_M 
         Caption         =   "关闭任务栏"
      End
      Begin VB.Menu 打开任务栏_M 
         Caption         =   "打开任务栏"
      End
      Begin VB.Menu nu2 
         Caption         =   "-"
      End
      Begin VB.Menu 关闭开始菜单_M 
         Caption         =   "关闭开始菜单"
      End
      Begin VB.Menu 打开开始菜单_M 
         Caption         =   "打开开始菜单"
      End
      Begin VB.Menu nu3 
         Caption         =   "-"
      End
      Begin VB.Menu 屏蔽任务管理器_M 
         Caption         =   "屏蔽任务管理器1"
      End
      Begin VB.Menu 屏蔽任务管理器_M2 
         Caption         =   "屏蔽任务管理器2"
      End
      Begin VB.Menu 屏蔽任务管理器_M3 
         Caption         =   "屏蔽任务管理器3"
      End
      Begin VB.Menu 取消屏蔽_M 
         Caption         =   "取消屏蔽"
      End
      Begin VB.Menu nu4 
         Caption         =   "-"
      End
      Begin VB.Menu 隐藏桌面图标_M 
         Caption         =   "隐藏桌面图标"
      End
      Begin VB.Menu 取消隐藏桌面图标_M 
         Caption         =   "取消隐藏桌面图标"
      End
      Begin VB.Menu nu11 
         Caption         =   "-"
      End
      Begin VB.Menu 关闭explorer_M 
         Caption         =   "关闭explorer1"
      End
      Begin VB.Menu 关闭explorer_M2 
         Caption         =   "关闭explorer2"
      End
      Begin VB.Menu 恢复explorer_M 
         Caption         =   "恢复explorer"
      End
      Begin VB.Menu 打开explorer_M 
         Caption         =   "打开explorer"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于本软件(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'可以使用FindWindow和ShowWindow来控制任务栏?下面给出一个例子程序?首先建立一个窗体和两个按钮?在窗体声明部分输入如下定义:
'{控制任务栏
Public DrvController As New clsProcess
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Any) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Const SW_HIDE = 0
    Const SW_SHOWNORMAL = 1
'控制任务栏}
'{控制开始菜单
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
  Private Const SW_RESTORE = 9
  Private Const SW_SHOW = 5
  Private Const SW_SHOWMAXIMIZED = 3
  'Private Const SW_HIDE = 0
  
  Private hhkLowLevelKybd As Long
''''''  制开始菜单状态
''''''  Dim hwnd     As Long
''''''  Dim ss     As Boolean
''''''  hwnd = FindWindow("Shell_TrayWnd", vbNullString)
''''''  hlong = FindWindowEx(hwnd, 0, "Button", vbNullString)
''''''  ss = IsWindowVisible(hlong)
''''''  If ss = True Then
''''''  MsgBox "visible"
''''''  Else
''''''  MsgBox "hide"
''''''  End If
  
'｝控制开始菜单
'{ 关机
Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_FORCE As Long = 4
Private Const EWX_POWEROFF As Long = 8

'The ExitWindowsEx function either logs off, shuts down, or shuts
'down and restarts the system.
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

'The GetLastError function returns the calling thread's last-error
'code value. The last-error code is maintained on a per-thread basis.
'Multiple threads do not overwrite each other's last-error code.
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
TheLuid As LUID
Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type

'The GetCurrentProcess function returns a pseudohandle for the
'current process.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'The OpenProcessToken function opens the access token associated with
'a process.
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'The LookupPrivilegeValue function retrieves the locally unique
'identifier (LUID) used on a specified system to locally represent
'the specified privilege name.
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

'The AdjustTokenPrivileges function enables or disables privileges
'in the specified access token. Enabling or disabling privileges
'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Private Const mlngWindows95 = 0
Private Const mlngWindowsNT = 1

Public glngWhichWindows32 As Long

'The GetVersion function returns the operating system in use.
Private Declare Function GetVersion Lib "kernel32" () As Long
'} 关机

''''''{热键
'''''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'''''Const WM_SETHOTKEY = &H32
'''''Const HOTKEYF_SHIFT = &H1
'''''Const HOTKEYF_CONTROL = &H2
'''''Const HOTKEYF_ALT = &H4
''''''}热键

'{禁用窗体右上角的关闭按钮
  Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
  Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'RemoveMenu GetSystemMenu(Me.hwnd, 0), &HF060, 0

'}禁用窗体右上角的关闭按钮



'{在VB程序中经常要使Ctrl-Alt-Delete和Ctrl-Esc 无效：

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'}在VB程序中经常要使Ctrl-Alt-Delete和Ctrl-Esc 无效：

Dim ProCount As Long    '当前进程数
Dim RamUse As Long  '当前已用内存
Dim IniFile1 As String  '系统进程解释文件
Dim SMJC_ss As Boolean '扫描进程
Dim MyNot As NOTIFYICONDATA '定义一个托盘结构
Dim Quit_C As Boolean  '退出状态


'{在VB程序中经常要使Ctrl-Alt-Delete和Ctrl-Esc 无效：

Sub DisableCtrlAltDelete(bDisabled As Boolean)

Dim x As Long

x = SystemParametersInfo(97, bDisabled, CStr(1), 0)

End Sub


Private Sub LoadDrv()
 With DrvController
        .szDrvFilePath = Replace(App.Path & "\TailList.sys", "\\", "\")
        .szDrvLinkName = "TailList"
        .szDrvDisplayName = "TailList"
        .szDrvSvcName = "TailList"
 End With
End Sub
Private Sub OpenDrv()
 With DrvController
        .InstDrv
        .StartDrv
        .OpenDrv
 End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IntR As Integer
    'IntR = MsgBox("确认要退出程序吗？", vbYesNo, "退出确认")
    'If IntR = vbNo Then
    '    If Quit_C = True Then Quit_C = False
    '    Cancel = -1
    '    Else
 ''       frmMain.Visible = False
    'End If
 ''    If Quit_C = False Then Cancel = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'{控制开始菜单
If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
'｝控制开始菜单

'{热键
Dim ret As Long
'取消Message的截取，使之送往原来的windows程序
ret = SetWindowLong(Me.hWnd, GWL_WNDPROC, preWinProc)
Call UnregisterHotKey(Me.hWnd, uVirtKey)
'}
 With DrvController
        .StopDrv
        .DelDrv
 End With

If trayflag = True Then

With MyNot

     .hIcon = frmMain.Icon '托盘图标指针

     .hWnd = frmMain.hWnd '窗体指针

     .szTip = "" '弹出提示字符串

     .uCallbackMessage = WM_USER + 100 '对应程序定义的消息

     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE '标志

     .uID = 1 '图标识别符

     .cbSize = Len(MyNot) '计算该结构所占字节数

     End With

    hh = Shell_NotifyIcon(NIM_DELETE, MyNot) '删除该图标

    trayflag = False '图标删除后trayflag为假


End If

     Unhook '退出消息循环


End Sub
'{ 关机
Private Sub AdjustToken()
'********************************************************************
'* This procedure sets the proper privileges to allow a log off or a
'* shut down to occur under Windows NT.
'********************************************************************

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2

Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long

'Set the error code of the last thread to zero using the
'SetLast Error function. Do this so that the GetLastError
'function does not return a value other than zero for no
'apparent reason.
SetLastError 0

'Use the GetCurrentProcess function to set the hdlProcessHandle
'variable.
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, _
(TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), _
hdlTokenHandle

'Get the LUID for shutdown privilege
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

tkp.PrivilegeCount = 1 ' One privilege to set
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED

'Enable the shutdown privilege in the access token of this process
AdjustTokenPrivileges hdlTokenHandle, _
False, _
tkp, _
Len(tkpNewButIgnored), _
tkpNewButIgnored, _
lBufferNeeded
End Sub
'} 关机
'列举进程
Public Sub ListProcess()
On Error Resume Next
    Dim I As Long, j As Long, n As Long
    Dim Jssl As Integer
    Dim proc As PROCESSENTRY32
    Dim snap As Long
    Dim exename As String
    Dim item As ListItem
    Dim lngHwndProcess As Long
    Dim lngModules(1 To 200) As Long
    Dim lngCBSize2 As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim WKSize As Long
    Dim strProcessName As String
    Dim strComment As String   '装载进程注释的字符串
    Dim ProClass As String
    Dim SMJC_S As Boolean  '扫描进程
    '开始进程循环
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    proc.dwSize = Len(proc)
    theloop = ProcessFirst(snap, proc)
    I = 0
    n = 0
    While theloop <> 0
        I = I + 1
        'exename = proc.szExeFile
        lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, proc.th32ProcessID)
        If lngHwndProcess <> 0 Then
            lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
            If lngReturn <> 0 Then
                strModuleName = Space(MAX_PATH)
                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, 500)
                strProcessName = Left(strModuleName, lngReturn)
                strProcessName = CheckPath(Trim$(strProcessName))
                If strProcessName <> "" Then
                    j = HaveItem(proc.th32ProcessID)
                    If j = 0 Then  '如果没有该进程
                        '获取短文件名
                        exename = Dir(strProcessName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
                        SMJC_S = True
                        
                        If SMJC_ss = True Then  '扫描进程 菜单的
                                For Jssl = 1 To GetINI("解锁文件", "数量", IniFile1)
                                    If exename = GetINI("解锁文件", Str(Jssl), IniFile1) Then SMJC_S = False
                                Next
                            Else
                            
                            SMJC_S = False
                        End If
                        
                        If SMJC_S = True Then
                                                        Dim hand As Long, id As Long
                                'If MsgBox("确定要结束进程 " & List1.SelectedItem.Text & " 吗？", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
                                id = CLng(proc.th32ProcessID)
                                If id <> 0 Then
                                    EndPro id
                                End If
                            Else
                            
                            
                            
                            
                            
                            exename = Dir(strProcessName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
                            If exename = "hh.exe" Then
                                'MsgBox SetProClass(proc.th32ProcessID, IDLE_PRIORITY_CLASS)
                            End If
                            
                            '添加进程item
                            Set item = List1.ListItems.Add(, "ID:" & CStr(proc.th32ProcessID), exename)
                            '进程ID
                            item.SubItems(1) = proc.th32ProcessID
                            '内存使用
                            pmc.cb = LenB(pmc)
                            lret = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                            n = n + pmc.WorkingSetSize
                            WKSize = pmc.WorkingSetSize / 1024
                            item.SubItems(2) = WKSize
                            '优先级
                            item.SubItems(5) = GetProClass(proc.th32ProcessID)
                            '进程路径
                            item.SubItems(6) = strProcessName
                            '进程图标
                            IM1.ListImages.Add , strProcessName, GetIcon(strProcessName)
                            item.SmallIcon = IM1.ListImages.item(strProcessName).Key
                            '这里判断是否为系统进程
                            strComment = ""
                            If UCase(Left$(strProcessName, Len(GetSysDir))) = UCase(GetSysDir) Then
                                strComment = GetINI("sysdir", Mid$(strProcessName, Len(GetSysDir) + 2), IniFile1)
                            ElseIf UCase(Left$(strProcessName, Len(GetWinDir))) = UCase(GetWinDir) Then
                                strComment = GetINI("windir", Mid$(strProcessName, Len(GetWinDir) + 2), IniFile1)
                            End If
                            If strComment <> "" Then
                                item.SubItems(3) = "系统"
                                item.SubItems(4) = Left$(strComment, 2)
                                item.SubItems(7) = Mid$(strComment, 4)
                            End If
                        End If
                    Else    '如果已经有该进程
                        pmc.cb = LenB(pmc)
                        lret = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                        n = n + pmc.WorkingSetSize
                        WKSize = pmc.WorkingSetSize / 1024
                        If CLng(List1.ListItems.item(j).SubItems(2)) <> WKSize Then List1.ListItems.item(j).SubItems(2) = WKSize
                        ProClass = GetProClass(proc.th32ProcessID)
                        If ProClass <> List1.ListItems.item(j).SubItems(5) Then List1.ListItems.item(j).SubItems(5) = ProClass
                    End If
                End If
            End If
        End If
        theloop = ProcessNext(snap, proc)
    Wend
    CloseHandle snap
    If I <> ProCount Then
        SB1.Panels.item(1) = "进程数：" & I
        ProCount = I
    End If
    If n <> RamUse Then
        SB1.Panels.item(2) = "已用内存：" & FormatLng(n)
        RamUse = n
    End If
End Sub
Public Sub ListProcess_s()
On Error Resume Next
    Dim I As Long, j As Long, n As Long
    Dim proc As PROCESSENTRY32
    Dim snap As Long
    Dim exename As String
    Dim item As ListItem
    Dim lngHwndProcess As Long
    Dim lngModules(1 To 200) As Long
    Dim lngCBSize2 As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim WKSize As Long
    Dim strProcessName As String
    Dim strComment As String   '装载进程注释的字符串
    Dim ProClass As String
    
    '开始进程循环
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    proc.dwSize = Len(proc)
    theloop = ProcessFirst(snap, proc)
    I = 0
    n = 0
    While theloop <> 0
        I = I + 1
        'exename = proc.szExeFile
        lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, proc.th32ProcessID)
        If lngHwndProcess <> 0 Then
            lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
            If lngReturn <> 0 Then
                strModuleName = Space(MAX_PATH)
                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, 500)
                strProcessName = Left(strModuleName, lngReturn)
                strProcessName = CheckPath(Trim$(strProcessName))
                If strProcessName <> "" Then
                    j = HaveItem(proc.th32ProcessID)
                    If j = 0 Then  '如果没有该进程
                        '获取短文件名
                            Dim hand As Long, id As Long
                                'If MsgBox("确定要结束进程 " & List1.SelectedItem.Text & " 吗？", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
                                id = CLng(proc.th32ProcessID)
                                If id <> 0 Then
                                    EndPro id
                                End If
                        
                    Else    '如果已经有该进程
                        pmc.cb = LenB(pmc)
                        lret = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                        n = n + pmc.WorkingSetSize
                        WKSize = pmc.WorkingSetSize / 1024
                        If CLng(List1.ListItems.item(j).SubItems(2)) <> WKSize Then List1.ListItems.item(j).SubItems(2) = WKSize
                        ProClass = GetProClass(proc.th32ProcessID)
                        If ProClass <> List1.ListItems.item(j).SubItems(5) Then List1.ListItems.item(j).SubItems(5) = ProClass
                    End If
                End If
            End If
        End If
        theloop = ProcessNext(snap, proc)
    Wend
    CloseHandle snap
    If I <> ProCount Then
        SB1.Panels.item(1) = "进程数：" & I
        ProCount = I
    End If
    If n <> RamUse Then
        SB1.Panels.item(2) = "已用内存：" & FormatLng(n)
        RamUse = n
    End If
End Sub
'设置进程优先级
Public Function SetProClass(ByVal PID As Long, ByVal ClassID As Long)
On Error Resume Next
    Dim hwd As Long
    hwd = OpenProcess(PROCESS_SET_INFORMATION, 0, PID)
    SetProClass = SetPriorityClass(hwd, ClassID)
End Function

'获取进程优先级
Public Function GetProClass(ByVal PID As Long) As String
On Error Resume Next
    Dim hwd As Long
    Dim Rtn As Long
    hwd = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
    Rtn = GetPriorityClass(hwd)
    Select Case Rtn
    Case IDLE_PRIORITY_CLASS
        GetProClass = "低"
    Case NORMAL_PRIORITY_CLASS
        GetProClass = "标准"
    Case HIGH_PRIORITY_CLASS
        GetProClass = "高"
    Case REALTIME_PRIORITY_CLASS
        GetProClass = "实时"
    Case 16384
        GetProClass = "较低"
    Case 32768
        GetProClass = "较高"
    End Select
End Function


'检查进程是否存在多余的已经结束的进程
Public Sub CheckProcess()
On Error Resume Next
    Dim lExit As Long
    Dim lngHwndProcess As Long
    Dim I As Long, j As Long
    If List1.ListItems.Count > 0 Then
        For I = List1.ListItems.Count To 1 Step -1
            j = CLng(List1.ListItems.item(I).SubItems(1))
            lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, j)
            If lngHwndProcess <> 0 Then
                GetExitCodeProcess lngHwndProcess, lExit
                If lExit <> STILL_ACTIVE Then List1.ListItems.Remove I
            Else
                List1.ListItems.Remove I
            End If
        Next
    End If
End Sub

'判断item是否存在
Public Function HaveItem(ByVal itemID As Long) As Long
On Error GoTo aaaa
    HaveItem = List1.ListItems("ID:" & CStr(itemID)).Index
Exit Function
aaaa:
    HaveItem = 0
End Function

Public Function CheckPath(ByVal PathStr As String) As String
On Error Resume Next
    PathStr = Replace(PathStr, "\??\", "")
    If UCase(Left$(PathStr, 12)) = "\SYSTEMROOT\" Then PathStr = GetWinDir & Mid$(PathStr, 12)
    CheckPath = PathStr
End Function



Private Sub Form_Load()
On Error Resume Next
App.Title = "进程管理器"
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hWnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


    If Quit_C = False Then
        If App.PrevInstance Then
            MsgBox "该程序已经运行！"
            End
        End If
    End If

        '隐藏应用
    App.TaskVisible = False
    'form1.showintaskbar=false
    '随着系统启动
        '******
                'If AddToStarup("abc", App.Path & "\进程管理器.exe") = True Then
                '      Debug.Print "OK"
                'Else
                '      Debug.Print "NO"
                'End If
        '******
        EnablePrivilege (SE_DEBUG)
 LoadDrv
  OpenDrv
 If DrvController.OpenDrv = False Then
 Call MsgBox("驱动加载失败！可能是您的操作系统版本不支持此驱动或被杀毒软件所拦截！", vbOKOnly, "驱动加载失败")
 SB1.Panels.item(3) = "驱动加载失败"
 Else
 SB1.Panels.item(3) = "病毒防御助手-进程管理器 驱动已加载"
 End If
 If WinVer >= 6 Then Label3.ForeColor = &HFF0000: Label3.Caption = "不支持": Command3.Enabled = False: MsgBox "您的系统版本不支持本驱动，十分抱歉！"
    Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    '初始化List1
    List1.ColumnHeaders.Add , , "进程名称", 120
    List1.ColumnHeaders.Add , , "PID", 45
    List1.ColumnHeaders.Add , , "内存(K)", 55
    List1.ColumnHeaders.Add , , "种类", 36
    List1.ColumnHeaders.Add , , "级别", 36
    List1.ColumnHeaders.Add , , "优先", 36
    List1.ColumnHeaders.Add , , "进程路径", 400
    List1.ColumnHeaders.Add , , "备注", 500
    List1.ColumnHeaders.item(3).Alignment = lvwColumnRight
    IniFile1 = GetApp & "sysset.ini"
'    '加载进程列表
    ListProcess
    
    
    SMJC_ss = False '扫描进程
    
   ' {热键
    Dim ret As Long

'记录原来的window程序地址
preWinProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
'用自定义程序代替原来的window程序
ret = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Wndproc)

idHotKey = 1
Modifiers = MOD_ALT + MOD_CONTROL 'Alt+Ctrl 键
uVirtKey = vbKeyJ  'J键
ret = RegisterHotKey(Me.hWnd, idHotKey, Modifiers, uVirtKey)
'}热键
''''''' {热键
''''''    Dim wHotkey As Long
''''''    '设置热键为Ctrl+Alt+A
''''''    wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * 256 + vbKeyJ
''''''    l = SendMessage(Me.hwnd, WM_SETHOTKEY, wHotkey, 0)
''''''
''''''
''''''' }热键


  '{ 托盘
  gHW = Me.hWnd '取得本窗体指针
     ''success = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)  '使窗口总显示在最前方
    '下一句调用钩子函数，将自制消息处理函数钩入Windows的消息循环
Digital1.Digital = Format(time(), "hh:mm:ss")

     hook
     
     Dim hh As Long

     With MyNot

     .hIcon = frmMain.Icon

     .hWnd = frmMain.hWnd

     .szTip = Str(Date) & Chr(&H0)

     .uCallbackMessage = WM_USER + 100

     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE

     .uID = 1

     .cbSize = Len(MyNot)

     End With

    hh = Shell_NotifyIcon(NIM_ADD, MyNot) '添加一个托盘图标

    trayflag = True '图标添加后trayflag为真
  '}
  
    '禁用窗体右上角的关闭按钮
RemoveMenu GetSystemMenu(Me.hWnd, 0), &HF060, 0
 Quit_C = False

    If Command = 1 Then
        全部锁定_M_Click
        frmMain.Visible = False
        'frmMain.WindowState = 1
    End If
    
    
    
    'frmMain.WindowState = 1
    Timer1.Interval = 200
    Timer2.Interval = 500
    
End Sub
Public Sub hook()

    '利用AddressOf取得消息处理函数WindowProc的指针，并将其传给SetWindowLong

    'lpPrevWndProc用来存储原窗口的指针

     lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)

    End Sub

    Public Sub Unhook()

    '本子程序用原窗口的指针替换WindowProc函数的指针，即关闭子类、退出消息循环

     Dim temp As Long

     temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)

    End Sub
Private Sub Form_Resize()
On Error Resume Next
    PicMain.Width = Me.Width / 15
    List1.Move 0, 0, PicMain.Width - 8, PicMain.Height
        
        If WindowState = 1 Then
            frmMain.Visible = False
            WindowState = 0
        End If
    
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    With List1
        If (ColumnHeader.Index - 1) = .SortKey Then
            .SortOrder = (.SortOrder + 1) Mod 2
            .Sorted = True
        Else
            .Sorted = False
            .SortOrder = 0
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End If
    End With
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Dim j As Long, I As Long
    If Button = 2 Then
        If List1.HitTest(x, y) Is Nothing Then Exit Sub
        j = List1.HitTest(x, y).Index
        List1.ListItems(j).Selected = True
        For I = 0 To 5
            mnuSetProClassSub(I).Checked = False
        Next
        Select Case List1.SelectedItem.SubItems(5)
        Case "实时": mnuSetProClassSub(0).Checked = True
        Case "高": mnuSetProClassSub(1).Checked = True
        Case "较高": mnuSetProClassSub(2).Checked = True
        Case "标准": mnuSetProClassSub(3).Checked = True
        Case "较低": mnuSetProClassSub(4).Checked = True
        Case "低": mnuSetProClassSub(5).Checked = True
        End Select
        If mnuProcess.Enabled = True Then PopupMenu mnuProcess
    End If
End Sub



Private Sub Mmqx_C_Click()
解锁_F.Visible = False
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next

    ShellAbout Me.hWnd, "进程管理器", "病毒防御助手", ByVal 0&
    'App.Title &
End Sub

Private Sub mnuDelPro_Click()
On Error Resume Next
    Dim I As Long, hand As Long, id As Long
    If MsgBox("确定要结束进程并删除以下文件吗？" & vbCrLf & vbCrLf & List1.SelectedItem.SubItems(6), vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
    id = CLng(List1.SelectedItem.SubItems(1))
    If id <> 0 Then
        EndPro id
        DoEvents
        DoEvents
        SuperSleep 1
        'FileDel CStr(List1.SelectedItem.SubItems(6))
        '换成kill的方法
        Call SetAttr(CStr(List1.SelectedItem.SubItems(6)), vbNormal)
        '这里必须先修改属性，不然会删不掉系统属性的文件
        Kill CStr(List1.SelectedItem.SubItems(6))
        SuperSleep 0.5
        If Not Dir(CStr(List1.SelectedItem.SubItems(6)), vbNormal Or vbSystem Or vbHidden Or vbReadOnly) = "" Then
          '删除失败
          MsgBox "删除文件" & CStr(List1.SelectedItem.SubItems(6)) & "失败！", vbOKOnly, "删除失败"
        End If
    End If
    ListProcess
End Sub

Private Sub mnuEndPro_Click()
On Error Resume Next
    Dim I As Long, hand As Long, id As Long
    If MsgBox("确定要结束进程 " & List1.SelectedItem.Text & " 吗？", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
    id = CLng(List1.SelectedItem.SubItems(1))
    If id <> 0 Then
        EndPro id
        With DrvController
        Call .IoControl(.CTL_CODE_GEN(&H801), VarPtr(id), 4, 0, 0)
  End With
    End If
    ListProcess
End Sub

'结束一个进程
Public Sub EndPro(ByVal PID As Long)
On Error Resume Next
    Dim lngHwndProcess As Long
    Dim hand As Long
    Dim exitCode As Long
    hand = OpenProcess(PROCESS_TERMINATE, True, PID)
    TerminateProcess hand, exitCode
    CloseHandle hand
End Sub

Private Sub mnuExit_Click()

    Quit_C = True
    Unload Me
    End
End Sub

Private Sub mnuFileProperties_Click()
    ShowProperties List1.SelectedItem.SubItems(6), Me.hWnd
End Sub

Private Sub mnuFolder_Click()
    ShellExecute hWnd, "open", GetAppF(List1.SelectedItem.SubItems(6)), "", "", 1
End Sub

Private Sub mnuHome_Click()
    ShellExecute hWnd, "open", "http://www.dvmsc.com", "", "", 1
End Sub

Private Sub mnuOnTop_Click()
    mnuOnTop.Checked = Not mnuOnTop.Checked
    SetTop Me, mnuOnTop.Checked
End Sub

Private Sub mnuRefurbish_Click()
    ListProcess
End Sub

Private Sub mnuSetProClassSub_Click(Index As Integer)
On Error Resume Next
    Dim PID As Long, Rtn As Long
    PID = CLng(List1.SelectedItem.SubItems(1))
    If mnuSetProClassSub(Index).Checked = True Then Exit Sub
    Select Case Index
    Case 0: Rtn = SetProClass(PID, REALTIME_PRIORITY_CLASS)
    Case 1: Rtn = SetProClass(PID, HIGH_PRIORITY_CLASS)
    Case 2: Rtn = SetProClass(PID, 32768)
    Case 3: Rtn = SetProClass(PID, NORMAL_PRIORITY_CLASS)
    Case 4: Rtn = SetProClass(PID, 16384)
    Case 5: Rtn = SetProClass(PID, IDLE_PRIORITY_CLASS)
    End Select
    If Rtn = 0 Then MsgBox "无法为进程 " & List1.SelectedItem.Text & " 设置优先级。", vbCritical
End Sub

Private Sub Timer1_Timer()
    ListProcess
End Sub

Private Sub Timer2_Timer()
    CheckProcess
End Sub



Private Sub XTSD_S_JS_Click()
'Timer3.Interval = 0
'Timer1.Interval = 200
SMJC_ss = False '扫描进程

End Sub

Private Sub XTSD_S_SD_Click()
'Timer1.Interval = 0
'Timer3.Interval = 200
SMJC_ss = True '扫描进程
End Sub

Private Sub 打开explorer_M_Click()

        Call Shell("explorer", 0)
        'Shell "Notepad", vbNormalFocus
        Shell "explorer", vbNormalFocus

        Shell "regedit", vbNormalFocus

End Sub

Private Sub 打开开始菜单_M_Click()
        Dim hWnd     As Long
        hWnd = FindWindow("Shell_TrayWnd", vbNullString)
        hlong = FindWindowEx(hWnd, 0, "Button", vbNullString)
       '"正常显示"
        ShowWindow hlong, SW_RESTORE
       '"最大化"
        'ShowWindow hlong, SW_SHOWMAXIMIZED
        '鍵
            UnhookWindowsHookEx hhkLowLevelKybd
            hhkLowLevelKybd = 0

End Sub

Private Sub 打开任务栏_M_Click()
     Dim hTaskBar As Long
     
     hTaskBar = FindWindow("Shell_TrayWnd", 0&)
     ShowWindow hTaskBar, SW_SHOWNORMAL
End Sub

Private Sub 关闭explorer_M_Click()

Dim I As Integer
    For I = 1 To 30
        Call Shell("tskill explorer", 0)
        Sleep 100
    Next
End Sub

Private Sub 关闭explorer_M2_Click()
Dim I As Integer
    For I = 1 To 30
        屏蔽任务管理器_M33_Click
      Sleep 100
    Next
End Sub

Private Sub 关闭开始菜单_M_Click()
        Dim hWnd     As Long
        hWnd = FindWindow("Shell_TrayWnd", vbNullString)
        hlong = FindWindowEx(hWnd, 0, "Button", vbNullString)
        ShowWindow hlong, SW_HIDE
        
        '鍵
        hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)

End Sub

Private Sub 关闭任务栏_M_Click()
     Dim hTaskBar As Long
     
     hTaskBar = FindWindow("Shell_TrayWnd", 0&)
     ShowWindow hTaskBar, SW_HIDE
End Sub

Private Sub 关机_M_Click()
Dim lngVersion As Long
    If MsgBox("确认要关机吗？", vbYesNo, "关机确认") = vbYes Then

        lngVersion = GetVersion()
        If ((lngVersion And &H80000000) = 0) Then
        glngWhichWindows32 = mlngWindowsNT
        Else
        glngWhichWindows32 = mlngWindows95
        End If

            If glngWhichWindows32 = mlngWindowsNT Then
                AdjustToken
            End If
            ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_POWEROFF), 0
    End If


End Sub

Private Sub 恢复explorer_M_Click()

    Call Savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe")

End Sub

Private Sub 开机运行_M_Click()
    If AddToStarup("abc", App.Path & "\进程管理器.exe 1") = True Then
            MsgBox "成功"
        Else
            MsgBox "失败"
    End If
End Sub

Private Sub 屏蔽任务管理器_M_Click()
    Open Environ("windir") & "\system32\taskmgr.exe" For Input Lock Read Write As #1
End Sub

Private Sub 屏蔽任务管理器_M2_Click()

    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 1)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstationr", 1)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableChangePassword", 1)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstation", 1)
    
    '删除注消
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 1)
    '运行
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1)
    
    

End Sub

Private Sub 屏蔽任务管理器_M33_Click()
On Error Resume Next
    Dim I As Long, j As Long, n As Long
    Dim Jssl As Integer
    Dim proc As PROCESSENTRY32
    Dim snap As Long
    Dim exename As String
    Dim item As ListItem
    Dim lngHwndProcess As Long
    Dim lngModules(1 To 200) As Long
    Dim lngCBSize2 As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim WKSize As Long
    Dim strProcessName As String
    Dim strComment As String   '装载进程注释的字符串
    Dim ProClass As String
    Dim SMJC_S As Boolean  '扫描进程
    '开始进程循环
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    proc.dwSize = Len(proc)
    theloop = ProcessFirst(snap, proc)
    I = 0
    n = 0
    While theloop <> 0
        I = I + 1
        'exename = proc.szExeFile
        lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, proc.th32ProcessID)
        If lngHwndProcess <> 0 Then
            lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
            If lngReturn <> 0 Then
                strModuleName = Space(MAX_PATH)
                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, 500)
                strProcessName = Left(strModuleName, lngReturn)
                strProcessName = CheckPath(Trim$(strProcessName))
                If strProcessName <> "" Then
                        '获取短文件名
                        exename = Dir(strProcessName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
                        If exename = "explorer.exe" Then
                            id = CLng(proc.th32ProcessID)
                            If id <> 0 Then
                                EndPro id
                            End If
                        End If
                    
                End If
            End If
        End If
        theloop = ProcessNext(snap, proc)
    Wend
    CloseHandle snap

End Sub




Private Sub 屏蔽任务管理器_M3_Click()
    Call DisableCtrlAltDelete(True)
End Sub

Private Sub 取消开机运行_M_Click()

    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "abc")

End Sub
Private Sub 取消屏蔽_M_Click()
    Close #1
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 0)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstationr", 0)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableChangePassword", 0)
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstation", 0)
    
    '删除注消
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0)
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0)

    Call DisableCtrlAltDelete(False)
    
    
End Sub

Private Sub 取消隐藏桌面图标_M_Click()
Dim wnd As Long
            wnd = FindWindow("Progman", vbNullString)
            wnd = FindWindowEx(wnd, 0, "ShellDll_DefView", vbNullString)
ShowWindow wnd, 5
End Sub

Private Sub 全部解锁_M_Click()

        txtPassword.Text = ""
        解锁_F.Caption = " 解 锁 "
        frmMain.解锁_F.Visible = True
        
End Sub


Private Sub 全部锁定_M_Click()

        XTSD_S_SD_Click
        关闭任务栏_M_Click
        关闭开始菜单_M_Click
        '屏蔽任务管理器_M_Click
        屏蔽任务管理器_M2_Click
        隐藏桌面图标_M_Click
        
        XTSD_S_JS.Enabled = False
        打开任务栏_M.Enabled = False
        打开开始菜单_M.Enabled = False
        取消屏蔽_M.Enabled = False
        取消隐藏桌面图标_M.Enabled = False
        修改密码_M.Enabled = False
        取消开机运行_M.Enabled = False
        mnuExit.Enabled = False
        mnuEndPro.Enabled = False
        mnuProcess.Enabled = False

End Sub

Private Sub 修改密码_M_Click()
txtPassword.Text = ""
解锁_F.Caption = " 修改密码 "
frmMain.解锁_F.Visible = True

End Sub

Private Sub 隐藏_M_Click()
    frmMain.Visible = False
    'frmMain.WindowState = 1
End Sub


Private Sub 隐藏桌面图标_M_Click()
'隐藏桌面图标_M.Checked = Not 隐藏桌面图标_M.Checked
Dim wnd As Long
            wnd = FindWindow("Progman", vbNullString)
            wnd = FindWindowEx(wnd, 0, "ShellDll_DefView", vbNullString)
ShowWindow wnd, 0
End Sub
Private Sub Mmqr_C_Click()
'密码信息文件的路径
Dim LoadFiles As String
Dim FilesTest As Boolean
Dim Cipher_Text As String

LoadFiles = Environ("windir") & "\sysset.dll"



'''''''''''''''''''''''''''''''''
'检验 key.dat 文件是否存在
If Dir(LoadFiles, vbHidden) = Empty Then
FilesTest = False
Else
FilesTest = True
End If
Filenum = FreeFile '提供一个尚未使用的文件号

'读取密码文件，把文件的信息赋值给 StrTarget 变量
Dim StrTarget As String
Open LoadFiles For Random As Filenum
Get #Filenum, 1, StrTarget
Close Filenum
If 解锁_F.Caption = " 解 锁 " Then
                            '如果 key.dat 文件已存在，则要求输入登录密码
                            
        If FilesTest = True Then
                                'Dim InputString As String
                                'InputString = InputBox("请你输入登录密码" & Chr(13) & Chr(13) & "万能密码：http://www.vbeden.com", "密码登录", InputString)
                            '将你输入的密码解密到 Plain_Text 变量
                Dim Plain_Text As String
                SubDecipher txtPassword, StrTarget, Plain_Text
        
                If txtPassword = Plain_Text Then
                
                        XTSD_S_JS.Enabled = True
                        打开任务栏_M.Enabled = True
                        打开开始菜单_M.Enabled = True
                        取消屏蔽_M.Enabled = True
                        取消隐藏桌面图标_M.Enabled = True
                        修改密码_M.Enabled = True
                        取消开机运行_M.Enabled = True
                        mnuExit.Enabled = True
                        mnuEndPro.Enabled = True
                        mnuProcess.Enabled = True
                        

                        
                        XTSD_S_JS_Click
                        打开任务栏_M_Click
                        打开开始菜单_M_Click
                        取消屏蔽_M_Click
                        取消隐藏桌面图标_M_Click
                        
                        
                        解锁_F.Visible = False
                    
                    Else
                
                        MsgBox "密码错误"
                        
                End If
        End If
    Else
    


                        SubCipher txtPassword.Text, txtPassword.Text, Cipher_Text
                        
                        '保存到文件并加密
                        Filenum = FreeFile
                        
                        Open LoadFiles For Random As Filenum
                        '把 Cipher_Text 的变量写入文件里
                        Put #Filenum, 1, Cipher_Text
                        Close Filenum
                        解锁_F.Visible = False

End If


If FilesTest = True Then SetAttr LoadFiles, vbHidden
End Sub
'加密子程序
Private Sub SubCipher(ByVal Password As String, ByVal From_Text As String, To_Text As String)
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim Str_len As Integer
Dim I As Integer
Dim ch As Integer

'得到了加密的数字
offset = NumericPassword(Password)

Rnd -1
'对随机数生成器做初始化的动作
Randomize offset

Str_len = Len(From_Text)
For I = 1 To Str_len
ch = Asc(Mid$(From_Text, I, 1))
If ch >= MIN_ASC And ch <= MAX_ASC Then
ch = ch - MIN_ASC
offset = Int((NUM_ASC + 1) * Rnd)
ch = ((ch + offset) Mod NUM_ASC)
ch = ch + MIN_ASC
To_Text = To_Text & Chr$(ch)
End If
Next I
End Sub

'解密子程序
Private Sub SubDecipher(ByVal Password As String, ByVal From_Text As String, To_Text As String)
Const MIN_ASC = 32 ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim Str_len As Integer
Dim I As Integer
Dim ch As Integer

offset = NumericPassword(Password)
Rnd -1
Randomize offset

Str_len = Len(From_Text)
For I = 1 To Str_len
ch = Asc(Mid$(From_Text, I, 1))
If ch >= MIN_ASC And ch <= MAX_ASC Then
ch = ch - MIN_ASC
offset = Int((NUM_ASC + 1) * Rnd)
ch = ((ch - offset) Mod NUM_ASC)
If ch < 0 Then ch = ch + NUM_ASC
ch = ch + MIN_ASC
To_Text = To_Text & Chr$(ch)
End If
Next I
End Sub

'将你输入的每个字符转换成密码数字
Private Function NumericPassword(ByVal Password As String) As Long
Dim Value As Long
Dim ch As Long
Dim Shift1 As Long
Dim Shift2 As Long
Dim I As Integer
Dim Str_len As Integer

'得到字符串内字符的数目
Str_len = Len(Password)
'给每个字符转换成密码数字
For I = 1 To Str_len
ch = Asc(Mid$(Password, I, 1))
Value = Value Xor (ch * 2 ^ Shift1)
Value = Value Xor (ch * 2 ^ Shift2)

Shift1 = (Shift1 + 7) Mod 19
Shift2 = (Shift2 + 13) Mod 23
Next I
NumericPassword = Value
End Function


Private Sub 注销_M_Click()

    lngVersion = GetVersion()
    
    If ((lngVersion And &H80000000) = 0) Then
        glngWhichWindows32 = mlngWindowsNT
    Else
        glngWhichWindows32 = mlngWindows95
    End If

    If glngWhichWindows32 = mlngWindowsNT Then
        AdjustToken
    End If
    
    ExitWindowsEx EWX_LogOff, 0  '注销
    
End Sub
