VERSION 5.00
Begin VB.Form Form_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病毒防御助手进程保护"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "Form_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10395
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   435
      Left            =   1305
      TabIndex        =   3
      Top             =   2145
      Width           =   1395
   End
   Begin VB.TextBox Text_Memory 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   3870
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form_Map.frx":169B1
      Top             =   390
      Width           =   6195
   End
   Begin VB.CommandButton cmdReadMem 
      Caption         =   "读共享内存变量"
      Height          =   435
      Left            =   2085
      TabIndex        =   2
      Top             =   1365
      Width           =   1575
   End
   Begin VB.CommandButton cmdWriteMem 
      Caption         =   "写共享内存变量"
      Height          =   435
      Left            =   165
      TabIndex        =   1
      Top             =   1365
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "共享内存变量值："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   660
      Width           =   1935
   End
End
Attribute VB_Name = "Form_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProtectObj(10) As String

Private sData As String
Const LenStr As Long = 65535 * 10
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const PROCESS_TERMINATE = &H1
Private Const TH32CS_SNAPHEAPLIST = 1
Private Const TH32CS_SNAPPROCESS = 2
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
    szExeFile As String * 500
End Type
Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Function ReadMem() As String
    sData = String(LenStr, vbNullChar)
    Call lstrcpyn(ByVal sData, ByVal lShareData, LenStr)
    Text_Memory.Text = ""
    Text_Memory.SelText = sData
    ReadMem = sData
End Function

Public Function WriteMem(ByVal WriteText)
    sData = WriteText
    Call lstrcpyn(ByVal lShareData, ByVal sData, LenStr)
End Function

Private Sub Command1_Click()
Dim Text
Text = "Protect::" & InputBox("...")
Form_Map.WriteMem Split(Text, "::")(1)
SuperSleep 1.5
Form_Map.WriteMem "Wait"
End Sub

Private Sub Command2_Click()
ReLoad
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "已经开启自我保护，无需再次开启"
Unload Me
End If
App.TaskVisible = False
ProtectObj(0) = App.Path & "\PhotonProtect.exe"
ProtectObj(1) = App.Path & "\ProcessRTA.exe"
ProtectObj(2) = App.Path & "\PRMonitor.exe"
ProtectObj(3) = App.Path & "\RegRTA.exe"
ProtectObj(4) = App.Path & "\ScanMod.exe"
ProtectObj(5) = App.Path & "\Protect.exe"
ProtectObj(6) = App.Path & "\ProtectProcess.exe"
'ProtectObj(7) = App.Path & "\PhotonProtect.exe"
'ProtectObj(8) = App.Path & "\PhotonProtect.exe"
'ProtectObj(9) = App.Path & "\PhotonProtect.exe"
'ProtectObj(10) = App.Path & "\PhotonProtect.exe"
    Text_Memory.Text = ""
    hMemShare = CreateFileMapping(&HFFFFFFFF, _
        0, _
        PAGE_READWRITE, _
         0, _
        LenStr, _
        "PPProtectSelf")
    If hMemShare = 0 Then
        'Err.LastDllError
        MsgBox "创建内存映射文件失败!", vbCritical, "错误"
    End If
    If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
        '指定内存文件已存在
    End If
    
    
    
    
    
    
    
    
    
    lShareData = MapViewOfFile(hMemShare, FILE_MAP_WRITE, 0, 0, 0)
    If lShareData = 0 Then
        MsgBox "为映射文件对象创建视失败!", vbCritical, "错误"
    End If
    'Debug.Print "lShareData="; lShareData
    'sData = String(LenStr, &H0)
    ReadMem
    Call WriteMem("Wait")
    OpenFile App.Path & "\ProtectProcess.exe"
    Load frmRec
    Me.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Text
Text = "Protect::Unload"
Form_Map.WriteMem Split(Text, "::")(1)
SuperSleep 1.5
Form_Map.WriteMem "Wait"

    If lShareData <> 0 Then
        Call UnmapViewOfFile(ByVal lShareData) '解除映射
        lShareData = 0
    End If
    
    If hMemShare <> 0 Then
        Call CloseHandle(hMemShare) '关闭映射
        hMemShare = 0
    End If
    Unload frmRec
End Sub


Public Function ReLoad()
On Error Resume Next
    Dim uProcess As PROCESSENTRY32
    Dim mSnapShot As Long
    Dim mName As String
    Dim I As Integer
    Dim Num As Integer
   Dim mresult
    Dim Msg As String
    DoEvents
    '获取进程长度？？
    uProcess.dwSize = Len(uProcess)
    '创建一个系统快照
    mSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If mSnapShot Then
        '获取第一个进程
        mresult = ProcessFirst(mSnapShot, uProcess)
        '失败则返回false
        Do While mresult
            '返回进程长度值+1，Chr(0)的作用：结束语，防止修改进程
            I = InStr(1, uProcess.szExeFile, Chr(0))
            '转换成小写
            mName = LCase$(Left$(uProcess.szExeFile, I - 1))
            '在listview控件中添加当前进程名
            '添加进程名
            For Num = 0 To UBound(ProtectObj)
            If LCase(GetProcessPath(mName)) = LCase(ProtectObj(Num)) Then
            '如果找到相同的进程
            WriteMem uProcess.th32ProcessID
            SuperSleep 1
            WriteMem "Wait"
            SuperSleep 0.5
            End If
            Next
            '获取下一个进程
            mresult = ProcessNext(mSnapShot, uProcess)
        Loop
    Else
        ErrMsgProc (Msg)
    End If

End Function
Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub


Public Function OpenFile(ByVal OpenName As String, Optional ByVal InitDir As String = vbNullString, Optional ByVal msgStyle As ShowStyle = vbNormalFocus)
    ShellExecute 0&, vbNullString, OpenName, vbNullString, InitDir, msgStyle
End Function

