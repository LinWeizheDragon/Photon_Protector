VERSION 5.00
Begin VB.Form frmFix 
   Caption         =   "修复映像劫持及文件关联"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "frmFix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6525
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateToolhelpSnapshot Lib "KERNEL32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "KERNEL32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "KERNEL32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = TH32CS_SNAPPROCESS + TH32CS_SNAPHEAPLIST + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE

Private Const PROCESS_TERMINATE = 1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const MAX_PATH = 260

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
    szExeFile As String * MAX_PATH
End Type

Public Function IsProcessRun(exeName As String) As Boolean
    Dim snap As Long, Ret As Long, lProcess As Long
    Dim proc As PROCESSENTRY32
    Dim mName As String * MAX_PATH, modName As String

    snap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    proc.dwSize = Len(proc)
    Ret = ProcessFirst(snap, proc)

    Do While Ret <> 0
        mName = ""
        Dim Modules(1 To MAX_PATH) As Long, cbMNeeded As Long
        lProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, True, proc.th32ProcessID)
        If lProcess <> 0 Then
            Ret = EnumProcessModules(lProcess, Modules(1), MAX_PATH, cbMNeeded)
            If Ret <> 0 Then Ret = GetModuleFileNameExA(lProcess, Modules(1), mName, Len(mName))
            modName = Trim(Left(mName, Ret))
            If InStr(LCase(modName), LCase(exeName)) Then
                CloseHandle snap
                IsProcessRun = True
                Exit Function
            End If
        End If
        Ret = ProcessNext(snap, proc)
    Loop
    CloseHandle snap
    IsProcessRun = False
End Function



Private Sub Command1_Click()
Shell App.Path & "\FixImage.exe /Quiet"
Text1.Text = Text1.Text & vbCrLf & "开始修复......请放行执行进程"
SuperSleep 1
Do Until IsProcessRun("FixImage.exe") = False
SuperSleep 1
Loop
Text1.Text = Text1.Text & vbCrLf & "修复完成......"
SuperSleep 1
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "即将开始修复映像劫持和文件关联，点击开始继续，点击取消结束"
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub
