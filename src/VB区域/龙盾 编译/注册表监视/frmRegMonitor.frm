VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmRegMonitor 
   Caption         =   "注册表监视"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmRegMonitor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5355
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   600
      Width           =   2295
   End
   Begin VB.CheckBox chkAllow 
      Caption         =   "不再提示，以后都这样处理"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Timer timerCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   2400
   End
   Begin ComctlLib.ProgressBar proBar 
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   5160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   30
   End
   Begin VB.OptionButton optDisaccord 
      Caption         =   "不同意修改"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton optAgree 
      Caption         =   "同意修改"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   4920
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Frame frameReg 
      Caption         =   "注册表监视"
      Height          =   2245
      Left            =   1680
      TabIndex        =   6
      Top             =   4800
      Width           =   5625
      Begin VB.TextBox txtProcessPath 
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Top             =   1760
         Width           =   4095
      End
      Begin VB.TextBox txtType 
         Height          =   270
         Left            =   1320
         TabIndex        =   1
         Top             =   1290
         Width           =   4095
      End
      Begin VB.TextBox txtRegPath 
         Height          =   775
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   300
         Width           =   4095
      End
      Begin VB.Label lblProcessPath 
         AutoSize        =   -1  'True
         Caption         =   "进程信息:"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label lType 
         AutoSize        =   -1  'True
         Caption         =   "键值/类型:"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lPath 
         AutoSize        =   -1  'True
         Caption         =   "注册表路径:"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   1200
      Top             =   1680
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuPopMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "退出程序"
      End
   End
End
Attribute VB_Name = "frmRegMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InstallRegHook Lib "RegistryInfo.dll" (ByVal strCheck As String) As Long
Private Declare Function UninstallRegHook Lib "RegistryInfo.dll" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private mintSum As Integer

Private Sub cmdOK_Click()
    timerCheck.Enabled = False  '停止记时
    mintSum = 0 '计数归0
    Me.proBar.Value = 0 '进度条进度归0
    gblnIsShow = False '设置不显示窗体标志状态
    'Me.Hide '隐藏窗体
End Sub

Private Sub Command1_Click()
MsgBox CheckProcess("C:\Documents and Settings\Administrator\桌面\光子防御网 编译\注册表监视\测试添加启动项.exe", "呵呵呵呵呵")
End Sub

Private Sub Command2_Click()
MsgBox ProcessScan("C:\Documents and Settings\Administrator\桌面\光子防御网 编译\注册表监视\测试添加启动项.exe")
End Sub

Private Sub Form_Initialize()
    If App.PrevInstance Then
    End '重复加载就直接退出
    End If
    InitCommonControls
    App.TaskVisible = False
    
End Sub

Private Sub Form_Load()
   Load frmData
   Load frmRec
    strIniFilePath = App.Path & "\Rules.ini" '设置设置文件路径

Call WriteString("RegRules", """" & App.Path & "\ProcessRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\RegRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\USBRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\ProgramUpdate.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\DragonShield.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\PRMonitor.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\Tools\FixFolders\龙盾-移动盘隐藏文件夹修复工具.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\Tools\KillFiles\KillFile.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\Tools\RegMonitor\RegTools.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\PhotonRepair.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\PhotonClear.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\PhotonMajorization.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\ProtectProcess.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\Protect.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\NetScanner\Photon-NetScanner.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\NetScanner\koemsec1.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\NetScanner\kdumprep.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("RegRules", """" & App.Path & "\NetScanner\kdumpfix.exe""", "1", App.Path & "\Rules.ini")



    'Me.Hide '隐藏主窗体
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE '最前端显示
    StartHook Me.hwnd '消息钩子主要是获取DLL传来的消息 ,消息名是WM_COPYDATA
   ' SendToTray '添加托盘图标
    InstallRegHook "http://blog.csdn.net/chenhui530/" '安装全局API钩子
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
Me.Hide
End Sub



Private Sub Form_Unload(Cancel As Integer)
  '  DeleteSysTray '删除托盘
    Unhook Me.hwnd '卸载消息钩子
    UninstallRegHook '卸载API钩子
    Dim I As Form
For Each I In Forms
Unload I
Next

End Sub

Private Sub mnuExit_Click()
    Erase gstrArray '清空消息
    gblnIsEnd = True '确定退出状态
    cmdOK_Click '使本次生效并且关闭记时器控件等
    Unload Me '卸载窗体准备退出
End Sub

Private Sub timerCheck_Timer()
    If mintSum >= 30 Then '当等于30秒时就隐藏界面
        timerCheck.Enabled = False
        mintSum = 0
        gblnIsShow = False
        Me.Hide
    End If
    mintSum = mintSum + 1 '增加计数当大于等于30时隐藏界面
    Me.proBar.Value = mintSum '显示进度
End Sub

Private Sub txtProcessPath_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '不允许输入
End Sub

Private Sub txtRegPath_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '不允许输入
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '不允许输入
End Sub
