VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   0  'None
   Caption         =   "病毒防御助手"
   ClientHeight    =   6000
   ClientLeft      =   8100
   ClientTop       =   4020
   ClientWidth     =   5385
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   0
      Picture         =   "frmTip.frx":169B1
      ScaleHeight     =   5985
      ScaleWidth      =   5385
      TabIndex        =   1
      Top             =   -23
      Width           =   5415
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4815
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "以后也这样处理（不想麻烦的话）"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   3960
            Width           =   3615
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2760
            Top             =   4320
         End
         Begin VB.PictureBox PicIcon 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   480
            Picture         =   "frmTip.frx":23E3B
            ScaleHeight     =   675
            ScaleWidth      =   750
            TabIndex        =   9
            Top             =   360
            Width           =   750
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1440
            Width           =   4215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "确定"
            Height          =   375
            Left            =   3480
            TabIndex        =   5
            Top             =   4200
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "终止进程"
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            Top             =   3600
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "允许此操作"
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "再过 30 秒自动替您选择"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   4320
            Width           =   2415
         End
         Begin VB.Label Tip 
            BackColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   1440
            TabIndex        =   7
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "信息："
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            X1              =   360
            X2              =   4440
            Y1              =   3480
            Y2              =   3480
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "可疑进程创建"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChooseNum As Integer
Public Choose As Boolean
'ChooseNum
'1:允许
'2:拒绝
'3:信任
'4:黑名单
Dim i As Integer


'窗体移动API声明


  Private Declare Function ReleaseCapture Lib "user32" () As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Const HTCAPTION = 2
  Private Const WM_NCLBUTTONDOWN = &HA1
'获取任务栏高度
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Private Sub Back_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 '拖动窗体
  If Button = 1 Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  End If
End Sub

Private Sub Command1_Click()
Choose = False
Unload Me
End Sub
Private Sub Command2_Click()
Choose = True
Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
Dim lRes As Long
    Dim rectVal As RECT
    Dim TaskbarHeight As Integer '任务栏高度
lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
TaskbarHeight = Screen.Height - rectVal.Bottom * Screen.TwipsPerPixelY
 Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - TaskbarHeight, Me.Width, Me.Height
i = 30
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check1.Value = 1 Then
If Option1.Value = True Then '判断选择
  ChooseNum = 3
Else
  ChooseNum = 4
End If
Else
If Option1.Value = True Then '判断选择
  ChooseNum = 1
Else
  ChooseNum = 2
End If
End If
DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub

Private Sub Form_Activate()
Call rgnform(Me, 5, 5) '调用子过程
End Sub

Private Sub Timer1_Timer()
i = i - 1
Label2.Caption = "再过 " & i & " 秒自动替您选择"
If i = 0 Then
Unload Me
Timer1.Enabled = False
End If

End Sub
