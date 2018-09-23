VERSION 5.00
Begin VB.Form frmVirusTip 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   8100
   ClientTop       =   4020
   ClientWidth     =   7515
   Icon            =   "frmVirusTip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   $"frmVirusTip.frx":169B1
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      Picture         =   "frmVirusTip.frx":169B9
      ScaleHeight     =   3465
      ScaleWidth      =   7470
      TabIndex        =   1
      Top             =   0
      Width           =   7500
      Begin VB.TextBox TextRes 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmVirusTip.frx":1F282
         Top             =   1680
         Width           =   5775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "处理(默认)"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3480
         Top             =   2760
      End
      Begin VB.CommandButton Command3 
         Caption         =   "暂不处理"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2880
         Top             =   2760
      End
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   600
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "再过 30 秒自动替您选择"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Tip 
         BackColor       =   &H00FFFFFF&
         Caption         =   "发现病毒"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   3480
         TabIndex        =   2
         Top             =   3480
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmVirusTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChooseMod As Boolean
'True ：处理
'False：暂不处理
Public Choose As Boolean
Dim I As Integer

'窗体移动API声明
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

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
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  End If
End Sub

Private Sub Command1_Click()
ChooseMod = True
Unload Me
End Sub
Private Sub Command2_Click()
ChooseMod = True
Unload Me
End Sub

Private Sub Command3_Click()
ChooseMod = False '暂不处理
Unload Me
End Sub


Private Sub Form_Load()
I = 30
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub

Private Sub Form_Activate()
Call rgnform(Me, 10, 10) '调用子过程
End Sub



Private Sub Timer1_Timer()

I = I - 1
Label2.Caption = "再过 " & I & " 秒自动替您选择"

If I = 0 Then
ChooseMod = True
Unload Me
Timer1.Enabled = False
End If

End Sub



