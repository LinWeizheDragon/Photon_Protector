VERSION 5.00
Begin VB.Form frmUSB 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "不了，谢谢"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修复"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin 扫描模块.jcbutton jcbutton1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "消息"
      Picture         =   "frmUSB.frx":0000
      PictureAlign    =   5
   End
   Begin VB.Label USB 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Tip 
      BackStyle       =   0  'Transparent
      Caption         =   "检测到以下移动设备插入，点击修复修复文件"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmUSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Time As Integer
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

Private Sub Command1_Click()
On Error Resume Next
Me.Hide
Dim Path
Dim MyFSO As New FileSystemObject
For i = 0 To frmMain.List.ListCount - 1
If Dir(frmMain.List.List(i)) <> "" Then
ShowFolderList frmMain.List.List(i)
End If
Next
For x = 0 To frmMain.List2.ListCount - 1
Path = frmMain.List2.List(x)
Debug.Print Path
Call SetAttr(Path, vbNormal)
FileCopy App.Path & "\desktop.ini", Path & "\desktop.ini"
FileCopy App.Path & "\文件夹安全验证图标.ico", Path & "\文件夹安全验证图标.ico"
Call SetAttr(Path, vbSystem)
Debug.Print App.Path & "\desktop.ini" & "----" & Path & "\desktop.ini"
Next
Unload Me
End Sub

Private Sub Command2_Click()
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.List.Clear
frmMain.List2.Clear
End Sub
