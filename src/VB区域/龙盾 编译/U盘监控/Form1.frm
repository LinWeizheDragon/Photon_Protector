VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10470
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3000
      Picture         =   "Form1.frx":169B1
      ScaleHeight     =   345
      ScaleWidth      =   510
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2400
      Picture         =   "Form1.frx":16EEA
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame MyFrame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   920
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   3135
      Begin VB.Frame BtnFrame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   3135
         Begin VB.Image IOut 
            Height          =   720
            Left            =   1800
            MouseIcon       =   "Form1.frx":174B9
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":1760B
            Top             =   80
            Width           =   1080
         End
         Begin VB.Image IFix 
            Height          =   720
            Left            =   960
            MouseIcon       =   "Form1.frx":178FD
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":17A4F
            Top             =   80
            Width           =   1080
         End
         Begin VB.Image IOpen 
            Height          =   720
            Left            =   120
            MouseIcon       =   "Form1.frx":17BD9
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":17D2B
            Top             =   80
            Width           =   1080
         End
         Begin VB.Image Image3 
            Height          =   900
            Left            =   0
            Picture         =   "Form1.frx":17EBB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3060
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   60
         Left            =   120
         Top             =   780
         Width           =   15
      End
      Begin VB.Label USBCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "剩余空间：100G"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   900
         Left            =   120
         Picture         =   "Form1.frx":18D73
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2940
      End
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":19C2B
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1A174
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1A6BD
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1AC06
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1B14F
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1B698
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   6
      Left            =   0
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1BBE1
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   7
      Left            =   0
      TabIndex        =   15
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1C12A
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   8
      Left            =   0
      TabIndex        =   16
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1C673
      PictureAlign    =   6
   End
   Begin 扫描模块.jcbutton USB 
      Height          =   975
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   9360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4210752
      Caption         =   "C:"
      ForeColor       =   16777215
      Picture         =   "Form1.frx":1CBBC
      PictureAlign    =   6
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "菜单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      MouseIcon       =   "Form1.frx":1D105
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   840
      X2              =   3960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   840
      X2              =   3840
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      MouseIcon       =   "Form1.frx":1D257
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":1D3A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "病毒防御助手 U盘监控"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "Form1.frx":33D5A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Const HTCAPTION = 2
  Private Const WM_NCLBUTTONDOWN = &HA1

Dim ShowBtn As Boolean
Dim DrivePathNow As String

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
MyFrame.Width = 0
    Me.Move ReadString("USBRTA", "Left", App.Path & "\Set.ini"), ReadString("USBRTA", "Top", App.Path & "\Set.ini")
    On Error Resume Next
    Dim rtn As Long
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA
    SetLayeredWindowAttributes hWnd, &HFF&, 0, LWA_COLORKEY
Debug.Print ReadString("USBRTA", "Left", App.Path & "\Set.ini")
End Sub



Private Sub IFix_Click()
StartMis DrivePathNow, 2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  '拖动窗体
  If Button = 1 Then '注意这里，你的无法实现就是这里的问题，没有定义鼠标按下
  ReleaseCapture
  SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  End If
   
WriteString "USBRTA", "left", Me.Left, App.Path & "\Set.ini"
WriteString "USBRTA", "Top", Me.Top, App.Path & "\Set.ini"

End Sub



Private Sub IOpen_Click()
StartMis DrivePathNow, 1
End Sub

Private Sub IOut_Click()
StartMis DrivePathNow, 3
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  '拖动窗体
  If Button = 1 Then '注意这里，你的无法实现就是这里的问题，没有定义鼠标按下
  ReleaseCapture
  SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
  End If
 
WriteString "USBRTA", "left", Me.Left, App.Path & "\Set.ini"
WriteString "USBRTA", "Top", Me.Top, App.Path & "\Set.ini"

End Sub



Private Sub Label2_Click()
PopupMenu frmMain.mnuPop
End Sub

'
'Dim MouseEnter     As Boolean
'Dim P As POINTAPI
'    GetCursorPos P '获取鼠标在屏幕中的位置
'    ScreenToClient Me.hwnd, P '转换为本窗体的坐标
'    Dim t As Boolean
'    t = P.X >= 0 And P.Y >= 0 And P.X < Me.Width / Screen.TwipsPerPixelX And P.Y <= Me.Height / Screen.TwipsPerPixelY
'    'If t Then Me.Caption = "x=" & P.x & "y=" & P.y '按像素显示坐标
'    'If t Then Me.Caption = "x=" & P.x * Screen.TwipsPerPixelX & "y=" & P.y * Screen.TwipsPerPixelY '按缇显示坐标
'
'  MouseEnter = (Image1.Left <= P.X * Screen.TwipsPerPixelX) And (P.X * Screen.TwipsPerPixelX <= Image1.Left + Image1.Width) And (Image1.Top <= P.Y * Screen.TwipsPerPixelY) And (P.Y * Screen.TwipsPerPixelY <= Image1.Top + Image1.Height)
'  If MouseEnter Then
'        Command1.Caption = "进入"
'        Command1.BackColor = &HFFC0FF
'        SetCapture Command1.hwnd
'  Else
'        Command1.Caption = "退出"
'        Command1.BackColor = &H8000000F
'        ReleaseCapture
'  End If


Public Function CallTip(ByVal TopVal As Integer)
MyFrame.Top = TopVal
MyFrame.Width = 0
SuperSleep 0.02
MyFrame.Width = 500
SuperSleep 0.02
MyFrame.Width = 1500
SuperSleep 0.02
MyFrame.Width = 2500
SuperSleep 0.02
MyFrame.Width = 3135
End Function

Private Sub USB1_Click()
ShowBtn = True
DrivePathNow = USB1.Caption
BtnFrame.Visible = True
CallTip USB1.Top
End Sub
Private Sub USB1_MouseEnter()
ShowBtn = False '不显示按钮
BtnFrame.Visible = False
CallTip USB1.Top
SetInfo USB1.Caption
End Sub
Private Sub USB1_MouseLeave()
If ShowBtn = False Then '如果不是显示按钮
MyFrame.Width = 0
End If
End Sub

Public Sub SetValue(ByVal Num)
On Error Resume Next
Dim now
Shape1.Width = Num * 2900
now = Shape1.Width / 2900

'If now < Num Then '如果需要增加
'Do Until now >= Num
'now = now + 0.01
'If now > 1 Then now = 1
'Shape1.Width = now * 2900
'SuperSleep 0.01
'If 0.5 > now And now > 0.25 Then
'Shape1.BackColor = &HFFFF&
'ElseIf 0.75 > now And now > 0.5 Then
'Shape1.BackColor = &H80FF&
'ElseIf now > 0.75 Then
'Shape1.BackColor = &HC0&
'ElseIf now < 0.25 Then
'Shape1.BackColor = &HFF00&
'End If
'Loop
'Else '如果需要减少
'Do Until now <= Num
'now = now - 0.01
'If now < 0 Then now = 0
'Shape1.Width = now * 2900
'SuperSleep 0.01
If 0.5 > now And now > 0.25 Then
Shape1.BackColor = &HFFFF&
ElseIf 0.75 > now And now > 0.5 Then
Shape1.BackColor = &H80FF&
ElseIf now > 0.75 Then
Shape1.BackColor = &HC0&
ElseIf now < 0.25 Then
Shape1.BackColor = &HFF00&
End If





End Sub


Public Function SetInfo(ByVal DrivePath As String)
On Error GoTo err:
USBCaption.Caption = "获取信息" & vbCrLf & ">>>>>"
Dim MyFSO As New FileSystemObject
Dim FreeStr, VolumStr
VolumStr = DrivePath & "(" & MyFSO.GetDrive(DrivePath).VolumeName & ")"
FreeStr = "剩余空间：" & KillNum(MyFSO.GetDrive(DrivePath).FreeSpace)
USBCaption.Caption = VolumStr & vbCrLf & vbCrLf & FreeStr
SetValue (MyFSO.GetDrive(DrivePath).TotalSize - MyFSO.GetDrive(DrivePath).FreeSpace) / MyFSO.GetDrive(DrivePath).TotalSize
Exit Function
err:
USBCaption.Caption = "获取信息失败，请检查磁盘" & vbCrLf & vbCrLf & ">>>>>"
SetValue 1

End Function

Public Function KillNum(ByVal Num)

If Num < 1024 Then '小于1024B
KillNum = Round(Num, 2) & " B"
Exit Function
End If
Num = Num / 1024
If Num < 1024 Then '小于1024KB
KillNum = Round(Num, 2) & " KB"
Exit Function
End If
Num = Num / 1024
If Num < 1024 Then '小于1024MB
KillNum = Round(Num, 2) & " MB"
Exit Function
End If
Num = Num / 1024
If Num < 1024 Then '小于1024GB
KillNum = Round(Num, 2) & " GB"
Exit Function
End If
Num = Num / 1024
If Num < 1024 Then '小于1024TB
KillNum = Round(Num, 2) & " TB"
Exit Function
End If
KillNum = Round(Num, 2) & " TB"
End Function

Private Sub USB_Click(Index As Integer)
ShowBtn = True
DrivePathNow = USB(Index).Caption
BtnFrame.Visible = True
CallTip USB(Index).Top
End Sub

Private Sub USB_MouseEnter(Index As Integer)
ShowBtn = False '不显示按钮
BtnFrame.Visible = False
CallTip USB(Index).Top
SetInfo USB(Index).Caption
End Sub

Private Sub USB_MouseLeave(Index As Integer)
If ShowBtn = False Then '如果不是显示按钮
MyFrame.Width = 0
End If
End Sub
