VERSION 5.00
Begin VB.UserControl USBList 
   BackColor       =   &H000000FF&
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   3225
   ScaleWidth      =   6000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   60
      Left            =   120
      Top             =   800
      Width           =   15
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   0
      Picture         =   "USBList.ctx":0000
      Top             =   2280
      Width           =   3540
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   0
      Picture         =   "USBList.ctx":1140
      Top             =   1920
      Width           =   3540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image IOut 
      Height          =   720
      Left            =   2550
      MouseIcon       =   "USBList.ctx":22D7
      MousePointer    =   99  'Custom
      Picture         =   "USBList.ctx":2429
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Image IFix 
      Height          =   720
      Left            =   1800
      MouseIcon       =   "USBList.ctx":271B
      MousePointer    =   99  'Custom
      Picture         =   "USBList.ctx":286D
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Image IOpen 
      Height          =   720
      Left            =   1050
      MouseIcon       =   "USBList.ctx":29F7
      MousePointer    =   99  'Custom
      Picture         =   "USBList.ctx":2B49
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Image Scroll 
      Height          =   900
      Left            =   0
      Picture         =   "USBList.ctx":2CD9
      Top             =   960
      Width           =   3540
   End
   Begin VB.Label Label2 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "USBList.ctx":3B91
      Top             =   0
      Width           =   3540
   End
End
Attribute VB_Name = "USBList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Scrolling As Boolean
Dim ScrolOver As Boolean
Const Topk = 60 '滚动时ImageList和背景高度差
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function ReleaseCapture Lib "user32" () As Long



Private Sub IFix_Click()
StartMis Left(Label3.Caption, 2), 2
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Form1
If ScrolOver = False Then '还没滚动完
If Timer1.Enabled = False Then '计时器还没开启
Timer1.Enabled = True

End If
End If
End With


End Sub


Private Function Follow()
'跟随
On Error Resume Next
IOpen.Top = Scroll.Top + Topk
IFix.Top = Scroll.Top + Topk
IOut.Top = Scroll.Top + Topk
Label3.Top = Scroll.Top + 240
End Function



Private Sub IOpen_Click()
StartMis Left(Label3.Caption, 2), 1
End Sub

Private Sub IOut_Click()
StartMis Left(Label3.Caption, 2), 3
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Do Until Scroll.Top = 0
SuperSleep 0.00001
Scroll.Top = Scroll.Top + 5
Follow
Loop
ScrolOver = True '滚动完了
Timer1.Enabled = False

End Sub



Private Sub UserControl_Initialize()
Scroll.Top = 0 - Scroll.Height  '隐藏
Follow
ScrolOver = False '还没滚动

End Sub
Public Function SetText(ByVal First As String, ByVal Second As String, ByVal Path As String)
Label2.Caption = Second
Label3.Caption = Path

End Function
Public Function ReSet()
Timer1.Enabled = False
Scroll.Top = 0 - Scroll.Height  '隐藏
Follow
ScrolOver = False '还没滚动
End Function
Public Sub SetType(ByVal DiskType As Integer)
If DiskType = 1 Then '移动硬盘
Image1.Picture = Image2.Picture
Else
Image1.Picture = Image3.Picture
End If
End Sub

Public Sub SetValue(ByVal Num)
On Error Resume Next
Dim now
now = Shape1.Width / 3300
If now < Num Then '如果需要增加
Do Until now >= Num
now = now + 0.01
If now > 1 Then now = 1
Shape1.Width = now * 3300
SuperSleep 0.01
If 0.5 > now And now > 0.25 Then
Shape1.BackColor = &HFFFF&
ElseIf 0.75 > now And now > 0.5 Then
Shape1.BackColor = &H80FF&
ElseIf now > 0.75 Then
Shape1.BackColor = &HC0&
ElseIf now < 0.25 Then
Shape1.BackColor = &HFF00&
End If
Loop
Else '如果需要减少
Do Until now <= Num
now = now - 0.01
If now < 0 Then now = 0
Shape1.Width = now * 3300
SuperSleep 0.01
If 0.5 > now And now > 0.25 Then
Shape1.BackColor = &HFFFF&
ElseIf 0.75 > now And now > 0.5 Then
Shape1.BackColor = &H80FF&
ElseIf now > 0.75 Then
Shape1.BackColor = &HC0&
ElseIf now < 0.25 Then
Shape1.BackColor = &HFF00&
End If
Loop
End If


End Sub
