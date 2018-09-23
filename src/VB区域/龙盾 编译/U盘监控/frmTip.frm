VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTip 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   8100
   ClientTop       =   4020
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   $"frmTip.frx":0000
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
      Picture         =   "frmTip.frx":0008
      ScaleHeight     =   3465
      ScaleWidth      =   7470
      TabIndex        =   1
      Top             =   0
      Width           =   7500
      Begin VB.PictureBox PicIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1080
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   7095
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2520
            Top             =   1680
         End
         Begin MSComctlLib.ListView ListView 
            Height          =   1575
            Left            =   240
            TabIndex        =   6
            Top             =   0
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2778
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CommandButton Command3 
            Caption         =   "暂不处理"
            Height          =   375
            Left            =   4080
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   3120
            Top             =   1680
         End
         Begin VB.CommandButton Command2 
            Caption         =   "处理(默认)"
            Height          =   375
            Left            =   5640
            TabIndex        =   3
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "再过 30 秒自动替您选择"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1800
            Width           =   2415
         End
      End
      Begin VB.Label Tip 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1920
         TabIndex        =   8
         Top             =   720
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
         TabIndex        =   7
         Top             =   3480
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChooseMod As Boolean
'True ：处理
'False：暂不处理
Public Choose As Boolean
'ChooseNum
'1:允许
'2:拒绝
'3:不再阻止
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
Choose = False
ChooseMod = True
Unload Me
End Sub
Private Sub Command2_Click()
DoEvents
Me.Hide
If ListView.ListItems.Count = 0 Then Exit Sub
For xy = 1 To ListView.ListItems.Count
 If ListView.ListItems(xy).Checked = True Then
 Call SetAttr(ListView.ListItems(xy).Text, vbNormal)
 Call Kill(ListView.ListItems(xy).Text)
 End If
Next
Choose = True
ChooseMod = True
Unload Me
End Sub

Private Sub Command3_Click()
Choose = True
ChooseMod = False '暂不处理
Unload Me
End Sub


Private Sub Form_Load()
I = 30
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
DoEvents
Me.Hide
On Error Resume Next
If ChooseMod = True Then
For y = 1 To ListView.ListItems.Count
Set itm = ListView.ListItems(y)
If itm.Checked = True Then
 If Dir(itm.Text, vbDirectory) <> "" Then '如果是文件夹
   Call SetAttr(itm.Text, vbNormal)
 Else '不是文件夹
   If Dir(itm.Text, vbSystem Or vbHidden Or vbNormal Or vbReadOnly) <> "" Then '如果存在这个文件
     Call SetAttr(itm.Text, vbNormal)
     Kill itm.Text
   End If
 End If
End If
Next
End If
DeleteObject outrgn '将圆角区域使用的所有系统资源释放
End Sub

Private Sub Form_Activate()
Call rgnform(Me, 10, 10) '调用子过程
End Sub



Private Sub ListView_ItemClick(ByVal item As MSComctlLib.ListItem)
PicIcon = GetIconFromFile(item.Text, 0, True)
Tip.Caption = item.SubItems(2)
End Sub

Private Sub Timer1_Timer()

I = I - 1
Label2.Caption = "再过 " & I & " 秒自动替您选择"

If I = 0 Then
Unload Me
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
With Me
Dim FilePath As String '将列于第一个的文件的图标显示，显示病毒描述
FilePath = .ListView.ListItems(1).Text
Dim StrDes As String
StrDes = .ListView.ListItems(1).SubItems(2)
If StrDes <> "" Then .Tip.Caption = StrDes
If FilePath <> "" Then
.PicIcon.Picture = GetIconFromFile(FilePath, 0, True)
Else
.PicIcon.Visible = False
.PicIcon.AutoRedraw = True
End If
Timer2.Enabled = False
End With
For x = 1 To ListView.ListItems.Count
ListView.ListItems(x).Checked = True
Next
End Sub

