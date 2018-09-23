VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "全盘扫描"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "frmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5250
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "金山云安全联网查杀"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "光子防御离线查杀"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.DoScan False, True, ""
Unload Me
End Sub

Private Sub Command2_Click()
If Not Dir(App.Path & "\NetScanner\Photon-NetScanner.exe") = "" Then
MsgBox "请手动运行“Photon-NetScanner.exe”程序以防服务加载失败！"
Shell "explorer.exe /select, """ & App.Path & "\NetScanner\Photon-NetScanner.exe""", vbNormalFocus
Else
Dim frmText As New frmMsg
frmText.Label1.Caption = "启动失败！您可能没有正确安装本软件！（错误号1000，找不到所需组件）"
frmText.Show
End If
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "光子防御网全盘查杀有两种模式：" & vbCrLf & _
                 "1.光子防御离线查杀，使用离线病毒库，查杀率较低，在没有网络的情况下提供服务。" & vbCrLf & _
                 "2.金山云安全联网查杀，使用金山云安全开放平台提供的API查杀，查杀率极高，但必须连接网络。"
                 
End Sub
