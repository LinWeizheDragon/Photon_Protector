VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   6960
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "自动更新"
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "开始更新"
         Height          =   495
         Left            =   4320
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label LblStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "等待中"
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   600
      Top             =   2760
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "自动更新"
      BeginProperty Font 
         Name            =   "方正综艺简体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3820
      TabIndex        =   1
      Top             =   220
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "自动更新"
      BeginProperty Font 
         Name            =   "方正综艺简体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Enabled = False
Dim lnginet As Long
    Dim lnginetconn As Long
    Dim blnRC As Boolean
lnginet = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, _
vbNullString, vbNullString, 0&)
If lnginet Then
  lnginetconn = InternetConnect(lnginet, "174.128.236.169", 21, "54473", "linweizhe", 1, 0, 0)
    If lnginetconn Then
       blnRC = FtpGetFile(lnginetconn, "/联机信息/info.ini", App.Path & "\Update.ini", 0, 0, 1, 0)
         If blnRC Then
            LblStatus.Caption = "配置文件下载成功"
         End If
       InternetCloseHandle lnginetconn
       InternetCloseHandle lnginet
    Else
       MsgBox "连接失败，程序自动关闭，请重新再试", vbOKOnly, "FTP服务器连接失败"
       'End
    End If
Else
       MsgBox "连接错误，程序自动关闭，请重新再试", vbOKOnly, "FTP服务器连接错误"
        'End
End If

Dim Version As String
Dim versioninfo As String
Version = ReadString("VersionInfo", "Main", App.Path & "\Update.ini")
Dim fver As String
Dim fso As FileSystemObject
Set fso = New FileSystemObject
fver = fso.GetFileVersion(App.Path & "\DragonSheild.exe") '文件路径
versioninfo = fver
LblStatus.Caption = "最新版本：" & Version & vbCrLf & "当前版本：" & versioninfo
If versioninfo = Version Then
MsgBox "无需更新！"
End
Else

End Sub

Private Sub Form_Load()
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

End Sub
