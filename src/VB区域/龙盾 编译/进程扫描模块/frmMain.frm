VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmMain 
   Caption         =   "龙盾-高级实时监控"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   3975
   ScaleWidth      =   6135
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "监控本地设备进程创建"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "监控移动设备进程创建"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "选项"
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "关闭"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Status 
         BackColor       =   &H00FFFFFF&
         Caption         =   "启动中……"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "状态："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   0
      Top             =   2520
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "实时监控"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Begin VB.Menu mnuShow 
         Caption         =   "显示主界面"
      End
      Begin VB.Menu Split 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "关闭防护"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
DoEvents
WriteString "ProcessRTA", "Removable", Check1.Value, App.Path & "\Set.ini"
End Sub

Private Sub Check2_Click()
DoEvents
WriteString "ProcessRTA", "HardDisk", Check2.Value, App.Path & "\Set.ini"

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
 MsgBox "已经开启本程序，禁止重复开启！否则会造成不可预知的系统严重错误。", vbOKOnly, "龙盾-高级实时防护"
 End
End If
App.TaskVisible = False
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

Load frmData
Load frmHookCreate
Load frmRec
Me.Hide
frmHookCreate.StartHookFunction
'CreatTray Me, "龙盾-高级实时监控", "龙盾-高级实时监控", "程序已开启", 4

Dim Removeable
Dim Harddisk
Removeable = ReadString("ProcessRTA", "Removable", App.Path & "\Set.ini")
Harddisk = ReadString("ProcessRTA", "HardDisk", App.Path & "\Set.ini")
If Removeable <> "" Then
If Removeable = 1 Then Check1.Value = 1
End If
If Harddisk <> "" Then
If Harddisk = 1 Then Check2.Value = 1
End If
Me.Hide
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Me.WindowState = vbNormal
Me.Hide
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim i As Form
For Each i In Forms
Unload i
Next



End Sub
