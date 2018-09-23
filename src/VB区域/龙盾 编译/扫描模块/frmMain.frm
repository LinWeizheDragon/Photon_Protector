VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.Controls.v15.3.1.ocx"
Begin VB.Form frmMain 
   Caption         =   "光子防御网-安全扫描"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":169B1
   ScaleHeight     =   8205
   ScaleWidth      =   11955
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Time 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1600
      Width           =   11775
      Begin VB.CommandButton Command6 
         Caption         =   "反选"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "全选"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   6000
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "详细"
         Height          =   1695
         Left            =   2880
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.Label Lbl_Tip 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label4"
            Height          =   1335
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "关闭"
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "开始处理，结束扫描"
         Height          =   375
         Left            =   9360
         TabIndex        =   10
         Top             =   5880
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListVirus 
         Height          =   3855
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "计算机扫描"
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   11535
         Begin VB.CommandButton Command5 
            Caption         =   "停止"
            Height          =   375
            Left            =   10560
            TabIndex        =   17
            Top             =   1250
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "暂停"
            Height          =   375
            Left            =   9600
            TabIndex        =   16
            Top             =   1250
            Width           =   855
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "信息"
            Height          =   975
            Left            =   8520
            TabIndex        =   15
            Top             =   120
            Width           =   2895
            Begin VB.Label Lbl_VirusNum 
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   255
               Left            =   1080
               TabIndex        =   21
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "发现威胁："
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Lbl_Time 
               BackStyle       =   0  'Transparent
               Caption         =   "00:00:00"
               Height          =   255
               Left            =   1080
               TabIndex        =   19
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "已用时间："
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
         End
         Begin XtremeSuiteControls.ProgressBar Progress 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   8895
            _Version        =   983043
            _ExtentX        =   15690
            _ExtentY        =   450
            _StockProps     =   93
            BackColor       =   12640511
            Appearance      =   6
            BarColor        =   8454016
         End
         Begin VB.Label Lbl_Progress 
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            Height          =   255
            Left            =   9120
            TabIndex        =   14
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Lbl_Status 
            BackStyle       =   0  'Transparent
            Caption         =   "等待中......"
            Height          =   615
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   6735
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "扫描状态："
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Lbl_Object 
            BackStyle       =   0  'Transparent
            Caption         =   "该目录下的所有文件"
            Height          =   255
            Left            =   1320
            TabIndex        =   5
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "扫描对象："
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_Target 
            BackStyle       =   0  'Transparent
            Caption         =   "暂无"
            Height          =   255
            Left            =   1320
            TabIndex        =   3
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "扫描目标："
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "完成后点击开始处理，获取扫描报告。"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   6000
         Width           =   5055
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2280
      Top             =   4920
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StopScan As Boolean
Public TeScan As Boolean
Public StartTime As String

Private Sub Command1_Click()
On Error Resume Next
'删除……
DoEvents
If frmRow.FileList.ListItems.Count = 1 Then '只有当前这个扫描任务
frmProc.Show vbModal
frmFix.Show vbModal
End If

Dim StrLog As String
StrLog = "病毒防御助手 扫描报告" & vbCrLf & "---------------------------------" & vbCrLf & "开始时间：   " & StartTime & _
vbCrLf & "所用时间：   " & Lbl_Time.Caption & _
vbCrLf & "扫描目标：   " & Lbl_Target.Caption & _
vbCrLf & "扫描对象：   " & Lbl_Object.Caption & _
vbCrLf & "---------------------------------" & _
vbCrLf & "共扫描文件： " & filePathNum & " 个" & _
vbCrLf & "发现威胁：   " & ListVirus.ListItems.Count & " 个" & _
vbCrLf & "---------------------------------"
For virus = 1 To ListVirus.ListItems.Count
DoEvents
 Set itm = ListVirus.ListItems(virus)
 StrLog = StrLog & vbCrLf & "文件：    " & itm.Text & _
 vbCrLf & "威胁名：  " & itm.SubItems(1) & _
 vbCrLf & "威胁描述：" & itm.SubItems(2)
 If itm.Checked = True Then
 SetAttr itm.Text, vbNormal
 Kill itm.Text
 If Dir(itm.Text, vbNormal Or vbHidden Or vbSystem Or vbReadOnly) <> "" Then
 StrLog = StrLog & vbCrLf & "处理结果：" & "失败"
 Else
 StrLog = StrLog & vbCrLf & "处理结果：" & "成功"
 End If
 Else
 StrLog = StrLog & vbCrLf & "处理结果：" & "未处理"
 End If
Next
frmLog.Text1.Text = StrLog
frmLog.Show vbModal

selectover = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
For x = 1 To ListVirus.ListItems.Count
ListVirus.ListItems(x).Checked = True
Next
End Sub

Private Sub Command4_Click()
If StopScan = False Then
StopScan = True
Command4.Caption = "继续"
Command5.Enabled = True
Else
StopScan = False
Command4.Caption = "暂停"
Command5.Enabled = False
End If
End Sub

Private Sub Command5_Click()
TeScan = True
End Sub

Private Sub Command6_Click()
For x = 1 To ListVirus.ListItems.Count
ListVirus.ListItems(x).Checked = Not ListVirus.ListItems(x).Checked
Next
End Sub

Private Sub Form_Load()
CreatTray Me, "病毒防御助手-安全扫描", "病毒防御助手-安全扫描", "扫描待命", 4
selectover = True
StopScan = False
TeScan = False
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


With Me
'----------列表初始化----------
    .ListVirus.ListItems.Clear               '清空列表
    .ListVirus.ColumnHeaders.Clear           '清空列表头
    .ListVirus.View = lvwReport              '设置列表显示方式
    .ListVirus.GridLines = True              '显示网络线
    .ListVirus.LabelEdit = lvwManual         '禁止标签编辑
    .ListVirus.FullRowSelect = True          '选择整行
    .ListVirus.Checkboxes = True
    .ListVirus.ColumnHeaders.Add , , "威胁", .ListVirus.Width / 2 '给列表中添加列名
    .ListVirus.ColumnHeaders.Add , , "威胁类型", .ListVirus.Width / 2 '给列表中添加列名
    .ListVirus.ColumnHeaders.Add , , "描述", .ListVirus.Width / 2 '给列表中添加列名

End With
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.WindowState = vbNormal
Me.Hide
End If
If Me.Width <> 12075 Or Me.Height <> 8715 Then
Me.Width = 12075
Me.Height = 8715
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadTray
If selectover = False Then '如果扫描未完成，或者还没有完成处理
 If MsgBox("扫描正在进行或者威胁未处理，退出将丢失全部数据，您确定要退出吗?", vbOKCancel, "程序正在运行中......") <> vbOK Then
  Cancel = 1
 Else
  For Each i In Forms
   Unload i
  Next
  End
 End If
Else
  For Each i In Forms
   Unload i
  Next
  End
End If
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub


Private Sub ListVirus_ItemClick(ByVal item As MSComctlLib.ListItem)
If item.SubItems(2) = "" Then
Frame3.Visible = False
Else
Lbl_Tip.Caption = item.SubItems(2)
End If
End Sub

Private Sub ListVirus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If ListVirus.SelectedItem Is Nothing Then Exit Sub
If ListVirus.SelectedItem.Index <> 0 And ListVirus.ListItems.Count <> 0 Then
Frame3.Visible = True
Frame3.Top = y + ListVirus.Top + 10
Frame3.Left = x + ListVirus.Left
End If

End Sub

Private Sub Time_Timer()
Dim Time_Int As String
Time_Int = Format(CDate(Now) - (CDate(StartTime)), "hh:mm:ss")
Lbl_Time.Caption = Time_Int
Lbl_VirusNum.Caption = ListVirus.ListItems.Count
End Sub
