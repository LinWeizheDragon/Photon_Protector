VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMain 
   Caption         =   "病毒防御助手-U盘实时防护"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9870
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   9870
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Fuck 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8280
      Top             =   1080
   End
   Begin MSComctlLib.ListView ListDrv 
      Height          =   5055
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "创建"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2940
      Left            =   5160
      TabIndex        =   3
      Top             =   -360
      Width           =   1935
   End
   Begin VB.ListBox List 
      Height          =   1680
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2280
      Top             =   4920
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuFresh 
         Caption         =   "刷新"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "关闭保护"
      End
      Begin VB.Menu mnuTop 
         Caption         =   "置顶/取消置顶"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "隐藏悬浮窗"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TopMost As Boolean
Dim ShowBtn As Boolean
Private Sub Command3_Click()
'Form1.USBList1.SetText "", "F:()" & vbCrLf & vbCrLf & "可用空间。。。", "99"
End Sub

Public Function ReReadDrv()
On Error Resume Next
Form1.Height = 720
ListDrv.ListItems.Clear
Dim MyFSO As New FileSystemObject
For Each i In MyFSO.Drives
If i = "A:" Then GoTo GNext:
If i = "A:\" Then GoTo GNext:

If MyFSO.GetDrive(i).DriveType = Removable Then
Set itm = ListDrv.ListItems.Add(, , i)
itm.SubItems(1) = MyFSO.GetDrive(i).VolumeName
itm.SubItems(2) = "USB"
End If
GNext:
Next
Dim HdDrv
Dim AllDrv
Form1.Label1.Caption = "病毒防御助手 刷新列表中"
DoEvents
AllDrv = findUsbHardDisk
For Y = 0 To UBound(Split(AllDrv, "|")) - 1
HdDrv = Split(AllDrv, "|")(Y)
Set itm = ListDrv.ListItems.Add(, , HdDrv)
itm.SubItems(1) = MyFSO.GetDrive(HdDrv).VolumeName
itm.SubItems(2) = "HardDisk"
Next


With Form1
Dim a
Dim Fist, Second, Path
Dim Num As Integer
Num = ListDrv.ListItems.Count
If Num = 0 Then
Form1.Hide
Else
Form1.Show
End If
a = 1
For a = 1 To Num
If a > 10 Then
Form1.Label1.Caption = "已达上限10个，其余不显示"
Exit Function
End If
If ListDrv.ListItems(a).text <> "" Then
Form1.USB(a - 1).Caption = ListDrv.ListItems(a).text
If ListDrv.ListItems(a).SubItems(2) = "USB" Then
Form1.USB(a - 1).SetPicture Form1.Picture2
Else
Form1.USB(a - 1).SetPicture Form1.Picture1
End If
If a <= Num Then '条数小于等于总条数
Form1.Height = Form1.Height + 960
End If

End If
Next

End With
Form1.Label1.Caption = "病毒防御助手 U盘监控"
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

Private Sub Command4_Click()
ReReadDrv
End Sub

Private Sub Command5_Click()
'Form1.USBList1.SetValue 0
'Form1.USBList1.SetValue 1
End Sub

Private Sub Form_Load()
On Error Resume Next
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hWnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

HookForm Me
Form1.Show
With frmMain
'----------列表初始化----------
    .ListDrv.ListItems.Clear               '清空列表
    .ListDrv.ColumnHeaders.Clear           '清空列表头
    .ListDrv.View = lvwReport              '设置列表显示方式
    .ListDrv.GridLines = True              '显示网络线
    .ListDrv.LabelEdit = lvwManual         '禁止标签编辑
    .ListDrv.FullRowSelect = True          '选择整行
    .ListDrv.Checkboxes = False
    .ListDrv.ColumnHeaders.Add , , "路径", 1000  '给列表中添加列名
    .ListDrv.ColumnHeaders.Add , , "名称", .ListDrv.Width '给列表中添加列名
    .ListDrv.ColumnHeaders.Add , , "类型", 0# '给列表中添加列名

End With
If ReadString("USBRTA", "TopMost", App.Path & "\Set.ini") = "1" Then
SetWindowPos Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
TopMost = True
Else
TopMost = False
SetWindowPos Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flag
End If
ReReadDrv
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnHookForm Me
For Each i In Forms
Unload i
Next
End
End Sub

Private Sub Fuck_Timer()
MyFrame.Width = 0
End Sub

Private Sub jcbutton1_Click()
ShowBtn = True
BtnFrame.Visible = True
MyFrame.Top = jcbutton1.Top
MyFrame.Width = 0
SuperSleep 0.02
MyFrame.Width = 500
SuperSleep 0.02
MyFrame.Width = 1500
SuperSleep 0.02
MyFrame.Width = 2500
SuperSleep 0.02
MyFrame.Width = 3135
End Sub

Private Sub jcbutton1_MouseEnter()
ShowBtn = False '不显示按钮
BtnFrame.Visible = False
MyFrame.Top = jcbutton1.Top
MyFrame.Width = 0
SuperSleep 0.02
MyFrame.Width = 500
SuperSleep 0.02
MyFrame.Width = 1500
SuperSleep 0.02
MyFrame.Width = 2500
SuperSleep 0.02
MyFrame.Width = 3135

End Sub

Private Sub jcbutton1_MouseLeave()
If ShowBtn = False Then '如果不是显示按钮
MyFrame.Width = 0
End If

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFresh_Click()
ReReadDrv
End Sub

Private Sub mnuHide_Click()
Form1.Hide
End Sub

Private Sub mnuTop_Click()
If TopMost = False Then
SetWindowPos Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
TopMost = True
Call WriteString("USBRTA", "TopMost", "1", App.Path & "\Set.ini")
Else
TopMost = False
SetWindowPos Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flag
Call WriteString("USBRTA", "TopMost", "0", App.Path & "\Set.ini")
End If
End Sub


Private Sub MyFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fuck.Enabled = False
End Sub
