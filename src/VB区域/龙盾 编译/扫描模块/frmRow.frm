VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRow 
   Caption         =   "光子防御网-扫描队列"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "frmRow.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4200
   ScaleWidth      =   5295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin MSComctlLib.ListView FileList 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ToFilePath As String
Public Function AddRow(ByVal Path As String, ByVal ScanMod As String)
Set item = FileList.ListItems.Add(, , Path)
item.SubItems(1) = "Wait"
item.SubItems(2) = ScanMod
DoCheck
End Function


Private Sub Command1_Click()
AddRow "M:\U盘恢复工具", "Adv"
End Sub

Private Sub Command2_Click()
frmProc.Show
End Sub

Private Sub Form_Load()
With Me
'----------列表初始化----------
    .FileList.ListItems.Clear               '清空列表
    .FileList.ColumnHeaders.Clear           '清空列表头
    .FileList.View = lvwReport              '设置列表显示方式
    .FileList.GridLines = True              '显示网络线
    .FileList.LabelEdit = lvwManual         '禁止标签编辑
    .FileList.FullRowSelect = True          '选择整行
    .FileList.Checkboxes = False
    .FileList.ColumnHeaders.Add , , "路径", .FileList.Width / 2 '给列表中添加列名
    .FileList.ColumnHeaders.Add , , "状态", .FileList.Width / 2 '给列表中添加列名
    .FileList.ColumnHeaders.Add , , "方式", .FileList.Width / 2 '给列表中添加列名
End With
End Sub

Public Function DoCheck()
With Me
If .FileList.ListItems.Count <> 0 Then '如果扫描队列中有项目
  If .FileList.ListItems(1).SubItems(1) = "Wait" And selectover = True Then '如果第一个项目的状态是等待
    Set item = .FileList.ListItems(1)
    DoScan item.Text, item.SubItems(2) '对第一个项目进行扫描
  Else '如果第一个项目的状态不是等待
  Exit Function '退出
  End If
Else '如果没有项目

Exit Function '退出
End If
End With
End Function

Public Function DoScan(ByVal Path As String, ScanMod As String)
'Path：文件路径
'ScanMod：Sim-只扫描根目录;Adv-扫描全部;Sin-单个文件
On Error Resume Next
frmMain.Show
selectover = False
frmMain.Command1.Enabled = False
frmMain.Lbl_Target.Caption = Path
Select Case ScanMod
 Case "Sin"
   frmMain.Lbl_Object.Caption = "单个文件"
 Case "Adv"
   frmMain.Lbl_Object.Caption = "该目录下的所有文件"
 Case "Sim"
   frmMain.Lbl_Object.Caption = "该目录根下的所有文件"
End Select
frmMain.Lbl_Status.Caption = "初始化......"
frmMain.Command4.Enabled = True
frmMain.Command5.Enabled = False
frmMain.StartTime = Now
frmMain.Time.Enabled = True
DoEvents
filePathNum = 0 '初始化计数
Erase FilePathGroup '初始化数组
If Dir(Path, vbDirectory Or vbHidden Or vbNormal Or vbSystem Or vbReadOnly) = "" Then
Debug.Print Dir(Path)
frmMain.Lbl_Object.Caption = ""
frmMain.Lbl_Status.Caption = "终止"
frmMain.Lbl_Target.Caption = "暂无"
GoTo Out:
End If
'恢复所有系统、隐藏文件夹

FileList.ListItems(1).SubItems(1) = "Check"
 Dim MyForm As New frmCheck
 Load MyForm
 With MyForm
 Dim FilePath As String
 Dim Result As String '接受结果
 Dim Num As Long
 Dim total As Long
    Dim StrSplit As String
 '以上声明部分
Select Case ScanMod
Case "Sim"
 Showfilelist Path, .List1
 If selectover = True Then
    frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   GoTo Out:
 End If
 total = filePathNum + 1
 For i = 0 To filePathNum
 FilePath = Path & "\" & FilePathGroup(i)
 If frmMain.StopScan = True Then
   Do Until frmMain.StopScan = False Or frmMain.TeScan = True
   SuperSleep 1
   Loop
   If frmMain.StopScan = False Then
   End If
   If frmMain.TeScan = True Then
   frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   GoTo Out:
   End If
End If
  Result = ProcessScan(FilePath)
  Num = i + 1
  frmMain.Progress.Value = Round((Num / total) * 100, 0)
  frmMain.Lbl_Status.Caption = "正在扫描......" & FilePath '"已完成  " & Str(frmMain.Progress.Value) & "%"
  frmMain.Lbl_Progress.Caption = Str(frmMain.Progress.Value) & "%"
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
    '是病毒
    Set itm = frmMain.ListVirus.ListItems.Add(, , FilePath)
    itm.SubItems(1) = Split(Result, "|")(0)
    itm.SubItems(2) = Split(Result, "|")(1)
    itm.Checked = True
  Else
    If UBound(Split(FilePath, "\")) = 1 Then
     StrSplit = Split(FilePath, ".")(UBound(Split(FilePath, ".")))
   If StrSplit = FilePath Then GoTo skip:
   Debug.Print Left(FilePath, Len(FilePath) - Len(StrSplit) - 1)
   If Dir(Left(FilePath, Len(FilePath) - Len(StrSplit) - 1), vbDirectory) <> "" And Right(FilePath, 4) = ".exe" Then
   '同名文件夹
   Set itm = frmMain.ListVirus.ListItems.Add(, , FilePath)
    itm.SubItems(1) = "文件夹同名文件"
    itm.SubItems(2) = "此文件与""" & Left(FilePath, Len(FilePath) - Len(StrSplit) - 1) & """文件夹同名，如果您不认识此文件，请删除！"
    itm.Checked = True
   End If
   End If
  End If
skip:
 Next
Case "Adv"
  sousuofile Path, .List1
  If selectover = True Then
     frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   GoTo Out:
  End If
 total = filePathNum + 1
 For i = 0 To filePathNum
 FilePath = FilePathGroup(i)
If frmMain.StopScan = True Then
   Do Until frmMain.StopScan = False Or frmMain.TeScan = True
   SuperSleep 1
   Loop
   If frmMain.StopScan = False Then
   End If
   If frmMain.TeScan = True Then
   frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   GoTo Out:
   End If
End If
  Result = ProcessScan(FilePath)
  Debug.Print FilePath & "|" & Result
  Num = i + 1
  frmMain.Progress.Value = Round((Num / total) * 100, 0)
  frmMain.Lbl_Status.Caption = "正在扫描......" & FilePath '"已完成  " & Str(frmMain.Progress.Value) & "%"
  frmMain.Lbl_Progress.Caption = Str(frmMain.Progress.Value) & "%"
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
    '是病毒
    Set itm = frmMain.ListVirus.ListItems.Add(, , FilePath)
    itm.SubItems(1) = Split(Result, "|")(0)
    itm.SubItems(2) = Split(Result, "|")(1)
    itm.Checked = True
  Else
   If UBound(Split(FilePath, "\")) = 1 Then
     StrSplit = Split(FilePath, ".")(UBound(Split(FilePath, ".")))
     If StrSplit = FilePath Then GoTo skip2:
   Debug.Print Left(FilePath, Len(FilePath) - Len(StrSplit) - 1)
   If Dir(Left(FilePath, Len(FilePath) - Len(StrSplit) - 1), vbDirectory) <> "" And Right(FilePath, 4) = ".exe" Then
   '同名文件夹
   Set itm = frmMain.ListVirus.ListItems.Add(, , FilePath)
    itm.SubItems(1) = "文件夹同名文件"
    itm.SubItems(2) = "此文件与""" & Left(FilePath, Len(FilePath) - Len(StrSplit) - 1) & """文件夹同名，如果您不认识此文件，请删除！"
    itm.Checked = True
   End If
   End If
  End If
skip2:
 Next
Case "Sin" '单文件
  Result = ProcessScan(Path)
  frmMain.Progress.Value = 100
  total = 1
  frmMain.Lbl_Status.Caption = "已完成  " & Str(frmMain.Progress.Value) & "%"
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
    Set itm = frmMain.ListVirus.ListItems.Add(, , Path)
    itm.SubItems(1) = Split(Result, "|")(0)
    itm.SubItems(2) = Split(Result, "|")(1)
    itm.Checked = True
  Else
  If UBound(Split(FilePath, "\")) = 1 Then
   StrSplit = Split(FilePath, ".")(UBound(Split(FilePath, ".")))
   If StrSplit = FilePath Then GoTo skip3:
   Debug.Print Left(FilePath, Len(FilePath) - Len(StrSplit) - 1)
   If Dir(Left(FilePath, Len(FilePath) - Len(StrSplit) - 1), vbDirectory) <> "" And Right(FilePath, 4) = ".exe" Then
   '同名文件夹
   Set itm = frmMain.ListVirus.ListItems.Add(, , FilePath)
    itm.SubItems(1) = "文件夹同名文件"
    itm.SubItems(2) = "此文件与""" & Left(FilePath, Len(FilePath) - Len(StrSplit) - 1) & """文件夹同名，如果您不认识此文件，请删除！"
    itm.Checked = True
   End If
  End If
  End If
skip3:
End Select
frmMain.Show
frmMain.Lbl_Status.Caption = "共扫描： " & total & " 个文件，发现威胁： " & frmMain.ListVirus.ListItems.Count & " 个"
frmMain.Command1.Enabled = True
frmMain.Time.Enabled = False
frmMain.Command4.Enabled = False
Do Until selectover = True
 SuperSleep 3
Loop

Out:
FileList.ListItems.Remove 1
frmMain.Time.Enabled = False
End With
DoClean
End Function



Public Function DoClean()
'扫地工作，吼吼
frmMain.Lbl_Object.Caption = ""
frmMain.Lbl_Status.Caption = "等待中......"
frmMain.Lbl_Target.Caption = "暂无"
frmMain.Progress.Value = 0
frmMain.ListVirus.ListItems.Clear
frmMain.Command1.Enabled = False
frmMain.StopScan = False
frmMain.TeScan = False
frmMain.Command4.Enabled = False
frmMain.Command4.Caption = "暂停"
frmMain.Command5.Enabled = False
frmMain.Lbl_Time.Caption = "00:00:00"
frmMain.Lbl_Progress.Caption = "0%"
frmMain.Lbl_VirusNum.Caption = 0
DoCheck
selectover = True
End Function
