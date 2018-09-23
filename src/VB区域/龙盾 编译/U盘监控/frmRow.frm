VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRow 
   Caption         =   "龙盾-扫描队列"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form2"
   ScaleHeight     =   4200
   ScaleWidth      =   5295
   StartUpPosition =   3  '窗口缺省
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
Public Function AddRow(ByVal Path As String, ByVal ScanMod As String)
Set item = FileList.ListItems.Add(, , Path)
item.SubItems(1) = "Wait"
item.SubItems(2) = ScanMod
DoCheck
End Function


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
  If .FileList.ListItems(1).SubItems(1) = "Wait" Then '如果第一个项目的状态是等待
    Set item = .FileList.ListItems(1)
    DoScan item.Text, item.SubItems(2) '对第一个项目进行扫描
  Else '如果第一个项目的状态不是等待
  Exit Function '退出
  End If
Else '如果没有项目
If frmUSB.Visible = True Then
Unload frmUSB
End If
Exit Function '退出
End If
End With
End Function

Public Function DoScan(ByVal Path As String, ScanMod As String)
'Path：文件路径
On Error Resume Next
'ScanMod：Sim-只扫描根目录;Adv-扫描全部;Sin-单个文件
DoEvents
If Dir(Path) = "" Then Exit Function '没有找到此文件或路径
'恢复所有系统、隐藏文件夹
Dim XFSO As New FileSystemObject
Set fso = XFSO.GetFolder(Path)
For Each Folder In fso.SubFolders
 Call SetAttr(Folder, vbNormal)
Next

FileList.ListItems(1).SubItems(1) = "Check"
 Dim MyForm As New frmCheck
 Load MyForm
 With MyForm
 Dim FilePath As String
 Dim Result As String '接受结果
 Dim VirusDes As String '用于组合传递
 Dim singlemod As Boolean
    Dim StrSplit As String
 '以上声明部分
Select Case ScanMod
Case "Sim"
 Showfilelist Path, .List1
 For i = 0 To .List1.ListCount - 1
  .List1.ListIndex = i
  FilePath = Path & "\" & .List1.Text
  Result = ProcessScan(FilePath)
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
    If VirusDes = "" Then '如果还没有任何一个病毒
      VirusDes = FilePath & "|" & Result '赋值
      singlemod = True '打开单个病毒模式
    Else
      VirusDes = VirusDes & "||" & FilePath & "|" & Result '附加
      singlemod = False
    End If
  Else
   '如果没有查到这是个病毒的话
   StrSplit = Split(FilePath, ".")(UBound(Split(FilePath, ".")))
   Debug.Print Left(FilePath, Len(FilePath) - Len(StrSplit) - 1)
   If Dir(Left(FilePath, Len(FilePath) - Len(StrSplit) - 1), vbDirectory) <> "" And Right(FilePath, 4) = ".exe" Then
   '同名文件夹
   Result = "文件夹同名文件|此文件与""" & Left(FilePath, Len(FilePath) - Len(StrSplit) - 1) & """文件夹同名，如果您不认识此程序，请删除！"
    If VirusDes = "" Then '如果还没有任何一个病毒
      VirusDes = FilePath & "|" & Result '赋值
      singlemod = True '打开单个病毒模式
    Else
      VirusDes = VirusDes & "||" & FilePath & "|" & Result '附加
      singlemod = False
    End If
   End If
  End If
 Next
Case "Adv"
  sousuofile Path, .List1
 For i = 0 To .List1.ListCount - 1
  .List1.ListIndex = i
  FilePath = .List1.Text
  Result = ProcessScan(FilePath)
  Debug.Print FilePath & ":" & Result
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
    If VirusDes = "" Then '如果还没有任何一个病毒
      VirusDes = FilePath & "|" & Result '赋值
      singlemod = True '打开单个病毒模式
    Else
      VirusDes = VirusDes & "||" & FilePath & "|" & Result '附加
      singlemod = False
    End If
  Else
   '如果没有查到这是个病毒的话
   If UBound(Split(FilePath, "\")) = 1 Then
   StrSplit = Split(FilePath, ".")(UBound(Split(FilePath, ".")))
   Debug.Print Left(FilePath, Len(FilePath) - Len(StrSplit) - 1)
   If Dir(Left(FilePath, Len(FilePath) - Len(StrSplit) - 1), vbDirectory) <> "" And Right(FilePath, 4) = ".exe" Then
   '同名文件夹
   Result = "文件夹同名文件|此文件与""" & Left(FilePath, Len(FilePath) - Len(StrSplit) - 1) & """文件夹同名，如果您不认识此程序，请删除！"
    If VirusDes = "" Then '如果还没有任何一个病毒
      VirusDes = FilePath & "|" & Result '赋值
      singlemod = True '打开单个病毒模式
    Else
      VirusDes = VirusDes & "||" & FilePath & "|" & Result '附加
      singlemod = False
    End If
   End If
   End If
  End If
 Next
Case "Sin" '单文件
  Result = ProcessScan(Path)
  If Result <> "SAFE" And Result <> "Error" Then '如果不出错也不是安全
      VirusDes = Path & "|" & Result '赋值
  End If
End Select
If VirusDes <> "" Then ' 如果有病毒
  ShowTextTip VirusDes, singlemod '传递！
End If
FileList.ListItems.Remove 1
End With
DoCheck
End Function
