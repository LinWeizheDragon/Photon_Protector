VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9060
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView ListFile 
      Height          =   4095
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   120
      Top             =   0
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hWnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics


'----------列表初始化----------
    ListFile.ListItems.Clear               '清空列表
    ListFile.ColumnHeaders.Clear           '清空列表头
    ListFile.View = lvwIcon              '设置列表显示方式
   'ListFile.GridLines = True              '显示网络线
   ' ListFile.LabelEdit = lvwManual         '禁止标签编辑
    ListFile.FullRowSelect = True          '选择整行
    ListFile.Checkboxes = False
  '  ListFile.ColumnHeaders.Add , , "特征码", ListFile.Width / 2 '给列表中添加列名
   ' ListFile.ColumnHeaders.Add , , "名称", ListFile.Width / 2 '给列表中添加列名
   ' ListFile.ColumnHeaders.Add , , "描述", ListFile.Width / 2 '给列表中添加列名

End Sub
