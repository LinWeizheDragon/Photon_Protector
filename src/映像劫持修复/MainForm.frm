VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form MainForm 
   Caption         =   "镜像劫持修复"
   ClientHeight    =   4155
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   3480
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3480
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "立即修复"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "共有 0 个映像劫持项目。"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   0
      Top             =   4200
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu YW_Regedit_File 
      Caption         =   "文件(&F)"
      Begin VB.Menu YW_Regedit_File_New 
         Caption         =   "新建(&N)"
      End
      Begin VB.Menu YW_Regedit_File_Refresh 
         Caption         =   "刷新(&E)"
      End
      Begin VB.Menu YW_Regedit_File_OneOk 
         Caption         =   "一键修复"
      End
      Begin VB.Menu YW_Regedit_File_Delimiter1 
         Caption         =   "-"
      End
      Begin VB.Menu YW_Regedit_File_Close 
         Caption         =   "关闭程序(&C)"
      End
   End
   Begin VB.Menu YW_Regedit_RightMenu 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu YW_Regedit_RightMenu_Point 
         Caption         =   "指向"
      End
      Begin VB.Menu YW_Regedit_RightMenu_Delete 
         Caption         =   "删除"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y
Private Sub Command1_Click()
Call YW_Regedit_File_OneOk_Click
End Sub

Private Sub Form_Load()
x = Me.Width
y = Me.Height
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hWnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

'打开注册表
Dim YW_Regedit_Return As Long
YW_Regedit_Return = RegOpenKey(HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options", _
YW_Regedit_Hkey) '打开注册表
If YW_Regedit_Return <> 0 Then
MsgBox ("打开注册表失败!")
End
End If

Select Case Command
Case ""
Case "/Quiet"
Me.Hide
Call YW_Regedit_Scanning
Call YW_Regedit_File_OneOk_Click
End
Case "/quiet"
Me.Hide
Call YW_Regedit_Scanning
Call YW_Regedit_File_OneOk_Click
End
End Select
Call YW_Regedit_Scanning
If List1.ListCount = 0 Then
MsgBox "没有镜像劫持的项", vbInformation, "病毒防御助手 映像劫持修复"
End If
Label2.Caption = "共有 " & List1.ListCount & " 个映像劫持项目"

End Sub

Private Sub Form_Resize()
If Me.Height <> y Or Me.Width <> x Then
Me.Height = y
Me.Width = x
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim YW_Regedit_Return As Long
YW_Regedit_Return = RegCloseKey(YW_Regedit_Hkey)
If YW_Regedit_Return <> 0 Then
MsgBox ("关闭注册表失败!")
End If
End Sub

Private Sub List1_Click()
YW_Regedit_Focus2 = List1.List(List1.ListIndex)
PopupMenu YW_Regedit_RightMenu
End Sub

Private Sub YW_Regedit_File_Close_Click()
End
End Sub

Private Sub YW_Regedit_File_OneOk_Click()
On Error GoTo YW_Regedit_File_OneOk_Click_Error
Dim YW_Regedit_File_OneOk_Click1 As Long
For YW_Regedit_File_OneOk_Click1 = 1 To List1.ListCount
YW_Regedit_RightMenu_Delete_Click1 = RegDeleteKey( _
YW_Regedit_Hkey, List1.List(YW_Regedit_File_OneOk_Click1 - 1))
If YW_Regedit_RightMenu_Delete_Click1 <> 0 Then
MsgBox ("删除失败")
Exit For
End If
Next
Call YW_Regedit_Scanning
Exit Sub
YW_Regedit_File_OneOk_Click_Error:
MsgBox ("无法删除键值!")
Call YW_Regedit_Scanning
End Sub

Private Sub YW_Regedit_File_Refresh_Click()
On Error GoTo YW_Regedit_File_Refresh_Click_Error
Call YW_Regedit_Scanning
Exit Sub
YW_Regedit_File_Refresh_Click_Error:
MsgBox ("无法调用扫描模块!")
End Sub

Private Sub YW_Regedit_RightMenu_Delete_Click()
On Error GoTo YW_Regedit_RightMenu_Delete_Click_Error
Dim YW_Regedit_RightMenu_Delete_Click1 As Long
YW_Regedit_RightMenu_Delete_Click1 = MsgBox("确认删除?", 33)
If YW_Regedit_RightMenu_Delete_Click1 = 1 Then
YW_Regedit_RightMenu_Delete_Click1 = RegDeleteKey( _
YW_Regedit_Hkey, YW_Regedit_Focus2)
If YW_Regedit_RightMenu_Delete_Click1 = 0 Then
MsgBox ("删除成功")
Else
MsgBox ("删除失败")
End If
Call YW_Regedit_Scanning
End If
Exit Sub
YW_Regedit_RightMenu_Delete_Click_Error:
MsgBox ("无法删除键值!")
Call YW_Regedit_Scanning
End Sub

Private Sub YW_Regedit_File_New_Click()
On Error GoTo YW_Regedit_File_New_Error
Dim YW_Regedit_RightMenu_New_Click1 As Long
Dim YW_Regedit_RightMenu_New_Click3 As String
Dim YW_Regedit_RightMenu_New_Click4 As String
Dim YW_Regedit_RightMenu_New_Click5 As Long
YW_Regedit_RightMenu_New_Click3 = InputBox("请输入名称:", "镜像劫持修复")
YW_Regedit_RightMenu_New_Click4 = InputBox("请输入指向地址:", "镜像劫持修复")
YW_Regedit_RightMenu_New_Click1 = RegCreateKey(YW_Regedit_Hkey, _
YW_Regedit_RightMenu_New_Click3, YW_Regedit_RightMenu_New_Click5)
If YW_Regedit_RightMenu_New_Click1 = 0 Then
YW_Regedit_RightMenu_New_Click1 = RegSetValueEx(YW_Regedit_RightMenu_New_Click5, _
"Debugger", 0, REG_SZ, ByVal YW_Regedit_RightMenu_New_Click4, _
Len(YW_Regedit_RightMenu_New_Click4))
Else
MsgBox ("写入注册表时发生错误!")
End If
If YW_Regedit_RightMenu_New_Click1 = 0 Then
MsgBox ("写入注册表成功")
Else
MsgBox ("写入注册表时发生错误!")
End If

Call YW_Regedit_Scanning
Exit Sub
YW_Regedit_File_New_Error:
MsgBox ("无法新建键值!")
Call YW_Regedit_Scanning
End Sub

Private Sub YW_Regedit_RightMenu_Point_Click()
On Error GoTo YW_Regedit_RightMenu_Point_Click_Error
Dim YW_Regedit_RightMenu_Point_Click1 As Long
Dim YW_Regedit_RightMenu_Point_Click2 As Long
Dim YW_Regedit_RightMenu_Point_Click3 As Long
Dim YW_Regedit_RightMenu_Point_Click4 As String
YW_Regedit_RightMenu_Point_Click1 = RegOpenKey(HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & _
YW_Regedit_Focus2, YW_Regedit_RightMenu_Point_Click2)
If YW_Regedit_RightMenu_Point_Click1 = 0 Then
YW_Regedit_RightMenu_Point_Click4 = Space(REG_SIZE)
YW_Regedit_RightMenu_Point_Click1 = RegQueryValueEx(YW_Regedit_RightMenu_Point_Click2, _
"Debugger", 0, YW_Regedit_RightMenu_Point_Click3, _
ByVal YW_Regedit_RightMenu_Point_Click4, REG_SIZE)
If YW_Regedit_RightMenu_Point_Click1 = 0 Then
MsgBox ("指向->" & YW_Regedit_RightMenu_Point_Click4)
Else
MsgBox ("读取注册表时发生错误!")
End If
Else
MsgBox ("打开注册表时发生错误!")
Exit Sub
End If
Call YW_Regedit_Scanning
Exit Sub
YW_Regedit_RightMenu_Point_Click_Error:
MsgBox ("无法读取键值!")
Call YW_Regedit_Scanning
End Sub
