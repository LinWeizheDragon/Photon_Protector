VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "龙盾-移动盘隐藏文件夹修复工具"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   5325
   ScaleWidth      =   7020
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "修复文件夹"
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "修复"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2760
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DoEvents
Dim Path As String
For x = 0 To List1.ListCount - 1
Path = List1.List(x)
Call SetAttr(Path, vbNormal)
FileCopy App.Path & "\Folder\desktop.ini", Path & "\desktop.ini"
FileCopy App.Path & "\Folder\文件夹安全验证图标.ico", Path & "\文件夹安全验证图标.ico"
Call SetAttr(Path, vbSystem)
Next
MsgBox "修复成功", vbInformation, "完成了"
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Drive
Dim MyFSO As New FileSystemObject
For Each I In MyFSO.Drives
If MyFSO.GetDrive(I).DriveType = Removable Then
ShowFolderList (I)
End If
Next
End Sub
Public Sub ShowFolderList(folderspec)
     Dim fs, f, f1, s, sf
     Dim hs, h, h1, hf
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set sf = f.SubFolders
     For Each f1 In sf
     Form1.List1.AddItem folderspec & f1.Name
     Next
End Sub
