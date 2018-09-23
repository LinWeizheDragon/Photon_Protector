VERSION 5.00
Begin VB.Form frmChoose 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "龙盾-实时防护"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6075
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "打开高级实时防护（Ring0）"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开初级实时防护（WMI）"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    龙盾实时防护带有两种方式："
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmWatch.Show
Unload Me
End Sub

Private Sub Command2_Click()
If frmMain.exitproc("ProcessRTA.exe") = True Then
  strShare = "ProcessRTA"
  SuperSleep 1
 strShare = "ProcessRTA.Show"
Else
  Shell App.Path & "\ProcessRTA.exe"
End If
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = "龙盾实时防护拥有两种方式：" & vbCrLf _
& "1.WMI监视 优点：稳定性高 缺点：只能在同一时间拦截同一个进程，有可能放任木马运行，且有可能占用10%—20%的系统资源" & vbCrLf _
& "2.Ring0 Hook监视 优点：写于系统底层，防护性能高，占用资源少 缺点：不稳定，在打开单个文件时可能出现延迟、停滞等现象。" & vbCrLf _
& "请选择一项进行启动（WMI监视绑定于主程序，Ring0 Hook监视绑定于另外的进程，请尽量不要将两个防御都打开，防止出现系统错误）"

End Sub
