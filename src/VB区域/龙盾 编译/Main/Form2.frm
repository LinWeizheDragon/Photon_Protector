VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提示"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5025
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "确定"
      DownPicture     =   "Form2.frx":324A
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "暂时不提供杀毒服务！"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
