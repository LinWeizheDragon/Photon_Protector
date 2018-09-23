VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUSBOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "U盘实时防护"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4665
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4471
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Max             =   1
   End
   Begin VB.Label Label2 
      Caption         =   "普通模式：只扫描U盘下根目录文件"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "严格搜模式：扫描U盘所有文件，大于8GB自动改为普通模式"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmUSBOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Slider.Value = 0 Then
WriteString "USBRTA", "CheckMod", "Adv", App.Path & "\Set.ini"
Else
WriteString "USBRTA", "CheckMod", "Sim", App.Path & "\Set.ini"
End If
Unload Me
End Sub
