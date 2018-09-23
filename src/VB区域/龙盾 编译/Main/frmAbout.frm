VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin PhotonProtect.jcbutton jcbutton1 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632064
      Caption         =   "确定"
   End
   Begin PhotonProtect.jcbutton jcbutton2 
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "X"
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
Unload Me
End Sub

Private Sub jcbutton2_Click()
Unload Me
End Sub
