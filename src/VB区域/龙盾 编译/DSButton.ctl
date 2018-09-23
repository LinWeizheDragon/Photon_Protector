VERSION 5.00
Begin VB.UserControl DSButton 
   BackColor       =   &H00000000&
   BackStyle       =   0  '透明
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ClipBehavior    =   0  '无
   Picture         =   "DSButton.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   1200
   Begin VB.Image Off2 
      Height          =   330
      Left            =   0
      MouseIcon       =   "DSButton.ctx":1542
      MousePointer    =   99  'Custom
      Picture         =   "DSButton.ctx":1694
      Top             =   0
      Width           =   1155
   End
   Begin VB.Image Off1 
      Height          =   330
      Left            =   0
      MouseIcon       =   "DSButton.ctx":2BD6
      MousePointer    =   99  'Custom
      Picture         =   "DSButton.ctx":2D28
      Top             =   0
      Width           =   1155
   End
   Begin VB.Image On2 
      Height          =   330
      Left            =   0
      MouseIcon       =   "DSButton.ctx":426A
      MousePointer    =   99  'Custom
      Picture         =   "DSButton.ctx":43BC
      Top             =   0
      Width           =   1155
   End
   Begin VB.Image On1 
      Height          =   330
      Left            =   0
      MouseIcon       =   "DSButton.ctx":58FE
      MousePointer    =   99  'Custom
      Picture         =   "DSButton.ctx":5A50
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "DSButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mode As Integer



Private Sub Off1_Click()
frmControl.CallMode mode
End Sub

Private Sub Off2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ShowPic Off1

End Sub

Private Sub On1_Click()
frmControl.CallMode mode
End Sub

Private Sub On2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ShowPic On1
End Sub

Private Sub UserControl_Initialize()
ShowPic Off2
End Sub

Private Sub ShowPic(ByRef Img As Image)
On1.Visible = False
On2.Visible = False
Off1.Visible = False
Off2.Visible = False
Img.Visible = True
End Sub
Public Sub Reset()
If On1.Visible = True Then ShowPic On2
If Off1.Visible = True Then ShowPic Off2
End Sub

Public Sub SetStatus(ByVal OnOrOff As Boolean)
If OnOrOff Then
ShowPic On2
Else
ShowPic Off2
End If

End Sub

Public Function GetStatus() As Boolean
If On2.Visible = True Or On1.Visible = True Then '任意一个是显示的
GetStatus = True
ElseIf Off2.Visible = True Or Off1.Visible = True Then
GetStatus = False
End If

End Function

Public Function SetClickMode(ByVal ModeIndex) As Boolean
mode = ModeIndex
End Function
