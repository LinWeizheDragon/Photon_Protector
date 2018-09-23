VERSION 5.00
Begin VB.Form frmRec 
   Caption         =   "消息接受窗体"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3630
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   960
   End
End
Attribute VB_Name = "frmRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    strShare.Create "TestShareString"
    Text1 = strShare
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If strShare = Text1 Then

Exit Sub
End If
Text1 = strShare
DoCommand strShare
End Sub

Private Sub DoCommand(ByVal Text As String)

End Sub
