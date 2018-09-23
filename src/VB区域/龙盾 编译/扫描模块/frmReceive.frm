VERSION 5.00
Begin VB.Form frmRec 
   Caption         =   "DragonShieldRec"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "txtReceive1"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Text1.Text = InputBox("内容", "hehe")
End Sub

Private Sub Text1_Change()
'当Text1的内容被更改则说明收到传来的数据
If Text1.Text <> "txtReceive1" Then AddRow Text1.Text
Text1.Text = "txtReceive1" 'Text1复位，等待下一次接收数据
End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub AddRow(ByVal RecText)
Dim isVirus As Boolean
If UBound(Split(RecText, "|")) = 3 Then '说明带有病毒描述等等
  isVirus = True
ElseIf UBound(Split(RecText, "|")) = 2 Then
  isVirus = False
End If
Set Item = frmRow.WaitList.FindItem(Split(RecText, "|")(0))
If isVirus = True Then
Item.SubItems(1) = Split(RecText, "|")(1) & Split(RecText, "|")(2)
ElseIf isVirus = False Then
Item.SubItems(1) = Split(RecText, "|")(1)
End If

End Sub
Private Function AddListItem(ByVal FirstText As String, ByVal SecondText As String, ByRef List As ListView)
Set itm = List.ListItems.Add(, , FirstText)
itm.SubItems(1) = SecondText
End Function


