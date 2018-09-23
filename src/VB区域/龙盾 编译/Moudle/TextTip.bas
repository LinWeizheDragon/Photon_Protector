Attribute VB_Name = "TextTip"




Public Function ShowTextTip(ByVal MainText As String, ByVal DesTip As String, ByVal FilePath As String)

Dim MyTip As New frmTip
With MyTip


MyTip.Tip.Caption = MainText
'Dim Astring() As String
'Dim Bstring() As String
'For I = 0 To UBound(Split(DesTip, "|"))
'Astring(I) = Split(DesTip, "|")(I)
'Next
'For X = 0 To UBound(Astring)
'AddListItem Split(Astring(X), ":")(0), Split(Astring(X), ":")(1), MyTip.ListView1
'Next
MyTip.Text1.Text = Replace(DesTip, "|", vbCrLf)
MyTip.Option2.Value = True
If FilePath <> "" Then
.PicIcon.Picture = GetIconFromFile(FilePath, 0, True)
Else
.PicIcon.Visible = False
.PicIcon.AutoRedraw = True
End If
.Show
Do Until .Visible = False
SuperSleep 1
Loop

ShowTextTip = MyTip.ChooseNum
End With
End Function

Private Function AddListItem(ByVal FirstText As String, ByVal SecondText As String, ByRef List As ListView)
Set itm = List.ListItems.Add(, , FirstText)
itm.SubItems(1) = SecondText
End Function


