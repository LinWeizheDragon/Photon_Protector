Attribute VB_Name = "TextTip"

Private Function AddListItem(ByVal FirstText As String, ByVal SecondText As String, ByRef List As ListView)
Set itm = List.ListItems.Add(, , FirstText)
itm.SubItems(1) = SecondText
End Function


