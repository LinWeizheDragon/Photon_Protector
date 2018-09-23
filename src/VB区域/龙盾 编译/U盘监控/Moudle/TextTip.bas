Attribute VB_Name = "TextTip"
Public Function ShowTextTip(ByVal VirusDes As String, ByVal singlemod As Boolean) As Boolean
'传入：文件|病毒名|病毒描述||文件|病毒名|病毒描述
Dim MyTip As New frmTip
With MyTip
'----------列表初始化----------
    .ListView.ListItems.Clear               '清空列表
    .ListView.ColumnHeaders.Clear           '清空列表头
    .ListView.View = lvwReport              '设置列表显示方式
    .ListView.GridLines = True              '显示网络线
    .ListView.LabelEdit = lvwManual         '禁止标签编辑
    .ListView.FullRowSelect = True          '选择整行
    .ListView.Checkboxes = True
    .ListView.ColumnHeaders.Add , , "文件", .ListView.Width / 2 '给列表中添加列名
    .ListView.ColumnHeaders.Add , , "病毒名", .ListView.Width / 2 - 100 '给列表中添加列名
    .ListView.ColumnHeaders.Add , , "病毒描述", 0 '给列表中添加列名

If Not singlemod = True Then
ReDim Astring(UBound(Split(VirusDes, "||"))) As String
'先把内容分割成一个文件一个字符串
For i = 0 To UBound(Split(VirusDes, "||"))
Astring(i) = Split(VirusDes, "||")(i)
Next
'再把内容分割成文件|病毒名|病毒描述各一个，写入Listview
For x = 0 To UBound(Astring)
AddListItem Split(Astring(x), "|")(0), Split(Astring(x), "|")(1), Split(Astring(x), "|")(2), MyTip.ListView
Next
Else
AddListItem Split(VirusDes, "|")(0), Split(VirusDes, "|")(1), Split(VirusDes, "|")(2), MyTip.ListView
End If


.Show
Do Until .Visible = False
SuperSleep 1
Loop

ShowTextTip = MyTip.ChooseMod
End With
End Function

Private Function AddListItem(ByVal FirstText As String, ByVal SecondText As String, ByVal ThirdText As String, ByRef List As ListView)
Set itm = List.ListItems.Add(, , FirstText)
itm.SubItems(1) = SecondText
itm.SubItems(2) = ThirdText
End Function



