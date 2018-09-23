Attribute VB_Name = "mdProcessScan"

Public Function ProcessScan(ByVal Path As String) As String
'=========进程扫描模块========
'返回值类型：
'安全：SAFE
'出错：Error
'正常：返回病毒内容，格式：病毒名|病毒描述
'=========函数内容========
'On Error GoTo Err:
DoEvents '转让控制权
Dim Filestring As String
Dim FindString As String
Filestring = GetChecksum(Path) '获得文件特征码
Debug.Print Filestring
FindString = FindVirus(Filestring) '查找病毒
If FindString <> "SAFE" Then '如果在病毒库中找到了相应的病毒
ProcessScan = FindString '返回病毒内容，格式：病毒名|病毒描述
Else  '如果安全
ProcessScan = "SAFE"
End If
Exit Function
Err:
ProcessScan = "Error" '如果出错，返回“ERROR”

End Function

Private Function FindVirus(ByVal MD5 As String) As String
'搜查病毒库
On Error Resume Next
Set FindString = frmData.VirusData.FindItem(MD5)
If FindString Is Nothing Then  '没找到
  '返回：“SAFE”
    FindVirus = "SAFE"
    Exit Function
Else
  '返回：病毒名|病毒描述
  FindVirus = frmData.VirusData.ListItems(FindString.Index).SubItems(1) & "|" & frmData.VirusData.ListItems(FindString.Index).SubItems(2)
  Exit Function
End If
FindVirus = "SAFE"

End Function

