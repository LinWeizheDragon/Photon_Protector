Attribute VB_Name = "mdScan"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'Í£¶ÙAPi
Public Function ProcessScan(ByVal Path As String) As String
'=========½ø³ÌÉ¨ÃèÄ£¿é========
'·µ»ØÖµÀàĞÍ£º
'°²È«£ºSAFE
'³ö´í£ºError
'Õı³££º·µ»Ø²¡¶¾ÄÚÈİ£¬¸ñÊ½£º²¡¶¾Ãû|²¡¶¾ÃèÊö
'=========º¯ÊıÄÚÈİ========
'On Error GoTo Err:
DoEvents '×ªÈÃ¿ØÖÆÈ¨
Dim Filestring As String
Dim FindString As String
Filestring = GetChecksum(Path) '»ñµÃÎÄ¼şÌØÕ÷Âë
'InputBox "hh", "hh", Filestring
FindString = FindVirus(Filestring) '²éÕÒ²¡¶¾
If FindString <> "SAFE" Then 'Èç¹ûÔÚ²¡¶¾¿âÖĞÕÒµ½ÁËÏàÓ¦µÄ²¡¶¾
ProcessScan = FindString '·µ»Ø²¡¶¾ÄÚÈİ£¬¸ñÊ½£º²¡¶¾Ãû|²¡¶¾ÃèÊö
Else  'Èç¹û°²È«
ProcessScan = "SAFE"
End If
Exit Function
Err:
ProcessScan = "Error" 'Èç¹û³ö´í£¬·µ»Ø¡°ERROR¡±

End Function

Private Function FindVirus(ByVal MD5 As String) As String
'ËÑ²é²¡¶¾¿â
On Error Resume Next
Set FindString = frmData.VirusData.FindItem(MD5)
If FindString Is Nothing Then  'Ã»ÕÒµ½
  '·µ»Ø£º¡°SAFE¡±
    FindVirus = "SAFE"
    Exit Function
Else
  '·µ»Ø£º²¡¶¾Ãû|²¡¶¾ÃèÊö
  FindVirus = frmData.VirusData.ListItems(FindString.Index).SubItems(1) & "|" & frmData.VirusData.ListItems(FindString.Index).SubItems(2)
  Exit Function
End If
FindVirus = "SAFE"

End Function

Public Function SuperSleep(DealyTime As Single) '´Ë´¦Ô­Îªlong£¬ĞŞ¸ÄÎªsingle¿ÉÑÓÊ±1ms :SK<2<8h
Dim TimerCount As Single
    TimerCount = Timer + DealyTime 'Ôö¼ÓXÃë
    While TimerCount - Timer > 0
        DoEvents
        Sleep 1
    Wend
    Text1 = "SuperSleep " & DealyTime
End Function

