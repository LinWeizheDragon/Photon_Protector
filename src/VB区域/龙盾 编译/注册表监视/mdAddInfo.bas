Attribute VB_Name = "mdAddInfo"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Load_virus As Boolean
Public Load_Main As Boolean
Public Load_Watch As Boolean

Public Sub AddInfo(ByVal TextT As String)
Set lst = frmMain.InfoView.ListItems.Add(1, , Time)
lst.SubItems(1) = TextT
End Sub
Public Function SuperSleep(DealyTime As Single) '´Ë´¦Ô­Îªlong£¬ÐÞ¸ÄÎªsingle¿ÉÑÓÊ±1ms :SK<2<8h
Dim TimerCount As Single
    TimerCount = Timer + DealyTime 'Ôö¼ÓXÃë ZJ9x6|q
    While TimerCount - Timer > 0
        DoEvents
        Sleep 1
    Wend
    Text1 = "SuperSleep " & DealyTime
End Function

