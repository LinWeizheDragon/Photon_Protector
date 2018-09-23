Attribute VB_Name = "modWinVer"

Public Function WinVer() As Single
 Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
 Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 For Each objItem In colItems
  DoEvents
  strOSversion = objItem.Version
 Next
 strOSversion = Left(strOSversion, 3)
 WinVer = Val(strOSversion)
 Exit Function '½áÊø
 Select Case strOSversion
   Case "5.0"
     strOSversion = "Windows 2000"
    Case "5.1"
     strOSversion = "Windows XP"
    Case "5.2"
     strOSversion = "Windows Server 2003"
    Case "6.0"
     strOSversion = "Windows Vista"
    Case "6.1"
     strOSversion = "Windows 7"
    End Select
  MsgBox strOSversion
End Function


