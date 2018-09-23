VERSION 5.00
Begin VB.Form frmControl 
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3420
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   3420
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function CallMode(ByVal mode) As Boolean
On Error Resume Next
Dim Pid As Double
Pid = 0
Select Case mode
Case 1
'进程实时防护
If exitproc("ProcessRTA.exe") = True Then
  strShare = "ProcessRTA"
  SuperSleep 1
  strShare = "ProcessRTA.Unload"
  Call WriteString("Main", "ProcessRTA", 0, IniPath)
  Call ShowTip("光子防御网", "进程/驱动实时防护已关闭", 4)
Else
  If Not Dir(App.Path & "\ProcessRTA.exe") = "" Then
  Pid = Shell(App.Path & "\ProcessRTA.exe")
  End If
  Call WriteString("Main", "ProcessRTA", 1, IniPath)
  Call ShowTip("光子防御网", "进程/驱动实时防护已开启", 4)
End If

Case 2
'驱动实时防护
If exitproc("ProcessRTA.exe") = True Then
'  Auto.Enabled = False
  strShare = "ProcessRTA"
  SuperSleep 1
  strShare = "ProcessRTA.Unload"
  'MsgBox "抱歉，暂时无法关闭，我正在努力解决这个问题，重启后将关闭。", vbOK, "光子防御网"
  Call WriteString("Main", "ProcessRTA", 0, IniPath)
  Call ShowTip("光子防御网", "进程/驱动实时防护已关闭", 4)
Else
  If Not Dir(App.Path & "\ProcessRTA.exe") = "" Then
  Pid = Shell(App.Path & "\ProcessRTA.exe")
  End If
  Call WriteString("Main", "ProcessRTA", 1, IniPath)
  Call ShowTip("光子防御网", "进程/驱动实时防护已开启", 4)
End If

Case 3
'注册表实时防护

If exitproc("RegRTA.exe") = True Then
  strShare = "RegRTA"
  SuperSleep 1
  strShare = "RegRTA.Unload"
  Call WriteString("Main", "RegRTA", 0, IniPath)
  Call ShowTip("光子防御网", "注册表实时防护已关闭", 4)
Else
  '开启模块
  If Not Dir(App.Path & "\RegRTA.exe") = "" Then
  Pid = Shell(App.Path & "\RegRTA.exe")
  End If
  Call WriteString("Main", "RegRTA", 1, IniPath)
  Call ShowTip("光子防御网", "注册表实时防护已开启", 4)
End If
Case 4
If exitproc("USBRTA.exe") = True Then
  strShare = "USBRTA"
  SuperSleep 1
  strShare = "USBRTA.Close"
  Call WriteString("Main", "USBRTA", 0, IniPath)
  Call ShowTip("光子防御网", "U盘插入防护已关闭", 4)
Else
  If Not Dir(App.Path & "\USBRTA.exe") = "" Then
  Pid = Shell(App.Path & "\USBRTA.exe")
  End If
  Call WriteString("Main", "USBRTA", 1, IniPath)
  Call ShowTip("光子防御网", "U盘插入防护已开启", 4)
End If
Case 5
If exitproc("ProtectProcess.exe") = True Then
  strShare = "Protect"
  SuperSleep 1
  strShare = "Protect.Unload"
  Call ShowTip("光子防御网", "自我保护已关闭", 4)
Else
  If Not Dir(App.Path & "\Protect.exe") = "" Then
  Pid = Shell(App.Path & "\Protect.exe")
  End If
  Call ShowTip("光子防御网", "自我保护已开启", 4)
End If
End Select
SuperSleep 1
strShare = "Protect"
SuperSleep 1
strShare = "Protect.ReLoad"

frmMain.ReRead
End Function

