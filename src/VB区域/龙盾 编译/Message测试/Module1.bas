Attribute VB_Name = "mdMsg"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Const WM_MYMESSAGE = &H400 + 100
Const WM_USER = &H400
Public c As Long
Public KillShit As Boolean
Public Declare Function CheckFileDigitalSignature_Ansi Lib "DSDigitalSignature.dll" (ByVal CheckFileDigitalSignature_Ansi As String) As Long


Public Function Wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo Err:
Dim RecStr As String
Dim Mode
Dim DesStr
Dim DesString
Dim cds As COPYDATASTRUCT
    If Msg = WM_MYMESSAGE Then
     'MsgBox "开始……"
    RecStr = Form_Map.ReadMem
    
      'MsgBox RecStr
      '获得内容
      
      DesString = Split(RecStr, "|")(0)
      DesStr = Split(RecStr, "|")(1)
      Mode = Split(RecStr, "|")(2)
      '切分
      Dim WtStr
     ' MsgBox DesString
     '  MsgBox DesStr
      '  MsgBox Mode
     ' '\Registry\Machine\System\CurrentControlSet\Services\SuperKillFile|services.exe|DRV
      Select Case Mode
        Case "PRO" '进程
       ' MsgBox "进程"
       KillShit = False
        WriteProLog "==============================="
         WriteProLog "时间：" & Now
         WriteProLog DesStr & "正在创建进程：" & DesString
         If CheckProcess(DesStr, DesString) = True Then '放行
         WtStr = "Allow"
         Else
         WtStr = "Disallow"
         If KillShit = True Then '要删除
           Dim NewKill As New Killer
           NewKill.KillName = DesString
           Load NewKill
         End If
         End If
          Call WriteString("RTA", "Message", WtStr, _
          App.Path & "\Chat.ini")
        '写入信息
        Case "DRV"
        WriteDrvLog "==============================="
        'MsgBox "驱动"
         WriteDrvLog "时间：" & Now
         WriteDrvLog "驱动：" & DesString
         If CheckDriver(DesStr, DesString) = True Then '放行
         WtStr = "Allow"
         Else
         WtStr = "Disallow"
         End If
         
      End Select
    Form_Map.WriteMem (WtStr)
    End If
Err:
    Wndproc = CallWindowProc(c, hwnd, Msg, wParam, lParam)
End Function




Public Function CheckProcess(ByVal FromProcessPath As String, ByVal ProcessPath As String) As Boolean
On Error GoTo Err:

Dim YesNo As Boolean

Dim ID As String, Result As String
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
If Result = "" Then '没有记录，就新增一个默认的记录
  WriteString "Rules", """" & ProcessPath & """", "2", App.Path & "\Rules.ini"
End If
'重新读取
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
 If Result = "1" Then '信任的东西
 WriteProLog "结果：信任的目标，自动放行"
 CheckProcess = True
 Exit Function
 ElseIf Result = "0" Then '不信任的东西
 WriteProLog "结果：不信任的目标，自动拦截"
 CheckProcess = False
 Exit Function
 End If
 '数字签名验证
 'MsgBox "验证" & ProcessPath
 If CheckFileDigitalSignature_Ansi(ProcessPath) = "1" Then '已经有数字签名
  WriteProLog "结果：拥有数字签名，自动放行"
  CheckProcess = True
  Exit Function
  Else
  WriteProLog "无数字签名"
  End If
 '默认2
  '进行病毒扫描
 Dim ScanResult

 ScanResult = ProcessScan(ProcessPath)
If ScanResult <> "SAFE" And ScanResult <> "Error" Then

  
   Dim VirusForm As New frmVirusTip
    VirusForm.PicIcon = GetIconFromFile(ProcessPath, 0, True)
    VirusForm.Tip.Caption = "发现病毒正在运行，已被病毒防御助手拦截，病毒无法启动。"
    VirusForm.TextRes.Text = ProcessPath & vbCrLf & _
    "病毒名：" & Split(ScanResult, "|")(0) & vbCrLf & _
    "病毒描述：" & Split(ScanResult, "|")(1)
    VirusForm.Show vbModal
     WriteProLog "病毒：" & ProcessPath
     WriteProLog "病毒描述：" & ScanResult
     WriteProLog "已结束进程"
     CheckProcess = False
     Exit Function
Else
Dim MyForm As New frmTip
MyForm.PicIcon = GetIconFromFile(ProcessPath, 0, True)
MyForm.Text1.Text = "进程：" & FromProcessPath & vbCrLf & "正在创建进程：" & ProcessPath
Dim MyFSO As New FileSystemObject
Dim StrDrv As String
StrDrv = Left(ProcessPath, 3)
If Right(StrDrv, 2) <> ":\" Then '如果不是标准路径名
  MyForm.Option2.Value = True
MyForm.Tip = "可疑进程正在创建中，请不要运行来历不明的文件！如果您是打开文件夹，则此为伪装的应用程序。"
GoTo Kip:
End If
If MyFSO.GetDrive(StrDrv).DriveType <> Fixed Then '如果是不是本地出现的东西
 MyForm.Option2.Value = True
MyForm.Tip = "非本地磁盘中正在运行可疑进程，请不要运行来历不明的文件！如果您是打开文件夹，则此为伪装的应用程序。"
Else
 MyForm.Option1.Value = True
MyForm.Tip = "本地磁盘中正在运行进程，请不要运行来历不明的文件！由于在本地磁盘中，可能是系统自动运行的程序，默认30秒放行。"
End If
Kip:
MyForm.Command1.Caption = "Ｘ"
MyForm.Show vbModal

'如果选择以后也这么处理
If MyForm.ChooseNum <> 1 And MyForm.ChooseNum <> 2 Then
Dim MyForm2 As New frmTip
MyForm2.PicIcon = GetIconFromFile(ProcessPath, 0, True)
   If MyForm.ChooseNum = 3 Then
   WriteString "Rules", """" & ProcessPath & """", "1", App.Path & "\Rules.ini"
   MyForm2.Text1 = "文件：" & ProcessPath & vbCrLf & "已经添加到龙盾的信任列表，不拦截，不扫描。"
   WriteProLog "结果：用户选择添加到信任列表"
   ElseIf MyForm.ChooseNum = 4 Then
   MyForm2.Text1 = "文件：" & ProcessPath & vbCrLf & "已经添加到龙盾的黑名单列表，禁止运行，禁止操作"
   WriteString "Rules", """" & ProcessPath & """", "0", App.Path & "\Rules.ini"
   WriteProLog "结果：用户选择添加到黑名单列表"
   End If
MyForm2.Option1.Visible = False
MyForm2.Option2.Visible = False
MyForm2.Check1.Visible = False
MyForm2.Command2.Caption = "我知道了"
MyForm2.Label2.Visible = False
MyForm2.Label3.Caption = "添加规则"
MyForm2.Show
End If

If MyForm.ChooseNum = 1 Then
WriteProLog "已允许运行"
CheckProcess = True
ElseIf MyForm.ChooseNum = 2 Then
WriteProLog "已阻止运行"
CheckProcess = False
ElseIf MyForm.ChooseNum = 3 Then
WriteProLog "已允许运行"
CheckProcess = True
ElseIf MyForm.ChooseNum = 4 Then
WriteProLog "已阻止运行"
CheckProcess = False
End If
Exit Function
End If
Err:
WriteProLog "执行出错"
CheckProcess = True
End Function


Public Function CheckDriver(ByVal FromProcessPath As String, ByVal ProcessPath As String) As Boolean
'用ProcessPath替代DrivePath
On Error GoTo Err:

Dim YesNo As Boolean

Dim ID As String, Result As String
Result = ReadString("DriverRules", """" & ProcessPath & """", App.Path & "\Rules.ini")
If Result = "" Then '没有记录，就新增一个默认的记录
  WriteString "DriverRules", """" & ProcessPath & """", "2", App.Path & "\Rules.ini"
End If
'重新读取
Result = ReadString("DriverRules", """" & ProcessPath & """", App.Path & "\Rules.ini")
 If Result = "1" Then '信任的东西
 WriteDrvLog "结果：信任的驱动，自动放行。"
 CheckDriver = True
 Exit Function
 ElseIf Result = "0" Then '不信任的东西
 WriteDrvLog "结果：黑名单中的驱动，自动阻止。"
 CheckDriver = False
 Exit Function
 End If
 '数字签名验证
  If CheckFileDigitalSignature_Ansi(FromProcessPath) = "1" Then '已经有数字签名
  WriteProLog "结果：拥有数字签名，自动放行"
  CheckDriver = True
  Exit Function
  End If
 '默认2
Dim MyForm As New frmTip
'MyForm.PicIcon = GetIconFromFile(ProcessPath, 0, True)
MyForm.Text1.Text = "进程：" & FromProcessPath & vbCrLf & "正在加载驱动：" & ProcessPath
MyForm.Label3.Caption = "可疑驱动加载"
MyForm.Tip = "发现有程序正在绕过安全软件的检测，驱动程序掌管着系统最高权限（Ring0），如果被未知程序加载将有很大风险，加载驱动后将无法拦截到驱动的操作。"
If LCase(FromProcessPath) = "services.exe" Then '如果是系统加载
MyForm.Option1.Value = True
Else
MyForm.Option2.Value = True
End If
MyForm.Option2.Caption = "阻止此操作"
Kip:
MyForm.Command1.Caption = "Ｘ"
MyForm.Show vbModal

'如果选择以后也这么处理
If MyForm.ChooseNum <> 1 And MyForm.ChooseNum <> 2 Then
Dim MyForm2 As New frmTip
MyForm2.PicIcon = GetIconFromFile(ProcessPath, 0, True)
   If MyForm.ChooseNum = 3 Then
   WriteString "DriverRules", """" & ProcessPath & """", "1", App.Path & "\Rules.ini"
   MyForm2.Text1 = "驱动：" & ProcessPath & vbCrLf & "已经添加到龙盾的信任列表，不拦截，不扫描。"
   WriteDrvLog "结果：用户选择添加到信任列表"
   ElseIf MyForm.ChooseNum = 4 Then
   MyForm2.Text1 = "驱动：" & ProcessPath & vbCrLf & "已经添加到龙盾的黑名单列表，禁止运行，禁止操作"
   WriteString "DriverRules", """" & ProcessPath & """", "0", App.Path & "\Rules.ini"
 WriteDrvLog "结果：用户选择添加到黑名单"
   End If
MyForm2.Option1.Visible = False
MyForm2.Option2.Visible = False
MyForm2.Check1.Visible = False
MyForm2.Command2.Caption = "我知道了"
MyForm2.Label2.Visible = False
MyForm2.Label3.Caption = "添加规则"
MyForm2.Show
End If

If MyForm.ChooseNum = 1 Then
WriteDrvLog "已允许操作"
CheckDriver = True
ElseIf MyForm.ChooseNum = 2 Then
WriteDrvLog "已拦截操作"
CheckDriver = False
ElseIf MyForm.ChooseNum = 3 Then
WriteDrvLog "已允许操作"
CheckDriver = True
ElseIf MyForm.ChooseNum = 4 Then
WriteDrvLog "已拦截操作"
CheckDriver = False
End If
Exit Function
Err:
CheckDriver = True
End Function


Public Function WriteProLog(ByVal Text As String)
Open App.Path & "\ProLog.dat" For Binary As #1
    Put #1, LOF(1) + 1, Text & vbCrLf
Close #1
End Function

Public Function WriteDrvLog(ByVal Text As String)
Open App.Path & "\DrvLog.dat" For Binary As #1
    Put #1, LOF(1) + 1, Text & vbCrLf
Close #1
End Function

