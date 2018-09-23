Attribute VB_Name = "Module1"
Public Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Download by http://www.codefans.net
Public Declare Function CallWindowProc Lib "User32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub RtlZeroMemory Lib "ntdll.dll" (dest As Any, ByVal numBytes As Long)
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long



Public Const GWL_WNDPROC = -4
Public oldWNDPROC, newWNDPROC As Long
Public ExeFiles As String
Public CommandLines As String

Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

        Dim tmp As Long

        If msg = &H400 Then

                tmp = pShareMem + 12
                
                wcscpy ByVal StrPtr(ExeFiles), ByVal tmp
                RtlZeroMemory ByVal tmp, 1000
                tmp = pShareMem + 12 + 1000
                
                wcscpy ByVal StrPtr(CommandLines), ByVal tmp
                RtlZeroMemory ByVal tmp, 1000
                
                tmp = InStr(1, ExeFiles, Chr(0))

                If tmp > 1 Then
                        ExeFiles = Left$(ExeFiles, tmp - 1)
                
                End If

                tmp = InStr(1, CommandLines, Chr(0))

                If tmp > 1 Then
                        CommandLines = Left$(CommandLines, tmp - 1)
                
                End If


                If Len(ExeFiles) = 0 Then
                        'tmp = MessageBox(Form1.hwnd, "程序 " & CommandLines & "试图运行,是否允许?", "提示", vbYesNo + 32)
                        tmp = CheckProcess(CommandLines, "")

                        If tmp = vbYes Then
                                WndProc = 1
                        Else
                                WndProc = 0
                        End If

                Else
                        'tmp = MessageBox(Form1.hwnd, "程序 " & ExeFiles & " 试图执行 " & Chr(13) & "命令行:" & CommandLines & Chr(13) & "是否允许?", "提示", vbYesNo + 32)
                        tmp = CheckProcess(CommandLines, ExeFiles)
                        If tmp = True Then
                                WndProc = 1
                        Else
                                WndProc = 0
                        End If
                
                End If
                CommandLines = Space$(1000)
                ExeFiles = Space$(1000)
        Else
       
                WndProc = CallWindowProc(oldWNDPROC, hwnd, msg, wParam, lParam)
    
        End If

End Function

Public Function CheckProcess(ByVal CmdLine As String, ByVal ProcessPath As String) As Boolean
If ProcessPath = "" Then
CheckProcess = True
Exit Function
End If
On Error GoTo ERR:
Dim YesNo As Boolean
Dim ID As String, Result As String
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
If Result = "" Then '没有记录，就新增一个默认的记录
  WriteString "Rules", """" & ProcessPath & """", "2", App.Path & "\Rules.ini"
End If
'重新读取
Result = ReadString("Rules", """" & ProcessPath & """", App.Path & "\Rules.ini")
 If Result = "1" Then '信任的东西
 CheckProcess = True
 Exit Function
 ElseIf Result = "0" Then '不信任的东西
 CheckProcess = False
 Exit Function
 End If
 '默认2
Dim MyForm As New frmTip
MyForm.PicIcon = GetIconFromFile(ProcessPath, 0, True)
MyForm.Text1.Text = "此进程正在创建中：" & ProcessPath & vbCrLf & "命令行：" & CmdLine
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
   ElseIf MyForm.ChooseNum = 4 Then
   MyForm2.Text1 = "文件：" & ProcessPath & vbCrLf & "已经添加到龙盾的黑名单列表，禁止运行，禁止操作"
   WriteString "Rules", """" & ProcessPath & """", "0", App.Path & "\Rules.ini"
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
CheckProcess = True
ElseIf MyForm.ChooseNum = 2 Then
CheckProcess = False
ElseIf MyForm.ChooseNum = 3 Then
CheckProcess = True
ElseIf MyForm.ChooseNum = 4 Then
CheckProcess = False
End If

ERR:
End Function
