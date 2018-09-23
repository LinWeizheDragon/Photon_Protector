Attribute VB_Name = "mod_Msg"
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
Public c As Long
Public Function Wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ERR:
Dim s As String
Dim cds As COPYDATASTRUCT
    If Msg = &H4A Then
    CopyMemory cds, ByVal lParam, Len(cds)
    s = Space(cds.cbData)
    CopyMemory ByVal s, ByVal cds.lpData, cds.cbData
    s = StrConv(s, vbFromUnicode)
    s = Left(s, InStr(1, s, Chr(0)) - 1)
        If CheckProcess(wParam, s) = True Then
            Wndproc = 1234
        Else
            Wndproc = 0
        End If
        Exit Function

'    s = "进程(Pid:" & wParam & ")要创建新进程: " & s & ",是否允许?"
'    Debug.Print s
'        If MessageBoxA(0, s, "", 4) = 6 Then
'            Wndproc = 1234
'        Else
'            Wndproc = 0
'        End If
'        Exit Function
    End If
ERR:
    Wndproc = CallWindowProc(c, hwnd, Msg, wParam, lParam)
End Function

Public Function CheckProcess(ByVal FromProcessID As String, ByVal ProcessPath As String) As Boolean
If FromProcessID = "0" Then
Exit Function
End If
On Error GoTo ERR:
Dim FromProcessPath As String
Dim YesNo As Boolean
FromProcessPath = GetProcessPath(FromProcessID)
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
MyForm.Text1.Text = "进程：" & FromProcessPath & vbCrLf & "父进程ID：" & FromProcessID & vbCrLf & "正在创建进程：" & ProcessPath
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
