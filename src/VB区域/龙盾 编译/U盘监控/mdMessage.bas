Attribute VB_Name = "mdMessage"
Public strShare As New CSharedString

Public Function StartMis(ByVal Path As String, ByVal way As Integer)
'Path:The path of the usb driver.
'Way:The message type of click
If way = 1 Then '打开U盘
If Dir(Path) <> "" Then
OpenFile Path, , vbNormalFocus
Else
MsgBox "路径不正确，打开U盘失败！"
End If
ElseIf way = 2 Then '修复U盘
FixUSB Path
'Shell App.Path & "\ScanMod.exe"
'strShare = "ScanMod"
'SuperSleep 1
'strShare = "ScanMod.Scan.Adv" & Path
ElseIf way = 3 Then '退出U盘
RemoveUSB Path
End If
frmMain.ReReadDrv
End Function

Public Function FixUSB(ByVal Path As String)
On Error Resume Next
DoEvents
Form1.Label1.Caption = "病毒防御助手 工作中"
                WriteLog now & ":"
                WriteLog "--修复可移动设备" & Path
SuperSleep 1
Dim OldPath
OldPath = Path
     Dim fs, f, f1, s, sf
     Dim hs, h, h1, hf
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(OldPath)
     Set sf = f.SubFolders
     For Each f1 In sf
     If Right(OldPath, 1) = "\" Then
     Path = OldPath & f1.Name
     Else
     Path = OldPath & "\" & f1.Name
     End If
     Debug.Print Path
     DoEvents
Call SetAttr(Path, vbNormal)
FileCopy App.Path & "\desktop.dll", Path & "\desktop.ini"
FileCopy App.Path & "\文件夹安全验证图标.dll", Path & "\文件夹安全验证图标.ico"
Call SetAttr(Path, vbSystem)
Debug.Print App.Path & "\desktop.dll" & "----" & Path & "\desktop.ini"
     Next
OpenFile OldPath, , vbNormalFocus
Form1.Label1.Caption = "病毒防御助手 U盘监控"

End Function

Public Function RemoveUSB(ByVal Path As String)
    Dim lngLenPath As Long, blnIsUsb As Boolean, strPath As String
    lngLenPath = Len(Path)
    If lngLenPath <= 3 And Dir(Path, 1 Or 2 Or 4 Or vbDirectory) <> "" Then
        If lngLenPath = 2 Then
            If GetDriveBusType(Path) <> "Usb" Then
             '   MsgBox "只能解锁USB设备分区！！", vbCritical, "错误"
                
                Exit Function
            End If
            strPath = Path & "\"
        ElseIf lngLenPath = 1 Then
            If GetDriveBusType(Path & ":") <> "Usb" Then
            '    MsgBox "只能解锁USB设备分区！！", vbCritical, "错误"
                
                Exit Function
            End If
            strPath = Path & ":\"
        Else
            If GetDriveBusType(Left(Path, 2)) <> "Usb" Then
               ' MsgBox "只能解锁USB设备分区！", vbCritical, "错误"
                txtUsbDrive.SetFocus
                Exit Function
            End If
            strPath = Path
        End If
        blnIsUsb = True
    Else
     '   MsgBox "USB盘符不符合要求！", vbCritical, "错误"
        
        Exit Function
    End If
    '这里只检测本进程因为在获取驱动器类型的时候会打开一个句柄但是WINDOWS没有自己关闭所以用这个来
    '解除锁定，当然你也可以使用CloseLoackFiles函数来检测所有进程
    If CloseLockFileHandle(Left(strPath, 2), GetCurrentProcessId) Then
        If blnIsUsb Then
            If RemoveUsbDrive("\\.\" & Left(strPath, 2), True) Then
                'MsgBox "卸载UBS设备成功！", , "提示"
            Else
                'MsgBox "卸载UBS设备失败！", vbCritical, "提示"
            End If
        End If
    Else
       ' MsgBox "发现有锁定文件还没解锁！", vbCritical, "提示"
    End If
End Function
