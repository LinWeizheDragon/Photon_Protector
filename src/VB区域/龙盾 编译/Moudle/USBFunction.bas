Attribute VB_Name = "HookFunction"
Option Explicit
'写入到配置文件中去
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'获取配置文件中的值
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



'子类化窗体消息处理函数时需要使用的API，很常见，不作过多说明。
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const GWL_WNDPROC = -4
Private Const WM_DEVICECHANGE As Long = &H219
Private Const DBT_DEVICEARRIVAL As Long = &H8000&
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&
'设备类型：逻辑卷标
Private Const DBT_DEVTYP_VOLUME As Long = &H2

'与WM_DEVICECHANGE消息相关联的结构体头部信息

Private Type DEV_BROADCAST_HDR
    lSize As Long
    lDevicetype As Long   '设备类型
    lReserved As Long
End Type

'设备为逻辑卷时对应的结构体信息
Private Type DEV_BROADCAST_VOLUME
    lSize As Long
    lDevicetype As Long
    lReserved As Long
    lUnitMask As Long   '和逻辑卷标对应的掩码
    iFlag As Integer
End Type

Private info As DEV_BROADCAST_HDR
Private info_volume As DEV_BROADCAST_VOLUME

Private PrevProc As Long  '原来的窗体消息处理函数地址


Public Sub HookForm(f As Form)

    PrevProc = SetWindowLong(f.hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub UnHookForm(f As Form)

    SetWindowLong f.hwnd, GWL_WNDPROC, PrevProc

End Sub


Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg
    
        '插入USB DISK 则接收到此消息
        Case WM_DEVICECHANGE
         If wParam = DBT_DEVICEARRIVAL Then
              '若插入USBDISK或者映射网络盘等则
              'info.lDevicetype =2
              '即DBT_DEVTYP_VOLUME
              '利用参数lParam获取结构体头部信息
              CopyMemory info, ByVal lParam, Len(info)
             If info.lDevicetype = DBT_DEVTYP_VOLUME Then
               CopyMemory info_volume, ByVal lParam, Len(info_volume)
               '检测到有逻辑卷添加到系统中，则显示该设备根目录下全部文件名
               Dim drivepathname
               drivepathname = Chr(GetDriveName(info_volume.lUnitMask))

               Call ShowTip("龙盾", "发现" & drivepathname & ":\插入，正在进行扫描", 4)
               '加载程序日志
Loap:
               If Dir(drivepathname & ":\") = "" Then
               GoTo Loap:
               End If
           '    frmMain.ScanOrder.AddItem drivepathname & ":\"
            '   CheckOrder
'               frmMain.CheckTimer.Enabled = True
'               frmMain.ScanList.Drive = DrivePathName & ":\"
'               If frmMain.OpSim.Value = True Then
'               ListFiles DrivePathName & ":\", frmMain.lstFiles
'               ElseIf frmMain.OpAdv.Value = True Then
'               sousuofile DrivePathName & ":\"
'               End If
             End If
         End If
         
         If wParam = DBT_DEVICEREMOVECOMPLETE Then
             '若移走USBDISK或者映射网络盘等则
             'info.lDevicetype =2
             '即DBT_DEVTYP_VOLUME
             '利用参数lParam获取结构体头部信息
             CopyMemory info, ByVal lParam, Len(info)
             If info.lDevicetype = DBT_DEVTYP_VOLUME Then
               CopyMemory info_volume, ByVal lParam, Len(info_volume)
        '       Call ShowTip("U盘Autorun病毒监控模块", "移动设备" & Chr(GetDriveName(info_volume.lUnitMask)) & ":\" & "已拔出！", 4)
         '      frmLog.LogText.Text = frmLog.LogText.Text & vbCrLf & "检测到" & Chr(GetDriveName(info_volume.lUnitMask)) & ":\拔出.....时间：" & Format(Now)

               '清除LIST中的内容
       '        frmMain.lstFiles.Clear
             End If
         End If
           
     End Select

    ' 调用原来的窗体消息处理函数
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)

End Function

'根据输入的32位LONG型数据（只有一位为1）返回对应的卷标的ASCII数值
'规则是1：A、2：B、4：C等等
Private Function GetDriveName(ByVal lUnitMask As Long) As Byte
    Dim I As Long
    I = 0
    
    While lUnitMask Mod 2 <> 1
       lUnitMask = lUnitMask \ 2
       I = I + 1
    Wend
    
    GetDriveName = Asc("A") + I
End Function

'显示插入逻辑卷根目录的文件名列表，需要在工程里引用Microsoft Scripting Runtime库。
Private Function ListFiles(strPath As String, ByRef List As ListBox)
  Dim fso As New Scripting.FileSystemObject
  Dim objFolder As Folder
  Dim objFile As File

  Set objFolder = fso.GetFolder(strPath)

  For Each objFile In objFolder.Files
    List.AddItem strPath & objFile.Name
  Next
End Function

Public Function ListDiskFiles(strPath As String, ByRef List As ListBox)
ListFiles strPath, List
End Function

'读ini文件
 Public Function ReadString(ByVal Caption As String, ByVal Item As String, ByVal Path As String) As String
    On Error Resume Next
    Dim sBuffer As String
    sBuffer = Space(32767)
    GetPrivateProfileString Caption, Item, vbNullString, sBuffer, 32766, Path
    ReadString = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
 End Function

'写ini文件
 Public Function WriteString(ByVal Caption As String, ByVal Item As String, ByVal ItemValue As String, ByVal Path As String) As Long
    Dim sBuffer As String
    sBuffer = Space(32766)
    sBuffer = ItemValue & vbNullChar
    WriteString = WritePrivateProfileString(Caption, Item, sBuffer, Path)
 End Function

