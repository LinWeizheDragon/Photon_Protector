VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmProc 
   Caption         =   "进程扫描"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmProc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdExit 
      Caption         =   "完成"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "删除选中项目"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwPrss 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '获得进程的句柄
    Private Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, _
            ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
              
    '终止进程
    Private Declare Function TerminateProcess Lib "KERNEL32" (ByVal ApphProcess As Long, _
            ByVal uExitCode As Long) As Long
    '创建一个系统快照
    Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" _
            (ByVal lFlags As Long, lProcessID As Long) As Long
    '获得系统快照中的第一个进程的信息
    Private Declare Function ProcessFirst Lib "KERNEL32" Alias "Process32First" _
            (ByVal mSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    '获得系统快照中的下一个进程的信息
    Private Declare Function ProcessNext Lib "KERNEL32" Alias "Process32Next" _
            (ByVal mSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Private Type PROCESSENTRY32
        dwSize As Long
        cntUsage As Long
        th32ProcessID As Long
        th32DefaultHeapID As Long
        th32ModuleID As Long
        cntThreads As Long
        th32ParentProcessID As Long
        pcPriClassBase As Long
        dwFlags As Long
        szExeFile As String * 260&
    End Type
    Private Const TH32CS_SNAPPROCESS As Long = 2&
    Dim mresult

Private Sub CmdDelete_Click()
On Error Resume Next
    If lvwPrss.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    Dim Name As String
        Dim mProcID As Long
        Dim i
        Dim itm
    For i = 1 To lvwPrss.ListItems.Count
      Set itm = lvwPrss.ListItems(i)
      If itm.Checked = True Then
       If Name <> "" Then
        Name = Name & vbCrLf & itm.SubItems(1)
       Else
        Name = itm.SubItems(1)
       End If
      End If
    Next
      If Name = "" Then Exit Sub
    
    For i = 1 To lvwPrss.ListItems.Count
      Set itm = lvwPrss.ListItems(i)
      If itm.Checked = True Then
    '打开进程
    mProcID = OpenProcess(1&, -1&, itm.Text)
    '终止进程
    TerminateProcess mProcID, 0&
       End If
      
    Next
    DoEvents
    doList

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
   '配置ListView控件。
    lvwPrss.ListItems.Clear
    lvwPrss.ColumnHeaders.Clear
    lvwPrss.ColumnHeaders.Add , , "进程ID", 1500
    lvwPrss.ColumnHeaders.Add , , "进程名", 3000
    lvwPrss.ColumnHeaders.Add , , "进程路径", 10000
    lvwPrss.LabelEdit = lvwManual
    lvwPrss.FullRowSelect = True
    lvwPrss.HideSelection = False
    lvwPrss.HideColumnHeaders = False
    lvwPrss.View = lvwReport
    lvwPrss.Checkboxes = True
    doList
End Sub
Private Sub doList()
On Error Resume Next
    Dim uProcess As PROCESSENTRY32
    Dim mSnapShot As Long
    Dim mName As String
    Dim i As Integer
    Dim mlistitem As ListItem
    Dim msg As String
    lvwPrss.ListItems.Clear
    DoEvents
    '获取进程长度？？
    uProcess.dwSize = Len(uProcess)
    '创建一个系统快照
    mSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If mSnapShot Then
        '获取第一个进程
        mresult = ProcessFirst(mSnapShot, uProcess)
        '失败则返回false
        Do While mresult
            '返回进程长度值+1，Chr(0)的作用：结束语，防止修改进程
            i = InStr(1, uProcess.szExeFile, Chr(0))
            '转换成小写
            mName = LCase$(Left$(uProcess.szExeFile, i - 1))
            '在listview控件中添加当前进程名
            Debug.Print ProcessScan(GetProcessPath(mName)) & "|" & mName
            
            If ProcessScan(GetProcessPath(mName)) <> "SAFE" And ProcessScan(GetProcessPath(mName)) <> "Error" Then
            Set mlistitem = lvwPrss.ListItems.Add(, , Text:=uProcess.th32ProcessID)
            '添加进程名
            mlistitem.SubItems(1) = mName
            mlistitem.SubItems(2) = GetProcessPath(mName)
            End If
            
            '获取下一个进程
            mresult = ProcessNext(mSnapShot, uProcess)
        Loop
    Else
        ErrMsgProc (msg)
    End If
    If lvwPrss.ListItems.Count = 0 Then
    lvwPrss.ListItems.Add , , "无"
    lvwPrss.ListItems(1).SubItems(1) = "未发现病毒进程，点击完成继续"
    CmdDelete.Enabled = False
    Else
    Dim x As Integer
    For x = 1 To lvwPrss.ListItems.Count
     lvwPrss.ListItems(x).Checked = True
    Next
    End If
End Sub
Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub
