VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "强制删除文件 (支持拖放)"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5835
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出程序"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdKillFile 
      Caption         =   "开始删除"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdClearLst 
      Caption         =   "清除列表"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdShowPath 
      Caption         =   "添加目录"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdShowOpen 
      Caption         =   "添加文件"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.ListBox LstFile 
      Appearance      =   0  'Flat
      Height          =   3090
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4800
      Top             =   4200
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "注意：本工具将会直接强行删除目标文件，就算该文件正在使用中，所以请慎重删除！"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   5535
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194
Dim DrvController As New cls_Driver


Private Sub Form_Initialize()
'    SetButtonFlat CmdShowOpen.hWnd
'    SetButtonFlat CmdShowPath.hWnd
'    SetButtonFlat CmdClearLst.hWnd
'    SetButtonFlat CmdKillFile.hWnd
'    SetButtonFlat CmdAbout.hWnd
'    SetButtonFlat CmdExit.hWnd
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hWnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

End Sub

Private Sub Form_Load()
Dim a As Boolean, b As Boolean, c As Boolean
    With DrvController
        .szDrvFilePath = Replace(App.Path & "\SuperKillFile.sys", "\\", "\")
        .szDrvLinkName = "\\.\SuperKillFile"
        .szDrvDisplayName = "SuperKillFile"
        .szDrvSvcName = "SuperKillFile"
        .szDrvDeviceName = "\Device\SuperKillFile"
        a = .InstDrv
        b = .StartDrv
        c = .OpenDrv
    End With
    If a = False Or b = False Or c = False Then MsgBox "加载驱动失败,程序退出": End
'如果前一次运行程序没有通过Unload语句的话这里就无法通过，这个问题暂时无法解决，驱动的问题


End Sub

Private Sub Form_Unload(Cancel As Integer)
    With DrvController
        .StopDrv
        .DelDrv
    End With
    '卸载驱动，扫地工作
End Sub

Private Sub LstFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'OLE拖拽时，添加文件
On Error GoTo Err
Dim i As Long
    Me.Caption = "正在加载文件...请稍等"
    For i = 1 To Data.Files.Count
    '如果是文件夹则把文件夹里面的文件全部列出来
        If (GetAttr(Data.Files.Item(i)) And vbDirectory) = vbDirectory Then
            Findfile Data.Files.Item(i), LstFile
        Else
        '只添加文件
            LstFile.AddItem Data.Files.Item(i)
        End If
    Next
    addHorScrlBarListBox LstFile
    Me.Caption = "强制删除文件 (支持拖放)"
Err:
    'MsgBox "出错"
    Me.Caption = "强制删除文件 (支持拖放)"
End Sub

Private Sub CmdShowOpen_Click()
Dim Filter As String
Dim FilePath As String
'调用ShowOpen 自定义函数，类似于CommonDialog的自定义函数
    Filter = "所有可以用来删除的文件|*.*"
    Filter = Replace(Filter, "|", Chr(0))
    ShowOpen Me.hWnd, FilePath, , Filter, , , 0
    If FilePath = "" Then Exit Sub
    LstFile.AddItem FilePath
    addHorScrlBarListBox LstFile
End Sub

Private Sub CmdShowPath_Click()
Dim FilePath As String, i As Long
    ShowDir Me.hWnd, FilePath
    Me.Caption = "正在加载文件...请稍等"
    '加载目录，使用自定义函数ShowDir
    For i = 1 To intresult
        mydirectory(i) = ""
    Next
    Findfile FilePath, LstFile
    Me.Caption = "强制删除文件 (支持拖放)"
    addHorScrlBarListBox LstFile
End Sub

Private Sub CmdClearLst_Click()
    LstFile.Clear '清空
End Sub

Private Sub CmdKillFile_Click()
If LstFile.ListCount = 0 Then Exit Sub '如果没有东西就退出
'我移到前面，节省资源
Dim FileName As String, i As Long, k As Long
Dim ret As Byte
'声明变量
    Call VarPtr("VIRTUALIZER_START") '呼叫驱动
    For i = 0 To LstFile.ListCount - 1
        FileName = LstFile.List(i)
        SetAttr FileName, vbNormal '改变文件的属性为正常
        Me.Caption = "正在删除: " & FileName
        With DrvController '呼叫驱动，开始删除
            Call .IoControl(.CTL_CODE_GEN(&H360), VarPtr("\??\" & FileName), 4, ret, Len(ret))
        End With
    Next
    '呼叫驱动，删除结束
    Call VarPtr("VIRTUALIZER_END")
    For k = intresult - 1 To 1 Step -1
        Me.Caption = "删除目录: " & mydirectory(k)
        SetAttr mydirectory(k), vbNormal
        fRmdir mydirectory(k)
        '改变目录属性，然后删除之
    Next
    Me.Caption = "强制删除文件 (支持拖放)"
    LstFile.Clear
    MsgBox "完成!"
End Sub

Private Sub CmdAbout_Click()
    MsgBox "原作者：ilisuan  jupiter" & vbNewLine & "网上开源代码，经过修改后收录于病毒防御助手。"
End Sub

Private Sub CmdExit_Click()
'    With DrvController
'        .StopDrv
'        .DelDrv
'    End With
    Unload Me
    '呼叫删除驱动，有点多余。
End Sub


Private Sub addHorScrlBarListBox(ByVal refControlListBox As Object)
'为ListBox添加自动的ScrollBar滚动条
Dim nRet As Long, l As Long, lMax As Long, nNewWidth As Integer
    For l = 0 To LstFile.ListCount - 1
        If lMax < TextWidth(LstFile.List(l)) Then
            lMax = TextWidth(LstFile.List(l))
        End If
    Next l
    lMax = lMax + 120
    nNewWidth = lMax / 15
    nRet = SendMessage(refControlListBox.hWnd, LB_SETHORIZONTALEXTENT, nNewWidth, ByVal 0&)
End Sub

Private Function fRmdir(Path As String)
On Error GoTo Err
'消灭文件夹
    RmDir Path
Err:
End Function
