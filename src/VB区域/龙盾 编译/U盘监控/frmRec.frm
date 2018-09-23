VERSION 5.00
Begin VB.Form frmRec 
   Caption         =   "消息接受窗体"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   3630
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   1200
   End
End
Attribute VB_Name = "frmRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LenStr As Long = 65535 * 10
Dim strShare As String
Private Sub Form_Load()
hMemShare = CreateFileMapping(&HFFFFFFFF, _
        0, _
        PAGE_READWRITE, _
         0, _
        LenStr, _
        "PhotonMemorySpace")
    If hMemShare = 0 Then
        'Err.LastDllError
        MsgBox "创建内存映射文件失败!", vbCritical, "错误"
    End If
    If (err.LastDllError = ERROR_ALREADY_EXISTS) Then
        '指定内存文件已存在
    End If
    
    lShareData = MapViewOfFile(hMemShare, FILE_MAP_WRITE, 0, 0, 0)
    If lShareData = 0 Then
        MsgBox "为映射文件对象创建视失败!", vbCritical, "错误"
    End If
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lShareData <> 0 Then
        Call UnmapViewOfFile(ByVal lShareData) '解除映射
        lShareData = 0
    End If
    
    If hMemShare <> 0 Then
        Call CloseHandle(hMemShare) '关闭映射
        hMemShare = 0
    End If
End Sub
Public Function ReadMem() As String
    sData = String(LenStr, vbNullChar)
    Call lstrcpyn(ByVal sData, ByVal lShareData, LenStr)
    ReadMem = sData
End Function

Private Sub WriteMem(ByVal text)
    sData = text
    Call lstrcpyn(ByVal lShareData, ByVal sData, LenStr)
End Sub

Private Sub Timer1_Timer()
Text2.text = ReadMem()
strShare = Text2.text
If strShare = Text1.text Then Exit Sub
Text1.text = strShare
DoCommand Text1.text
End Sub

Private Sub DoCommand(ByVal text As String)
If text = "USBRTA.Unload" Then
 Unload frmMain
End If

End Sub
