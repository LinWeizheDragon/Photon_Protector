VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "病毒防御助手-病毒记录器"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   6540
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "描述"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "名称："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "文件："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRC32 As New clsGetCRC32

Public Function GetChecksum(sFile As String) As String
    On Error GoTo ErrHandle
    Dim cb0 As Byte
    Dim cb1 As Byte
    Dim cb2 As Byte
    Dim cb3 As Byte
    Dim cb4 As Byte
    Dim cb5 As Byte
    Dim cb6 As Byte
    Dim cb7 As Byte
    Dim cb8 As Byte
    Dim cb9 As Byte
    Dim cb10 As Byte
    Dim cb11 As Byte
    Dim cb12 As Byte
    Dim cb13 As Byte
    Dim cb14 As Byte
    Dim cb15 As Byte
    Dim cb16 As Byte
    Dim cb17 As Byte
    Dim cb18 As Byte
    Dim cb19 As Byte
    Dim cb20 As Byte
    Dim cb21 As Byte
    Dim cb22 As Byte
    Dim cb23 As Byte
    Dim buff As String
    Open sFile For Binary Access Read As #1
        buff = Space$(1)
        Get #1, , buff
    Close #1
    Open sFile For Binary Access Read As #2
        Get #2, 512, cb0
        Get #2, 1024, cb1
        Get #2, 2048, cb2
        Get #2, 3000, cb3
        Get #2, 4096, cb4
        Get #2, 5000, cb5
        Get #2, 6000, cb6
        Get #2, 7000, cb7
        Get #2, 8192, cb8
        Get #2, 9000, cb9
        Get #2, 10000, cb10
        Get #2, 11000, cb11
        Get #2, 12288, cb12
        Get #2, 13000, cb13
        Get #2, 14000, cb14
        Get #2, 15000, cb15
        Get #2, 16384, cb16
        Get #2, 17000, cb17
        Get #2, 18000, cb18
        Get #2, 19000, cb19
        Get #2, 20480, cb20
        Get #2, 21000, cb21
        Get #2, 22000, cb22
        Get #2, 23000, cb23
    Close #2
    buff = cb0
    buff = buff & cb1
    buff = buff & cb2
    buff = buff & cb3
    buff = buff & cb4
    buff = buff & cb5
    buff = buff & cb6
    buff = buff & cb7
    buff = buff & cb8
    buff = buff & cb9
    buff = buff & cb10
    buff = buff & cb11
    buff = buff & cb12
    buff = buff & cb13
    buff = buff & cb14
    buff = buff & cb15
    buff = buff & cb16
    buff = buff & cb17
    buff = buff & cb18
    buff = buff & cb19
    buff = buff & cb20
    buff = buff & cb21
    buff = buff & cb22
    buff = buff & cb23
    GetChecksum = CRC32.StringChecksum(buff)
    Set CRC32 = Nothing
    Exit Function
ErrHandle:
    Close #2
End Function

Function GetFullCRC(sFile As String) As String
    GetFullCRC = CRC32.FileChecksum(sFile)
End Function




Private Sub Command1_Click()
Dim str1, str2, str3 As String
If Text1.Text = "" Then MsgBox "请输入文件！", vbInformation, "错误": Exit Sub
If Text2.Text = "" Then MsgBox "请输入名称！", vbInformation, "错误": Exit Sub
str1 = Text1.Text
str2 = Text2.Text
str3 = Text3.Text
If Text3.Text = "" Then str3 = "未知描述"
str1 = GetChecksum(Text1.Text)
Text4.Text = str1 & "|" & str2 & "|" & str3
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    On Error GoTo err_filecopy
    '检查放下的东西是不是文件名，以及是否只拖放一个文件
    If Data.GetFormat(vbCFFiles) And Data.Files.Count = 1 Then
    Dim sFileName$
    '只读取第一条记录的信息
    sFileName = Data.Files(1)
     Text1.Text = sFileName
    End If
    Exit Sub
err_filecopy:
    MsgBox "文件拷贝出错：" & Err.Description

End Sub

Private Sub Text2_Change()
If UBound(Split(Text2.Text, "|")) <> 0 Then
MsgBox "不得带有：“|” 符号！"
Text2.Text = Replace(Text2.Text, "|", vbNullString)
End If
End Sub

Private Sub Text3_Change()
If UBound(Split(Text3.Text, "|")) <> 0 Then
MsgBox "不得带有：“|” 符号！"
Text3.Text = Replace(Text3.Text, "|", vbNullString)
End If
End Sub
