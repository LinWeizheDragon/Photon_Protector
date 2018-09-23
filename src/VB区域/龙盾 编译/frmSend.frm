VERSION 5.00
Begin VB.Form frmSend 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   3600
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'**模 块 名：发送数据
'**说    明：根据窗体标题确定发送对象,对其发送字符串
'**创 建 人：LionKing1990
'**日    期：2010年3月19日
'**版    本：V1.0
'**备    注：'部分API被我动过手术了 , 跟API浏览器里面不太一样
'*************************************************************************

'查找窗口及进程
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'其实是PostMessage,不需要立即返回,'投递一条消息
Private Declare Function SendMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'读写进程
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'自定义的消息
Private Const WM_USER = &H400
Private Const Msg_GetAddress = WM_USER + 1
Private Const Msg_GetData = WM_USER + 2
Private Const Msg_AddressReady = WM_USER + 3
'HOOK,很方便
Private WithEvents Hook As cSubclass
Attribute Hook.VB_VarHelpID = -1

'消息处理
Private Sub Hook_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
If bBefore Then
    Select Case uMsg
    Case Msg_AddressReady
        '内存申请完毕,并得到地址,开始写入
        WriteData wParam, lParam
    Case Else
    End Select
End If
End Sub


'HOOK开始及结束
Private Sub Form_Load()
    Text1.Text = "作者:Lionking1990"
    Set Hook = New cSubclass
    Hook.AddWindowMsgs Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hook.DeleteWindowMsg Me.hwnd
End Sub

