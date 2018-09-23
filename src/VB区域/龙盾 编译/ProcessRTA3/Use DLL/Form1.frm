VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "全局 Hook CreateProcessInternalW 测试"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4770
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "开始Hook"
      Height          =   360
      Left            =   1710
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3600
      Top             =   360
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Private Declare Sub FreeHook Lib "exports.dll" ()
Private Declare Sub EnableHook Lib "exports.dll" ()
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Sub cmdCommand1_Click()

        
        Dim tmp As ShareMemStruct

        EnableHook
        cmdCommand1.Enabled = False
End Sub

Private Sub cmdCommand2_Click()

End Sub

Private Sub Command1_Click()
MsgBox CheckProcess("C:\Windows\System32\explorer.exe", "G:\360data\重要数据\桌面\FlashFXP_xp911.com\FlashFXP_4.2.2.1760-Special\FlashFXP.exe")
End Sub

Private Sub Form_Load()
        ExeFiles = Space$(1000)
        CommandLines = Space$(1000)
        MapMemFile
        
        ShareMem.AntiHwnd = Me.hwnd
        SetData ShareMem
        
        oldWNDPROC = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
        Dim tmp As ShareMemStruct

        EnableHook
        cmdCommand1.Enabled = False
        Me.Hide
        Load frmRec
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnMapMemFile
For Each i In Forms
Unload i
Next
End Sub


