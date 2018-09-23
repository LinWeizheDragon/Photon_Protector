VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hook CreateProcessW"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3030
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3030
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   600
      Top             =   840
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   180
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   135
   End
   Begin VB.CommandButton cmdUnHook 
      Caption         =   "UnHook"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StartHook Lib "Hook.dll" () As Boolean
Private Declare Function UnLoadHook Lib "Hook.dll" () As Boolean

Private Sub cmdHook_Click()
    If StartHook Then cmdHook.Enabled = False
End Sub

Private Sub cmdUnHook_Click()
    If UnLoadHook Then cmdUnHook.Enabled = False: MsgBox "ok": Unload Me
End Sub



Private Sub Form_Initialize()
    SetButtonFlat cmdHook.hwnd
    SetButtonFlat cmdUnHook.hwnd
    Load frmRec
    If StartHook Then cmdHook.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Hide
    If App.PrevInstance Then
     MsgBox "已经开启进程保护！"
     Unload Me
     End
     End If
    c = GetWindowLong(hwnd, -4)
    SetWindowLong hwnd, -4, AddressOf Wndproc
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics



End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdUnHook.Enabled Then UnLoadHook: Unload Me
End Sub

Private Function SetButtonFlat(ByVal hwnd As Long) As Boolean
Dim style As Long
    style = GetWindowLong(hwnd, (-16))
    style = style Or &H8000&
    SetButtonFlat = SetWindowLong(hwnd, (-16), style)
End Function
