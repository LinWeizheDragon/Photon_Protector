VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form MYRECEIVER 
   Caption         =   "MYRECEIVER"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   4650
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   960
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开启"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "关闭"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   360
      Top             =   960
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "MYRECEIVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const TH32CS_SNAPHEAPLIST = &H1

Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
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
szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long



Function exitproc(ByVal exefile As String) As Boolean
Dim r
exitproc = False
Dim hSnapShot As Long, uProcess As PROCESSENTRY32
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
uProcess.dwSize = Len(uProcess)
r = Process32First(hSnapShot, uProcess)
Do While r
If Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)) = exefile Then
exitproc = True
Exit Do
End If
r = Process32Next(hSnapShot, uProcess)
Loop
End Function

Private Sub Command1_Click()
Call WriteString("RTA", "status", 1, App.Path & "\Chat.ini")
Shell App.Path & "\PRMonitor.exe"
End Sub



Private Sub Command2_Click()
MsgBox Form_Map.ReadMem
End Sub

Private Sub Command3_Click()
Call WriteString("RTA", "status", 0, App.Path & "\Chat.ini")
End Sub



Private Sub Form_Load()

If App.PrevInstance Then
Unload Me
End If
App.TaskVisible = False
'-------------皮肤控件加载----------------
Dim FileName As String
Dim IniFile As String
FileName = App.Path & "\Skin\Office2007.cjstyles"
IniFile = "NormalBlue.ini"
SkinFramework1.LoadSkin FileName, IniFile
SkinFramework1.ApplyWindow Me.hwnd
SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

Call WriteString("Rules", """" & App.Path & "\ProcessRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\RegRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\USBRTA.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\ProgramUpdate.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\DragonShield.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\PRMonitor.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\Tools\FixFolders\龙盾-移动盘隐藏文件夹修复工具.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\Tools\KillFiles\KillFile.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\Tools\ProcessMonitor\ProcessMonitor.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\Tools\RegMonitor\RegTools.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\PhotonRepair.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\PhotonClear.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\PhotonMajorization.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\ProtectProcess.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\Protect.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\NetScanner\Photon-NetScanner.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\NetScanner\koemsec1.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\NetScanner\kdumprep.exe""", "1", App.Path & "\Rules.ini")
Call WriteString("Rules", """" & App.Path & "\NetScanner\kdumpfix.exe""", "1", App.Path & "\Rules.ini")






c = GetWindowLong(Me.hwnd, -4)
   SetWindowLong Me.hwnd, -4, AddressOf Wndproc

Shell App.Path & "\PRMonitor.exe"
Timer1.Enabled = True
Load frmRec
Me.Hide
End Sub

Private Sub Form_Resize()
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i
Call Form_Map.WriteMem("ProcessRTA.Close")
SuperSleep 1
For Each i In Forms
Unload i
Next
End Sub

Private Sub Timer1_Timer()
If exitproc("PRMonitor.exe") = False Then
Unload Me
End If
End Sub

