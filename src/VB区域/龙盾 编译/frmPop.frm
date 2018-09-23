VERSION 5.00
Begin VB.Form frmPop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   1200
   End
   Begin 工程1.jcbutton jcbutton1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "打开主界面"
   End
   Begin 工程1.jcbutton Exit 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "退出"
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  Private Declare Function ReleaseCapture Lib "user32" () As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Const HTCAPTION = 2
  Private Const WM_NCLBUTTONDOWN = &HA1
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
        'Retrieve information about the next process recorded in our system snapshot
        r = Process32Next(hSnapShot, uProcess)
    Loop
End Function

Private Sub Exit_Click(Index As Integer)
Unload frmMain
End Sub

Private Sub Form_Load()
Debug.Print Me.Left
Debug.Print Me.Top
Dim p As POINTAPI
    GetCursorPos p
Timer1.Enabled = True
Me.Move gFunPixelsToTwips(p.x, DIRECTION_HORIZONTAL) - Me.Width, gFunPixelsToTwips(p.y, DIRECTION_VERTICAL) - Me.Height
Debug.Print Me.Left
Debug.Print Me.Top
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub jcbutton1_Click(Index As Integer)
frmMain.Show
End Sub

Private Sub Timer1_Timer()
Dim i
Dim Top As Single
Top = Me.Top
For i = 1 To 100
Me.Move Me.Left, Top - i
SuperSleep 0.001
Next
Timer1.Enabled = False
End Sub
