VERSION 5.00
Begin VB.Form frmHookCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2175
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Timer"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Text            =   "HookNtCreateProcessEx"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unload"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmHookCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Load_Drv As New cls_Driver

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal Dst As Long, ByVal Src As Long, ByVal uLen As Long)
Private Declare Sub GetMem4 Lib "msvbvm60.dll" (ByVal Address As Long, ByVal Dst As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ZwOpenProcess _
               Lib "ntdll.dll" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientId As CLIENT_ID) As Long


Const FILE_DEVICE_ROOTKIT As Long = &H2A7B
Const METHOD_BUFFERED     As Long = 0
Const METHOD_IN_DIRECT    As Long = 1
Const METHOD_OUT_DIRECT   As Long = 2
Const METHOD_NEITHER      As Long = 3
Const FILE_ANY_ACCESS     As Long = 0
Const FILE_READ_ACCESS    As Long = &H1     '// file & pipe
Const FILE_WRITE_ACCESS   As Long = &H2     '// file & pipe
Const FILE_READ_DATA      As Long = &H1
Const FILE_WRITE_DATA     As Long = &H2

Const TA_ALLOWCREATE      As Long = &H1
Const TA_UNALLOWCREATE    As Long = &H2
Const TA_LOOPING          As Long = &H1


Private Sub Command1_Click()
    
    Call Load_Drv.IoControl(Load_Drv.CTL_CODE_GEN(&H805), 0, 0, 0, 0)
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
 

    With Load_Drv
        .DelDrv
    End With
    Timer1.Enabled = False
    
End Sub

Private Sub Command3_Click()

Timer1.Enabled = Not Timer1.Enabled

End Sub




Private Sub Command4_Click()

    Call Load_Drv.IoControl(Load_Drv.CTL_CODE_GEN(&H806), 0, 0, 0, 0)
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    If EnablePrivilege(SE_DEBUG) = False Then
       If Not EnablePrivilege1(SE_DEBUG_PRIVILEGE, True) Then
          MsgBox "程序初始化失败。", 16, "错误"
          'End
       End If
    End If
    
    With Load_Drv
        .szDrvFilePath = App.Path & "\HookNtCreateProcessEx.sys"
        .szDrvLinkName = "HookNtCreateProcessEx"
        .szDrvSvcName = "HookNtCreateProcessEx"
        .szDrvDisplayName = "HookNtCreateProcessEx"
        .InstDrv
        .StartDrv
        .OpenDrv
        
    End With
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With Load_Drv
        .DelDrv
    End With
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
    Dim i As Long
    With Load_Drv
    Call .IoControl(.CTL_CODE_GEN(&H801), 0, 0, 0, 0, i)
    
    'MsgBox i
    Debug.Print i
    If i = TA_LOOPING Then
        
        Dim allow As Long
        Dim ProcessName As String
        Dim ProcessID As Long
        ProcessName = String$(260, 0)
        
        Call .IoControl(.CTL_CODE_GEN(&H802), 0, 0, VarPtr(ProcessID), 4)
        Call .IoControl(.CTL_CODE_GEN(&H803), 0, 0, StrPtr(ProcessName), 260)
        
        ProcessName = StrConv(ProcessName, vbUnicode)
        ProcessName = Left(ProcessName, InStr(1, ProcessName, Chr(0)) - 1)
         Timer1.Enabled = False
         If MsgBox("进程: " & ProcessID & vbCrLf & "尝试创建进程: " & ProcessName & vbCrLf & "是否允许？", vbOKCancel) = vbOK Then
           allow = TA_ALLOWCREATE
         Else
           allow = TA_UNALLOWCREATE
         End If

         Call .IoControl(.CTL_CODE_GEN(&H804), VarPtr(allow), 4, 0, 0)
         
         Dim hProcess As Long
         hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
         NtResumeProcess hProcess '继续
         ZwClose hProcess
         
         Sleep 1000
         
         hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
         NtResumeProcess hProcess '继续
         ZwClose hProcess
         Timer1.Enabled = True
    End If
    End With
    

End Sub
