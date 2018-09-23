Attribute VB_Name = "mdChangeNum"
Public strShare As New CSharedString

'User
Public Enum UserMsgStyle
    UserYes
    UserRetry
    UserYesNo
    UserOK
    UserOKOnly
End Enum
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

Public Function exitproc(ByVal exefile As String) As Boolean
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
        'Retrieve information about the next process recorded in our system snapshot
        r = Process32Next(hSnapShot, uProcess)
    Loop
End Function

Public Function ChangeNum()
On Error GoTo Err:
Dim DTime
DTime = ReadString("Log", "Time", App.Path & "\Set.ini")
If Len(DTime) < 7 Then
Do Until Len(DTime) >= 7
 DTime = "0|" & DTime
Loop
End If
Dim CTime As Integer
Debug.Print DTime
CTime = Split(DTime, "|")(0)
frmMain.Tm1.Picture = frmMain.ListNum(Val(CTime)).Picture
CTime = Split(DTime, "|")(1)
frmMain.Tm2.Picture = frmMain.ListNum(Val(CTime)).Picture
CTime = Split(DTime, "|")(2)
frmMain.Tm3.Picture = frmMain.ListNum(Val(CTime)).Picture
CTime = Split(DTime, "|")(3)
frmMain.Tm4.Picture = frmMain.ListNum(Val(CTime)).Picture
Exit Function

Err:
frmMain.Tm1.Picture = frmMain.ListNum(0).Picture
frmMain.Tm2.Picture = frmMain.ListNum(1).Picture
frmMain.Tm3.Picture = frmMain.ListNum(2).Picture
frmMain.Tm4.Picture = frmMain.ListNum(3).Picture
End Function

Public Function MsgUserBox(ByVal Text As String, Optional ByVal MsgType As UserMsgStyle = UserOK) As UserMsgStyle
Dim MyMsgform As New frmMsg
Select Case MsgType
Case UserOK
frmMsg.Show
frmMsg.Label1.Caption = Text
Do Until MyMsgform Is Nothing
SuperSleep 1
Loop
MsgUserBox = UserOK
End Select

End Function

