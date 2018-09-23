Attribute VB_Name = "MDebug"
'原作者：Bruce Meckinney
Option Explicit



Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

#If iVBVer <= 5 Then
' This seems to have been left out of VB5, although it is documented
Public Enum LogModeConstant
    vbLogAuto
    vbLogOff
    vbLogToFile
    vbLogToNT
    vbLogOverwrite = &H10
    vbLogThreadID = &H20
End Enum
#End If


Public Const pNull = 0                '声明一个NULL指针

#Const afLogfile = 1
#Const afMsgBox = 2
#Const afDebugWin = 4
#Const afAppLog = 8         ' Log to file
#Const afAppLogNT = 16      ' NT event log
#Const afAppLogMask = 8 Or 16

'NOTE: 暂时在这儿定义调试方式，发行时须移入到命令行参数定义内
#Const afDebug = afMsgBox

Private secFreq As Currency
#If afDebug And afLogfile Then
Private iLogFile As Integer
#End If
#If afDebug And (afAppLog Or afAppLogNT) Then
Private fAppLog As Boolean
#End If

Function BugInit() As Boolean
    If secFreq = 0 Then BugInit = QueryPerformanceCounter(secFreq)
End Function

Sub BugTerm()
#If afDebug And afLogfile Then
    ' Close log file
    Close iLogFile
    iLogFile = 0
#End If
End Sub

' Display appropriate error message, and then stop
' program.  These errors should NOT be possible in
' shipping product.
Sub BugAssert(ByVal fExpression As Boolean, _
              Optional sExpression As String)
#If afDebug Then
    If fExpression Then Exit Sub
    BugMessage "BugAssert failed: " & sExpression
    Stop
#End If
End Sub
    
    
Sub BugMessage(sMsg As String)
#If afDebug And afLogfile Then
    If iLogFile = 0 Then
        iLogFile = FreeFile
        ' Warning: multiple instances can overwrite log file
        Open App.EXEName & ".DBG" For Output Shared As iLogFile
        ' Challenge: Rewrite to give each instance its own log file
    End If
    Print #iLogFile, sMsg
#End If
#If afDebug And afMsgBox Then
    MsgBox sMsg
#End If
#If afDebug And afDebugWin Then
    Debug.Print sMsg
#End If
#If afDebug And afAppLogMask Then
    If fAppLog = False Then
        fAppLog = True
#If (afDebug And afAppLogMask) = afAppLogNT Then
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogToNT
#ElseIf (afDebug And afAppLogMask) = afAppLog Then
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogToFile Or vbLogOverwrite
#Else
        App.StartLogging App.Path & "\" & App.EXEName & ".LOG", _
                         vbLogAuto Or vbLogOverwrite
#End If
    End If
    App.LogEvent sMsg
#End If
End Sub

Sub BugLocalMessage(sMsg As String)
#If fDebugLocal Then
    BugMessage sMsg
#End If
End Sub

Sub ProfileStart(secStart As Currency)
    If secFreq = 0 Then QueryPerformanceFrequency secFreq
    QueryPerformanceCounter secStart
End Sub

Sub ProfileStop(secStart As Currency, secTiming As Currency)
    QueryPerformanceCounter secTiming
    If secFreq = 0 Then
        secTiming = 0 ' Handle no high-resolution timer
    Else
        secTiming = (secTiming - secStart) / secFreq
    End If
End Sub

Sub ProfileStopMessage(sOutput As String, sPrefix As String, _
                       secStart As Currency, sPost As String)
#If afDebug Then
    Static secTiming As Currency
    QueryPerformanceCounter secTiming
    If secFreq = 0 Then
        secTiming = 0 ' Handle no high-resolution timer
    Else
        secTiming = (secTiming - secStart) / secFreq
    End If
    ' Return through parameter so that routine can be Sub
    sOutput = sPrefix & secTiming & sPost
#End If
End Sub

Sub BugProfileStop(sPrefix As String, secStart As Currency)
#If afDebug Then
    Static secTiming As Currency
    QueryPerformanceCounter secTiming
    If secFreq = 0 Then
        secTiming = 0 ' Handle no high-resolution timer
    Else
        secTiming = secTiming - secStart / secFreq
    End If
    BugMessage sPrefix & secTiming & " sec "
#End If
End Sub

Sub ApiRaise(ByVal e As Long)
    Err.Raise vbObjectError + e, _
              App.EXEName & ".Windows", ApiError(e)
End Sub

Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      pNull, e, 0&, s, Len(s), ByVal pNull)
    If c Then ApiError = Left$(s, c)
End Function

Function LastApiError() As String
    LastApiError = ApiError(Err.LastDllError)
End Function

Function BasicError(ByVal e As Long) As Long
    BasicError = e And &HFFFF&
End Function

Function COMError(e As Long) As Long
    COMError = e Or vbObjectError
End Function
'


