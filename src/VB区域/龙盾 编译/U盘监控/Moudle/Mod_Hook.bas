Attribute VB_Name = "Module1"
Public Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Public Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function ZwClose _
               Lib "ntdll.dll" (ByVal ObjectHandle As Long) As Long

Public Const OBJ_INHERIT = &H2
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SYNCHRONIZE As Long = &H100000
Public Const JOB_OBJECT_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1F
Public Const PROCESS_DUP_HANDLE As Long = &H40
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
