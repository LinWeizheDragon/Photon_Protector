Attribute VB_Name = "modFileMap"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShowStyle
    vbHide
    vbMaximizedFocus
    vbMinimizedFocus
    vbMinimizedNoFocus
    vbNormalFocus
    vbNormalNoFocus
End Enum
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" _
    (ByVal hFile As Long, _
    ByVal lpFileMappingAttributes As Long, _
    ByVal flProtect As Long, _
    ByVal dwMaximumSizeHigh As Long, _
    ByVal dwMaximumSizeLow As Long, _
    ByVal lpName As String) As Long
Public Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" _
    (ByVal hFileMappingObject As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwFileOffsetHigh As Long, _
    ByVal dwFileOffsetLow As Long, _
    ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" _
    (lpBaseAddress As Any) As Long
Public Const ERROR_ALREADY_EXISTS = 183&

Public Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" _
    (DesStr As Any, _
    SrcStr As Any, _
    ByVal MaxLen As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, _
    ByVal cbCopy As Long)

    
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
    
'Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const FILE_MAP_WRITE = &H2
Public Const FILE_MAP_READ = &H4
Public Const PAGE_READWRITE = 4&
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const CREATE_ALWAYS = 2
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public hMemShare As Long 'Ó³ÉäÎÄ¼þ¾ä±ú
Public lShareData As Long 'Ó³ÉäµØÖ·

Public Function SuperSleep(DealyTime As Single) '´Ë´¦Ô­Îªlong£¬ÐÞ¸ÄÎªsingle¿ÉÑÓÊ±1ms :SK<2<8h
Dim TimerCount As Single
    TimerCount = Timer + DealyTime 'Ôö¼ÓXÃë ZJ9x6|q
    While TimerCount - Timer > 0
        DoEvents
        Sleep 1
    Wend
End Function
