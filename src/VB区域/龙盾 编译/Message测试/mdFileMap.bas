Attribute VB_Name = "mdFileMap"
Option Explicit
'Download by http://www.codefans.net

Public sData As String
Public Const LenStr As Long = 65535 * 10
Public Declare Function CreateFileMapping Lib "KERNEL32" Alias "CreateFileMappingA" _
    (ByVal hFile As Long, _
    ByVal lpFileMappingAttributes As Long, _
    ByVal flProtect As Long, _
    ByVal dwMaximumSizeHigh As Long, _
    ByVal dwMaximumSizeLow As Long, _
    ByVal lpName As String) As Long
Public Declare Function OpenFileMapping Lib "KERNEL32" Alias "OpenFileMappingA" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "KERNEL32" _
    (ByVal hFileMappingObject As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwFileOffsetHigh As Long, _
    ByVal dwFileOffsetLow As Long, _
    ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "KERNEL32" _
    (lpBaseAddress As Any) As Long
Public Const ERROR_ALREADY_EXISTS = 183&

Public Declare Function lstrcpyn Lib "KERNEL32" Alias "lstrcpynA" _
    (DesStr As Any, _
    SrcStr As Any, _
    ByVal MaxLen As Long) As Long
Public Declare Sub RtlMoveMemory Lib "KERNEL32" (lpvDest As Any, lpvSource As Any, _
    ByVal cbCopy As Long)

    
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Public Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" _
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



