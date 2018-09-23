Attribute VB_Name = "ShareMemMod"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Download by http://www.codefans.net
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long

Public Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long

Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long

Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Sub wcscpy Lib "NTDLL.DLL" (dest As Any, src As Any)
Public Declare Function strlen Lib "NTDLL.DLL" (dest As Any) As Long

Public Const FILE_MAP_ALL_ACCESS = 983071

Public Const PAGE_EXECUTE_READWRITE = 64

Type ShareMemStruct
    AntiHwnd As Long
    AntiPID As Long
    hHookID As Long
End Type
Public ShareMem As ShareMemStruct
Public pShareMem As Long
Public hFileMap As Long

Function MapMemFile() As Boolean

        hFileMap = OpenFileMapping(FILE_MAP_ALL_ACCESS, False, "shearememory535200701")

        If hFileMap = 0 Then
    
                hFileMap = CreateFileMapping(-1, ByVal 0, PAGE_EXECUTE_READWRITE, 0, 2048, "shearememory535200701")
        End If

        If hFileMap = 0 Then
    
                MapMemFile = False

                Exit Function
        
        End If

        pShareMem = MapViewOfFile(hFileMap, FILE_MAP_ALL_ACCESS, 0, 0, 0)

        If (pShareMem = 0) Then
                MapMemFile = False

                Exit Function

        End If
  
        MapMemFile = True

End Function

Function UnMapMemFile() As Boolean
        UnmapViewOfFile ShareMem
        CloseHandle hFileMap
End Function

Function SetData(ByRef data As ShareMemStruct)
    CopyMemory ByVal pShareMem, data, LenB(data)
End Function

Function GetData(ByRef data As ShareMemStruct)
    CopyMemory data, ByVal pShareMem, LenB(data)
End Function


