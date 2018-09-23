Attribute VB_Name = "modLockFileInfo"

Option Explicit
Private Declare Function NtQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ProcessInformationClass As PROCESSINFOCLASS, _
                                ByVal ProcessInformation As Long, _
                                ByVal ProcessInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
Private Enum PROCESSINFOCLASS
    ProcessBasicInformation = 0
    ProcessQuotaLimits
    ProcessIoCounters
    ProcessVmCounters
    ProcessTimes
    ProcessBasePriority
    ProcessRaisePriority
    ProcessDebugPort
    ProcessExceptionPort
    ProcessAccessToken
    ProcessLdtInformation
    ProcessLdtSize
    ProcessDefaultHardErrorMode
    ProcessIoPortHandlers
    ProcessPooledUsageAndLimits
    ProcessWorkingSetWatch
    ProcessUserModeIOPL
    ProcessEnableAlignmentFaultFixup
    ProcessPriorityClass
    ProcessWx86Information
    ProcessHandleCount
    ProcessAffinityMask
    ProcessPriorityBoost
    ProcessDeviceMap
    ProcessSessionInformation
    ProcessForegroundInformation
    ProcessWow64Information
    ProcessImageFileName
    ProcessLUIDDeviceMapsEnabled
    ProcessBreakOnTermination
    ProcessDebugObjectHandle
    ProcessDebugFlags
    ProcessHandleTracing
    ProcessIoPriority
    ProcessExecuteFlags
    ProcessResourceManagement
    ProcessCookie
    ProcessImageInformation
    MaxProcessInfoClass
End Enum
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long 'NTSTATUS
    PebBaseAddress As Long 'PPEB
    AffinityMask As Long 'ULONG_PTR
    BasePriority As Long 'KPRIORITY
    UniqueProcessId As Long 'ULONG_PTR
    InheritedFromUniqueProcessId As Long 'ULONG_PTR
End Type
Private Type FILE_NAME_INFORMATION
     FileNameLength As Long
     FileName(3) As Byte
End Type
Private Type NM_INFO
    Info As FILE_NAME_INFORMATION
    strName(259) As Byte
End Type
Private Enum FileInformationClass
    FileDirectoryInformation = 1
    FileFullDirectoryInformation = 2
    FileBothDirectoryInformation = 3
    FileBasicInformation = 4
    FileStandardInformation = 5
    FileInternalInformation = 6
    FileEaInformation = 7
    FileAccessInformation = 8
    FileNameInformation = 9
    FileRenameInformation = 10
    FileLinkInformation = 11
    FileNamesInformation = 12
    FileDispositionInformation = 13
    FilePositionInformation = 14
    FileFullEaInformation = 15
    FileModeInformation = 16
    FileAlignmentInformation = 17
    FileAllInformation = 18
    FileAllocationInformation = 19
    FileEndOfFileInformation = 20
    FileAlternateNameInformation = 21
    FileStreamInformation = 22
    FilePipeInformation = 23
    FilePipeLocalInformation = 24
    FilePipeRemoteInformation = 25
    FileMailslotQueryInformation = 26
    FileMailslotSetInformation = 27
    FileCompressionInformation = 28
    FileObjectIdInformation = 29
    FileCompletionInformation = 30
    FileMoveClusterInformation = 31
    FileQuotaInformation = 32
    FileReparsePointInformation = 33
    FileNetworkOpenInformation = 34
    FileAttributeTagInformation = 35
    FileTrackingInformation = 36
    FileMaximumInformation
End Enum
Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                                ByVal pSystemInformation As Long, _
                                ByVal SystemInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
                                
Private Enum SYSTEM_INFORMATION_CLASS
    SystemBasicInformation
    SystemProcessorInformation             '// obsolete...delete
    SystemPerformanceInformation
    SystemTimeOfDayInformation
    SystemPathInformation
    SystemProcessInformation
    SystemCallCountInformation
    SystemDeviceInformation
    SystemProcessorPerformanceInformation
    SystemFlagsInformation
    SystemCallTimeInformation
    SystemModuleInformation
    SystemLocksInformation
    SystemStackTraceInformation
    SystemPagedPoolInformation
    SystemNonPagedPoolInformation
    SystemHandleInformation
    SystemObjectInformation
    SystemPageFileInformation
    SystemVdmInstemulInformation
    SystemVdmBopInformation
    SystemFileCacheInformation
    SystemPoolTagInformation
    SystemInterruptInformation
    SystemDpcBehaviorInformation
    SystemFullMemoryInformation
    SystemLoadGdiDriverInformation
    SystemUnloadGdiDriverInformation
    SystemTimeAdjustmentInformation
    SystemSummaryMemoryInformation
    SystemMirrorMemoryInformation
    SystemPerformanceTraceInformation
    SystemObsolete0
    SystemExceptionInformation
    SystemCrashDumpStateInformation
    SystemKernelDebuggerInformation
    SystemContextSwitchInformation
    SystemRegistryQuotaInformation
    SystemExtendServiceTableInformation
    SystemPrioritySeperation
    SystemVerifierAddDriverInformation
    SystemVerifierRemoveDriverInformation
    SystemProcessorIdleInformation
    SystemLegacyDriverInformation
    SystemCurrentTimeZoneInformation
    SystemLookasideInformation
    SystemTimeSlipNotification
    SystemSessionCreate
    SystemSessionDetach
    SystemSessionInformation
    SystemRangeStartInformation
    SystemVerifierInformation
    SystemVerifierThunkExtend
    SystemSessionProcessInformation
    SystemLoadGdiDriverInSystemSpace
    SystemNumaProcessorMap
    SystemPrefetcherInformation
    SystemExtendedProcessInformation
    SystemRecommendedSharedDataAlignment
    SystemComPlusPackage
    SystemNumaAvailableMemory
    SystemProcessorPowerInformation
    SystemEmulationBasicInformation
    SystemEmulationProcessorInformation
    SystemExtendedHandleInformation
    SystemLostDelayedWriteInformation
    SystemBigPoolInformation
    SystemSessionPoolTagInformation
    SystemSessionMappedViewInformation
    SystemHotpatchInformation
    SystemObjectSecurityMode
    SystemWatchdogTimerHandler
    SystemWatchdogTimerInformation
    SystemLogicalProcessorInformation
    SystemWow64SharedInformation
    SystemRegisterFirmwareTableInformationHandler
    SystemFirmwareTableInformation
    SystemModuleInformationEx
    SystemVerifierTriageInformation
    SystemSuperfetchInformation
    SystemMemoryListInformation
    SystemFileCacheInformationEx
    MaxSystemInfoClass  '// MaxSystemInfoClass should always be the last enum
End Enum
Private Type SYSTEM_HANDLE
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    pObject As Long
    GrantedAccess As Long
End Type
Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Enum SYSTEM_HANDLE_TYPE
    OB_TYPE_UNKNOWN = 0
    OB_TYPE_TYPE = 1
    OB_TYPE_DIRECTORY
    OB_TYPE_SYMBOLIC_LINK
    OB_TYPE_TOKEN
    OB_TYPE_PROCESS
    OB_TYPE_THREAD
    OB_TYPE_UNKNOWN_7
    OB_TYPE_EVENT
    OB_TYPE_EVENT_PAIR
    OB_TYPE_MUTANT
    OB_TYPE_UNKNOWN_11
    OB_TYPE_SEMAPHORE
    OB_TYPE_TIMER
    OB_TYPE_PROFILE
    OB_TYPE_WINDOW_STATION
    OB_TYPE_DESKTOP
    OB_TYPE_SECTION
    OB_TYPE_KEY
    OB_TYPE_PORT
    OB_TYPE_WAITABLE_PORT
    OB_TYPE_UNKNOWN_21
    OB_TYPE_UNKNOWN_22
    OB_TYPE_UNKNOWN_23
    OB_TYPE_UNKNOWN_24
    OB_TYPE_IO_COMPLETION
    OB_TYPE_FILE
End Enum
'typedef struct _SYSTEM_HANDLE_INFORMATION
'{
'   ULONG           uCount;
'   SYSTEM_HANDLE   aSH[];
'} SYSTEM_HANDLE_INFORMATION, *PSYSTEM_HANDLE_INFORMATION;
Private Type SYSTEM_HANDLE_INFORMATION
    uCount As Long
    aSH() As SYSTEM_HANDLE
End Type
Private Declare Function NtDuplicateObject Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, _
                                ByVal SourceHandle As Long, _
                                ByVal TargetProcessHandle As Long, _
                                ByRef TargetHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByVal HandleAttributes As Long, _
                                ByVal Options As Long) As Long
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const DUPLICATE_SAME_ATTRIBUTES = &H4
Private Declare Function NtOpenProcess Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientID As CLIENT_ID) As Long
Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type
Private Type CLIENT_ID
    UniqueProcess As Long
    UniqueThread  As Long
End Type
Private Type IO_STATUS_BLOCK
    Status As Long
    uInformation As Long
End Type
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const PROCESS_DUP_HANDLE As Long = (&H40)
Private Declare Function NtClose Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal Length As Long)
                                      
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'typedef struct _OBJECT_NAME_INFORMATION
'{
'    UNICODE_STRING  Name;
'} OBJECT_NAME_INFORMATION, *POBJECT_NAME_INFORMATION;
'typedef enum _OBJECT_INFORMATION_CLASS
'{
'    ObjectBasicInformation,             // 0    Y       N
'    ObjectNameInformation,              // 1    Y       N
'    ObjectTypeInformation,              // 2    Y       N
'    ObjectAllTypesInformation,          // 3    Y       N
'    ObjectHandleInformation             // 4    Y       Y
'} OBJECT_INFORMATION_CLASS;
Private Enum OBJECT_INFORMATION_CLASS
    ObjectBasicInformation = 0
    ObjectNameInformation
    ObjectTypeInformation
    ObjectAllTypesInformation
    ObjectHandleInformation
End Enum
'
'typedef struct _UNICODE_STRING
'{
'    USHORT Length;
'    USHORT MaximumLength;
'    PWSTR Buffer;
'} UNICODE_STRING, *PUNICODE_STRING;
Private Type UNICODE_STRING
    uLength As Integer
    uMaximumLength As Integer
    pBuffer(3) As Byte
End Type
Private Type OBJECT_NAME_INFORMATION
    pName As UNICODE_STRING
End Type
Private Const STATUS_INFO_LEN_MISMATCH = &HC0000004
Private Const HEAP_ZERO_MEMORY = &H8
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function NtQueryObject Lib "NTDLL.DLL" (ByVal ObjectHandle As Long, _
'                                                        ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, _
'                                                        ObjectInformation As Any, ByVal ObjectInformationLength As Long, _
'                                                        ReturnLength As Long) As Long
Private Declare Function NtQueryObject Lib "NTDLL.DLL" (ByVal ObjectHandle As Long, _
                                                        ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, _
                                                        ByVal ObjectInformation As Long, ByVal ObjectInformationLength As Long, _
                                                        ReturnLength As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Function NT_SUCCESS(ByVal nStatus As Long) As Boolean
    NT_SUCCESS = (nStatus >= 0)
End Function
Public Function GetFileFullPath(ByVal hFile As Long) As String
    Dim hHeap As Long, dwSize As Long, objName As UNICODE_STRING, pName As Long
    Dim ntStatus As Long, i As Long, lngNameSize As Long, strDrives As String, strArray() As String
    Dim dwDriversSize As Long, strDrive As String, strTmp As String, strTemp As String
    On Error GoTo ErrHandle
    hHeap = GetProcessHeap
    pName = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, &H1000)
    ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000, dwSize)
    If (NT_SUCCESS(ntStatus)) Then
        i = 1
        Do While (ntStatus = STATUS_INFO_LEN_MISMATCH)
            pName = HeapReAlloc(hHeap, HEAP_ZERO_MEMORY, pName, &H1000 * i)
            ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000, ByVal 0)
            i = i + 1
        Loop
    End If
    HeapFree hHeap, 0, pName
    strTemp = String(512, Chr(0))
    lstrcpyW strTemp, pName + Len(objName)
    strTemp = StrConv(strTemp, vbFromUnicode)
    strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    strDrives = String(512, Chr(9))
    dwDriversSize = GetLogicalDriveStrings(512, strDrives)
    If dwDriversSize Then
        strArray = Split(strDrives, Chr(0))
        For i = 0 To UBound(strArray)
            If strArray(i) <> "" Then
                strDrive = Left(strArray(i), 2)
                strTmp = String(260, Chr(0))
                Call QueryDosDevice(strDrive, strTmp, 256)
                strTmp = Left(strTmp, InStr(strTmp, Chr(0)) - 1)
                If InStr(LCase(strTemp), LCase(strTmp)) = 1 Then
                    GetFileFullPath = strDrive & Mid(strTemp, Len(strTmp) + 1, Len(strTemp) - Len(strTmp))
                    Exit Function
                End If
            End If
        Next
    End If
ErrHandle:
End Function
Public Function CloseLockFileHandle(ByVal strFileName As String, ByVal dwProcessId As Long) As Boolean
    Dim ntStatus As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim lngHandles As Long
    Dim i As Long
    Dim objInfo As SYSTEM_HANDLE_INFORMATION, lngType As Long
    Dim hProcess As Long, hProcessToDup As Long, hFileHandle As Long
    Dim hFile As Long
    'Dim objIo As IO_STATUS_BLOCK, objFn As FILE_NAME_INFORMATION, objN As NM_INFO
    Dim bytBytes() As Byte, strSubPath As String, strTmp As String
    Dim blnIsOk As Boolean
    strSubPath = Mid(strFileName, 3, Len(strFileName) - 2)
    hFile = CreateFile("NUL", &H80000000, 0, ByVal 0&, 3, 0, 0)
    If hFile = -1 Then
        CloseLockFileHandle = False
        Exit Function
    End If
    objOa.Length = Len(objOa)
    objCid.UniqueProcess = dwProcessId
    ntStatus = 0
    Dim bytBuf() As Byte
    Dim nSize As Long
    nSize = 1
    Do
        ReDim bytBuf(nSize)
        ntStatus = NtQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), nSize, 0&)
        If (Not NT_SUCCESS(ntStatus)) Then
            If (ntStatus <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        nSize = nSize * 2
        ReDim bytBuf(nSize)
    Loop
    lngHandles = 0
    CopyMemory objInfo.uCount, bytBuf(0), 4
    lngHandles = objInfo.uCount
    ReDim objInfo.aSH(lngHandles - 1)
    Call CopyMemory(objInfo.aSH(0), bytBuf(4), Len(objInfo.aSH(0)) * lngHandles)
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).HandleValue = hFile And objInfo.aSH(i).UniqueProcessId = GetCurrentProcessId Then
            lngType = objInfo.aSH(i).ObjectTypeIndex
            Exit For
        End If
    Next
    NtClose hFile
    blnIsOk = True
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).ObjectTypeIndex = lngType And objInfo.aSH(i).UniqueProcessId = dwProcessId Then
            ntStatus = NtOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, objOa, objCid)
            If hProcessToDup <> 0 Then
                ntStatus = NtDuplicateObject(hProcessToDup, objInfo.aSH(i).HandleValue, GetCurrentProcess, hFileHandle, 0, 0, DUPLICATE_SAME_ATTRIBUTES)
                If (NT_SUCCESS(ntStatus)) Then
                    '这里如果直接调用NtQueryObject可能会挂起解决方法是用线程去处理当线程处理时间超过一定时间就把它干掉
                    '由于VB对多线程支持很差，其实应该说是对CreateThread支持很差，什么原因不要问我，相信网上也写有不少
                    '文件是关于它的,这里我选择了另一个函数也可以建立线程但是它是建立远程线程的，不过它却很稳定正好解决了
                    '我们这里的问题它就是CreateRemoteThread,^_^还记得我说过它很强大吧~~哈哈。
                    ntStatus = MyGetFileType(hFileHandle)
                    If ntStatus Then
                        strTmp = GetFileFullPath(hFileHandle)
                    End If
                    NtClose hFileHandle
                    If InStr(LCase(strTmp), LCase(strFileName)) Then
                        If Not CloseRemoteHandle(dwProcessId, objInfo.aSH(i).HandleValue, strFileName) Then
                            blnIsOk = False
                        End If
                    End If
                End If
            End If
        End If
    Next
    CloseLockFileHandle = blnIsOk
End Function
'检测所有进程
Public Function CloseLoackFiles(ByVal strFileName As String) As Boolean
    Dim ntStatus As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim lngHandles As Long
    Dim i As Long
    Dim objInfo As SYSTEM_HANDLE_INFORMATION, lngType As Long
    Dim hProcess As Long, hProcessToDup As Long, hFileHandle As Long
    Dim hFile As Long, blnIsOk As Boolean, strProcessName As String
    'Dim objIo As IO_STATUS_BLOCK, objFn As FILE_NAME_INFORMATION, objN As NM_INFO
    Dim bytBytes() As Byte, strSubPath As String, strTmp As String
    strSubPath = Mid(strFileName, 3, Len(strFileName) - 2)
    hFile = CreateFile("NUL", &H80000000, 0, ByVal 0&, 3, 0, 0)
    If hFile = -1 Then
        CloseLoackFiles = False
        Exit Function
    End If
    objOa.Length = Len(objOa)
    ntStatus = 0
    Dim bytBuf() As Byte
    Dim nSize As Long
    nSize = 1
    Do
        ReDim bytBuf(nSize)
        ntStatus = NtQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), nSize, 0&)
        If (Not NT_SUCCESS(ntStatus)) Then
            If (ntStatus <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        nSize = nSize * 2
        ReDim bytBuf(nSize)
    Loop
    lngHandles = 0
    CopyMemory objInfo.uCount, bytBuf(0), 4
    lngHandles = objInfo.uCount
    ReDim objInfo.aSH(lngHandles - 1)
    Call CopyMemory(objInfo.aSH(0), bytBuf(4), Len(objInfo.aSH(0)) * lngHandles)
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).HandleValue = hFile And objInfo.aSH(i).UniqueProcessId = GetCurrentProcessId Then
            lngType = objInfo.aSH(i).ObjectTypeIndex
            Exit For
        End If
    Next
    NtClose hFile
    blnIsOk = True
    For i = 0 To lngHandles - 1
        If objInfo.aSH(i).ObjectTypeIndex = lngType Then
            objCid.UniqueProcess = objInfo.aSH(i).UniqueProcessId
            ntStatus = NtOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, objOa, objCid)
            If hProcessToDup <> 0 Then
                ntStatus = NtDuplicateObject(hProcessToDup, objInfo.aSH(i).HandleValue, GetCurrentProcess, hFileHandle, 0, 0, DUPLICATE_SAME_ATTRIBUTES)
                If (NT_SUCCESS(ntStatus)) Then
                    '这里如果直接调用NtQueryObject可能会挂起解决方法是用线程去处理当线程处理时间超过一定时间就把它干掉
                    '由于VB对多线程支持很差，其实应该说是对CreateThread支持很差，什么原因不要问我，相信网上也写有不少
                    '文件是关于它的,这里我选择了另一个函数也可以建立线程但是它是建立远程线程的，不过它却很稳定正好解决了
                    '我们这里的问题它就是CreateRemoteThread,^_^还记得我说过它很强大吧~~哈哈。
                    ntStatus = MyGetFileType(hFileHandle)
                    If ntStatus Then
                        strTmp = GetFileFullPath(hFileHandle)
                    Else
                        strTmp = ""
                    End If
                    NtClose hFileHandle
                    If InStr(LCase(strTmp), LCase(strFileName)) Then
                        If Not CloseRemoteHandle(objInfo.aSH(i).UniqueProcessId, objInfo.aSH(i).HandleValue, strTmp) Then
                            blnIsOk = False
                        End If
                    End If
                End If
            End If
        End If
    Next
    CloseLoackFiles = blnIsOk
End Function
Private Function GetProcessCommandLine(ByVal dwProcessId As Long) As String
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim ntStatus As Long, hKernel As Long, strName As String
    Dim hProcess As Long, dwAddr As Long, dwRead As Long
    objOa.Length = Len(objOa)
    objCid.UniqueProcess = dwProcessId
    ntStatus = NtOpenProcess(hProcess, &H10, objOa, objCid)
    If hProcess = 0 Then
        GetProcessCommandLine = ""
        Exit Function
    End If
    hKernel = GetModuleHandle("kernel32")
    dwAddr = GetProcAddress(hKernel, "GetCommandLineA")
    CopyMemory dwAddr, ByVal dwAddr + 1, 4
    If ReadProcessMemory(hProcess, ByVal dwAddr, dwAddr, 4, dwRead) Then
        strName = String(260, Chr(0))
        If ReadProcessMemory(hProcess, ByVal dwAddr, ByVal strName, 260, dwRead) Then
            strName = Left(strName, InStr(strName, Chr(0)) - 1)
            NtClose hProcess
            GetProcessCommandLine = strName
            Exit Function
        End If
    End If
    NtClose hProcess
End Function
'解锁指定进程的锁定文件
Public Function CloseRemoteHandle(ByVal dwProcessId, ByVal hHandle As Long, Optional ByVal strLockFile As String = "") As Boolean
    Dim hMyProcess  As Long, hRemProcess As Long, blnResult As Long, hMyHandle As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES
    Dim ntStatus As Long, strProcessName As String, hProcess As Long, strMsg As String
    objCid.UniqueProcess = dwProcessId
    objOa.Length = Len(objOa)
    hMyProcess = GetCurrentProcess()
    ntStatus = NtOpenProcess(hRemProcess, PROCESS_DUP_HANDLE, objOa, objCid)
    If hRemProcess Then
        ntStatus = NtDuplicateObject(hRemProcess, hHandle, GetCurrentProcess, hMyHandle, 0, 0, DUPLICATE_CLOSE_SOURCE Or DUPLICATE_SAME_ACCESS)
        If (NT_SUCCESS(ntStatus)) Then
        'If DuplicateHandle(hRemProcess, hMyProcess, hHandle, hMyHandle, 0, 0, DUPLICATE_CLOSE_SOURCE Or DUPLICATE_SAME_ACCESS) Then
            blnResult = NtClose(hMyHandle)
            If blnResult >= 0 Then
                strProcessName = GetProcessCommandLine(dwProcessId)
                'If InStr(LCase(strProcessName), LCase(strLockFile)) Then
                If InStr(LCase(strProcessName), "explorer.exe") = 0 And dwProcessId <> GetCurrentProcessId Then
                    objCid.UniqueProcess = dwProcessId
                    ntStatus = NtOpenProcess(hProcess, 1, objOa, objCid)
                    If hProcess <> 0 Then TerminateProcess hProcess, 0
                End If
            End If
        End If
        Call NtClose(hRemProcess)
    End If
    CloseRemoteHandle = blnResult >= 0
End Function

'解锁指定进程的锁定文件
Public Function CloseRemoteHandleEx(ByVal dwProcessId, ByVal hHandle As Long, Optional ByVal strLockFile As String = "") As Boolean
    Dim hRemProcess As Long, hThread As Long, lngResult As Long, pfnThreadRtn As Long, hKernel As Long
    Dim objCid As CLIENT_ID
    Dim objOa As OBJECT_ATTRIBUTES, strMsg As String
    Dim ntStatus As Long, strProcessName As String, hProcess As Long
    objCid.UniqueProcess = dwProcessId
    objOa.Length = Len(objOa)
    ntStatus = NtOpenProcess(hRemProcess, PROCESS_QUERY_INFORMATION Or PROCESS_CREATE_THREAD Or PROCESS_VM_OPERATION Or PROCESS_VM_WRITE, objOa, objCid)
'    hMyProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_CREATE_THREAD Or PROCESS_VM_OPERATION Or PROCESS_VM_WRITE, 0, dwProcessId)
    If hRemProcess = 0 Then
        CloseRemoteHandleEx = False
        Exit Function
    End If
    hKernel = GetModuleHandle("kernel32")
    If hKernel = 0 Then
        CloseRemoteHandleEx = False
        Exit Function
    End If
    pfnThreadRtn = GetProcAddress(hKernel, "CloseHandle")
    If pfnThreadRtn = 0 Then
        FreeLibrary hKernel
        CloseRemoteHandleEx = False
        Exit Function
    End If
    hThread = CreateRemoteThread(hRemProcess, ByVal 0&, 0&, ByVal pfnThreadRtn, ByVal hHandle, 0, 0&)
    If hThread = 0 Then
        FreeLibrary hKernel
        CloseRemoteHandleEx = False
        Exit Function
    End If
    GetExitCodeThread hThread, lngResult
    CloseRemoteHandleEx = CBool(lngResult)
    strProcessName = GetProcessCommandLine(dwProcessId)
    If InStr(strProcessName, strLockFile) Then
        objCid.UniqueProcess = dwProcessId
        ntStatus = NtOpenProcess(hProcess, 1, objOa, objCid)
        If hProcess <> 0 Then TerminateProcess hProcess, 0
    End If
    NtClose hThread
    NtClose hRemProcess
    FreeLibrary hKernel
End Function
Private Function MyGetFileType(ByVal hFile As Long) As Long
    Dim hRemProcess As Long, hThread As Long, lngResult As Long, pfnThreadRtn As Long, hKernel As Long
    Dim dwEax As Long, dwTimeOut As Long
    hRemProcess = GetCurrentProcess
    hKernel = GetModuleHandle("kernel32")
    If hKernel = 0 Then
        MyGetFileType = 0
        Exit Function
    End If
    pfnThreadRtn = GetProcAddress(hKernel, "GetFileType")
    If pfnThreadRtn = 0 Then
        FreeLibrary hKernel
        MyGetFileType = 0
        Exit Function
    End If
    hThread = CreateRemoteThread(hRemProcess, ByVal 0&, 0&, ByVal pfnThreadRtn, ByVal hFile, 0, ByVal 0&)
    dwEax = WaitForSingleObject(hThread, 100)
    If dwEax = &H102 Then
        Call GetExitCodeThread(hThread, dwTimeOut)
        Call TerminateThread(hThread, dwTimeOut)
        NtClose hThread
        MyGetFileType = 0
        Exit Function
    End If
    If hThread = 0 Then
        FreeLibrary hKernel
        MyGetFileType = False
        Exit Function
    End If
    GetExitCodeThread hThread, lngResult
    MyGetFileType = lngResult
    NtClose hThread
    NtClose hRemProcess
    FreeLibrary hKernel
End Function

