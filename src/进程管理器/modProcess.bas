Attribute VB_Name = "modProcess"
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long

Public Const PROCESS_ALL_ACCESS = 0
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_VM_READ = 16
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_SET_INFORMATION = 612

Public Const STILL_ACTIVE = &H103

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Const MAX_PATH As Integer = 260

'''''''''''''''''''本部分用于获取和设置进程优先级
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
'''''''''''''''''''

Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Type PROCESSENTRY32
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



'原modProcess.bas

'API 常量 数据类型的声明和一些公共函数。
Public Declare Function AdjustTokenPrivileges _
                Lib "advapi32.dll" (ByVal TokenHandle As Long, _
                                    ByVal DisableAllPriv As Long, _
                                    ByRef NewState As TOKEN_PRIVILEGES, _
                                    ByVal BufferLength As Long, _
                                    ByRef PreviousState As TOKEN_PRIVILEGES, _
                                    ByRef pReturnLength As Long) As Long
Public Declare Function GetCurrentProcess _
                Lib "kernel32" () As Long
Public Declare Function LookupPrivilegeValue _
                Lib "advapi32.dll" _
                Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, _
                                               ByVal lpName As String, _
                                               lpLuid As LUID) As Long
Public Declare Function OpenProcessToken _
                Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
                                    ByVal DesiredAccess As Long, _
                                    TokenHandle As Long) As Long

Public Type MEMORY_CHUNKS
        Address As Long
        pData As Long
        Length As Long
End Type
Public Type LUID
        UsedPart As Long
        IgnoredForNowHigh32BitPart As Long
End Type '

Public Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
End Type
Public Const SE_CREATE_TOKEN = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT = "SeMachineAccountPrivilege"
Public Const SE_TCB = "SeTcbPrivilege"
Public Const SE_SECURITY = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP = "SeBackupPrivilege"
Public Const SE_RESTORE = "SeRestorePrivilege"
Public Const SE_SHUTDOWN = "SeShutdownPrivilege"
Public Const SE_DEBUG = "SeDebugPrivilege"
Public Const SE_AUDIT = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN = "SeRemoteShutdownPrivilege"
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const LVM_FIRST As Long = &H1000
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Public Enum SYSTEM_INFORMATION_CLASS
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
Public Declare Function ZwQuerySystemInformation _
                Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
                                 ByVal pSystemInformation As Long, _
                                 ByVal SystemInformationLength As Long, _
                                 ByRef ReturnLength As Long) As Long
Public Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
        UniqueProcessId As Integer
        CreatorBackTraceIndex As Integer
        ObjectTypeIndex As Byte
        HandleAttributes As Byte
        HandleValue As Integer
        pObject As Long
        GrantedAccess As Long
End Type
Public Type SYSTEM_HANDLE_INFORMATION
        NumberOfHandles As Long
        Handles(1 To 1) As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type
Public Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Public Const STATUS_ACCESS_DENIED = &HC0000022
Public Declare Function ZwWriteVirtualMemory _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal BaseAddress As Long, _
                                ByVal pBuffer As Long, _
                                ByVal NumberOfBytesToWrite As Long, _
                                ByRef NumberOfBytesWritten As Long) As Long
Public Declare Function ZwOpenProcess _
               Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, _
                                ByVal AccessMask As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
                                ByRef ClientId As CLIENT_ID) As Long
Public Type OBJECT_ATTRIBUTES
        Length As Long
        RootDirectory As Long
        ObjectName As Long 'PUNICODE_STRING 的指针
        Attributes As Long
        SecurityDescriptor As Long
        SecurityQualityOfService As Long
End Type
Public Type CLIENT_ID
        UniqueProcess As Long
        UniqueThread  As Long
End Type
Public Const PROCESS_QUERY_INFORMATION01 As Long = &H400
Public Const STATUS_INVALID_CID As Long = &HC000000B
Public Declare Function ZwClose _
               Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long
Public Const ZwGetCurrentProcess As Long = -1 '//0xFFFFFFFF
Public Const ZwGetCurrentThread As Long = -2 '//0xFFFFFFFE
Public Const ZwCurrentProcess As Long = ZwGetCurrentProcess
Public Const ZwCurrentThread As Long = ZwGetCurrentThread
Public Declare Function ZwCreateJobObject _
               Lib "NTDLL.DLL" (ByRef JobHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByRef ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Public Declare Function ZwAssignProcessToJobObject _
               Lib "NTDLL.DLL" (ByVal JobHandle As Long, _
                                ByVal ProcessHandle As Long) As Long
Public Declare Function ZwTerminateJobObject _
               Lib "NTDLL.DLL" (ByVal JobHandle As Long, _
                                ByVal ExitStatus As Long) As Long
Public Const OBJ_INHERIT = &H2
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const SYNCHRONIZE As Long = &H100000
Public Const JOB_OBJECT_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1F
Public Const PROCESS_DUP_HANDLE As Long = &H40
Public Const PROCESS_ALL_ACCESS01 As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)
Public Const OB_TYPE_PROCESS As Long = &H5 '// hard code
Public Type PROCESS_BASIC_INFORMATION
        ExitStatus As Long 'NTSTATUS
        PebBaseAddress As Long 'PPEB
        AffinityMask As Long 'ULONG_PTR
        BasePriority As Long 'KPRIORITY
        UniqueProcessId As Long 'ULONG_PTR
        InheritedFromUniqueProcessId As Long 'ULONG_PTR
End Type
Public Declare Function ZwDuplicateObject _
               Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, _
                                ByVal SourceHandle As Long, _
                                ByVal TargetProcessHandle As Long, _
                                ByRef TargetHandle As Long, _
                                ByVal DesiredAccess As Long, _
                                ByVal HandleAttributes As Long, _
                                ByVal Options As Long) As Long
Public Const DUPLICATE_CLOSE_SOURCE = &H1            '// winnt
Public Const DUPLICATE_SAME_ACCESS = &H2                '// winnt
Public Const DUPLICATE_SAME_ATTRIBUTES = &H4
Public Declare Function ZwQueryInformationProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ProcessInformationClass As PROCESSINFOCLASS, _
                                ByVal ProcessInformation As Long, _
                                ByVal ProcessInformationLength As Long, _
                                ByRef ReturnLength As Long) As Long
Public Enum PROCESSINFOCLASS
        ProcessBasicInformation
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
        ProcessIoPortHandlers           '// Note: this is kernel mode only
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
        MaxProcessInfoClass             '// MaxProcessInfoClass should always be the last enum
End Enum
Public Const STATUS_SUCCESS As Long = &H0
Public Const STATUS_INVALID_PARAMETER As Long = &HC000000D
Public Declare Function EnumProcesses _
                Lib "psapi.dll" (ByVal lpidProcess As Long, _
                                 ByVal cb As Long, _
                                 ByRef cbNeeded As Long) As Long
Public Declare Function Api_GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function ZwTerminateProcess _
               Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, _
                                ByVal ExitStatus As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function NT_SUCCESS(ByVal Status As Long) As Boolean
        NT_SUCCESS = (Status >= 0)
End Function

Public Sub CopyMemory(ByVal Dest As Long, ByVal Src As Long, ByVal cch As Long)
Dim Written As Long
        Call ZwWriteVirtualMemory(ZwCurrentProcess, Dest, Src, cch, Written)
End Sub

Public Function IsItemInArray(ByVal dwItem, ByRef dwArray() As Long) As Boolean
Dim Index As Long
        For Index = LBound(dwArray) To UBound(dwArray)
                If (dwItem = dwArray(Index)) Then IsItemInArray = True: Exit Function '// found
        Next
        IsItemInArray = False
End Function

Public Sub AddItemToArray(ByVal dwItem As Long, ByRef dwArray() As Long)
On Error GoTo ErrHdl

        If (IsItemInArray(dwItem, dwArray)) Then Exit Sub '// found
        
        ReDim Preserve dwArray(UBound(dwArray) + 1)
        dwArray(UBound(dwArray)) = dwItem
ErrHdl:
        
End Sub

'进程控制类函数 - 可以绕过一些API HOOK。
'函数需要 SE_DEBUG 特权。
'FlowerCode 的方法
'dwDesiredAccess: 需要的访问权
'bInhert: 句柄可以继承
'ProcessId: 进程 ID
Public Function OpenProcess01(ByVal dwDesiredAccess As Long, ByVal bInhert As Boolean, ByVal ProcessId As Long) As Long
        Dim st As Long
        Dim cid As CLIENT_ID
        Dim oa As OBJECT_ATTRIBUTES
        Dim NumOfHandle As Long
        Dim pbi As PROCESS_BASIC_INFORMATION
        Dim I As Long
        Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
        oa.Length = Len(oa)
        If (bInhert) Then oa.Attributes = oa.Attributes Or OBJ_INHERIT
        cid.UniqueProcess = ProcessId + 1 '// 呵呵.
        st = ZwOpenProcess(hProcessToRet, dwDesiredAccess, oa, cid)
        If (NT_SUCCESS(st)) Then OpenProcess01 = hProcessToRet: Exit Function
        st = STATUS_SUCCESS
        Dim bytBuf() As Byte
        Dim arySize As Long: arySize = &H20000 '// 128KB
        Do
                ReDim bytBuf(arySize)
                st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
                If (Not NT_SUCCESS(st)) Then
                        If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                                Erase bytBuf
                                Exit Function
                        End If
                Else
                        Exit Do
                End If

                arySize = arySize * 2
                ReDim bytBuf(arySize)
        Loop
        NumOfHandle = 0
        Call CopyMemory(VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle))
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        Call CopyMemory(VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle)
        For I = LBound(h_info) To UBound(h_info)
                With h_info(I)
                        If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then 'OB_TYPE_PROCESS is hardcode, you'd better get it dynamiclly
                                cid.UniqueProcess = .UniqueProcessId
                                st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, oa, cid)
                                If (NT_SUCCESS(st)) Then
                                        st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwCurrentProcess, hProcessCur, PROCESS_ALL_ACCESS01, 0, DUPLICATE_SAME_ATTRIBUTES)
                                        If (NT_SUCCESS(st)) Then
                                                st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)
                                                If (NT_SUCCESS(st)) Then
                                                        If (pbi.UniqueProcessId = ProcessId) Then
                                                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, OBJ_INHERIT, DUPLICATE_SAME_ATTRIBUTES)
                                                                If (NT_SUCCESS(st)) Then OpenProcess01 = hProcessToRet
                                                        End If
                                                End If
                                        End If
                                        st = ZwClose(hProcessCur)
                                End If
                                st = ZwClose(hProcessToDup)
                        End If
                End With
        Next
        Erase h_info
End Function

'ret val: bSuccess
'willy123 的方法
'hProcess: 进程句柄
'ExitStatus: 退出状态
Public Function TerminateProcess01(ByVal hProcess As Long, ByVal ExitStatus As Long) As Boolean
        Dim st As Long
        Dim hJob As Long
        Dim oa As OBJECT_ATTRIBUTES
        TerminateProcess01 = False
        oa.Length = Len(oa)
        st = ZwCreateJobObject(hJob, JOB_OBJECT_ALL_ACCESS, oa)
        If (NT_SUCCESS(st)) Then
                st = ZwAssignProcessToJobObject(hJob, hProcess)
                If (NT_SUCCESS(st)) Then
                        st = ZwTerminateJobObject(hJob, ExitStatus)
                        If (NT_SUCCESS(st)) Then TerminateProcess01 = True
                End If
                ZwClose (hJob)
        End If
End Function


Public Function EnablePrivilege(ByVal seName As String) As Boolean
        On Error Resume Next
        Dim p_lngRtn As Long
        Dim p_lngToken As Long
        Dim p_lngBufferLen As Long
        Dim p_typLUID As LUID
        Dim p_typTokenPriv As TOKEN_PRIVILEGES
        Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
        p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        If Err.LastDllError <> 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_typTokenPriv.PrivilegeCount = 1
        p_typTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        p_typTokenPriv.TheLuid = p_typLUID
        EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function




