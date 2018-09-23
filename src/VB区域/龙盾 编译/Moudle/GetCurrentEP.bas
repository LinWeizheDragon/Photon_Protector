Attribute VB_Name = "GetEP"
Option Explicit

Public Const STATUS_ACCESS_DENIED = &HC0000022
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const NO_INHERITANCE = 0
Public Const DACL_SECURITY_INFORMATION = &H4

Public Const PROCESS_QUERY_INFORMATION As Long = &H400


Public Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As Long
End Type

Public Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDeor As Long
    SecurityQualityOfService As Long
End Type

Public Enum ACCESS_MODE
    NOT_USED_ACCESS
    GRANT_ACCESS
    SET_ACCESS
    DENY_ACCESS
    REVOKE_ACCESS
    SET_AUDIT_SUCCESS
    SET_AUDIT_FAILURE
End Enum

Public Enum MULTIPLE_TRUSTEE_OPERATION
    NO_MULTIPLE_TRUSTEE
    TRUSTEE_IS_IMPERSONATE
End Enum

Public Enum TRUSTEE_FORM
    TRUSTEE_IS_SID
    TRUSTEE_IS_NAME
End Enum

Public Enum TRUSTEE_TYPE
    TRUSTEE_IS_UNKNOWN
    TRUSTEE_IS_USER
    TRUSTEE_IS_GROUP
End Enum

Public Type TRUSTEE
    pMultipleTrustee            As Long
    MultipleTrusteeOperation    As MULTIPLE_TRUSTEE_OPERATION
    TrusteeForm                 As TRUSTEE_FORM
    TrusteeType                 As TRUSTEE_TYPE
    ptstrName                   As String
End Type

Public Type EXPLICIT_ACCESS
    grfAccessPermissions        As Long
    grfAccessMode               As ACCESS_MODE
    grfInheritance              As Long
    TRUSTEE                     As TRUSTEE
End Type

Public Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
End Enum

Public Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any) As Long
Public Declare Function GetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any, ppSecurityDeor As Long) As Long
Public Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal cCountOfExplicitEntries As Long, pListOfExplicitEntries As EXPLICIT_ACCESS, ByVal OldAcl As Long, NewAcl As Long) As Long
Public Declare Sub RtlInitUnicodeString Lib "ntdll.dll" (DestinationString As UNICODE_STRING, ByVal SourceString As Long)
Public Declare Function ZwOpenSection Lib "ntdll.dll" (SectionHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As Any) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
   
Public verinfo As OSVERSIONINFO
   
Public g_pMapPhysicalMemory As Long
Public g_hMPM As Long
Public aByte(3) As Byte

Public Type CLIENT_ID
        UniqueProcess As Long
        UniqueThread  As Long
End Type

Public Function GetEProcess() As Long
Dim thread As Long, process As Long, fw As Long, bw As Long
    If OpenPhysicalMemory <> 0 Then
        thread = GetData(&HFFDFF124)
        process = GetData(thread + &H44)
        GetEProcess = process
        'SetData process + lOffsetPID, FalsePID
        CloseHandle g_hMPM
    End If
End Function

Public Sub SetPhyscialMemorySectionCanBeWrited(ByVal hSection As Long)
    Dim pDacl As Long
    Dim pNewDacl As Long
    Dim pSD As Long
    Dim dwRes As Long
    Dim ea As EXPLICIT_ACCESS
   
    GetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, pDacl, 0, pSD
         
    ea.grfAccessPermissions = SECTION_MAP_WRITE
    ea.grfAccessMode = GRANT_ACCESS
    ea.grfInheritance = NO_INHERITANCE
    ea.TRUSTEE.TrusteeForm = TRUSTEE_IS_NAME
    ea.TRUSTEE.TrusteeType = TRUSTEE_IS_USER
    ea.TRUSTEE.ptstrName = "CURRENT_USER" & vbNullChar

    SetEntriesInAcl 1, ea, pDacl, pNewDacl
   
    SetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, ByVal pNewDacl, 0
                                
CleanUp:
    LocalFree pSD
    LocalFree pNewDacl
End Sub
Public Function OpenPhysicalMemory() As Long
Dim Status As Long
Dim PhysmemString As UNICODE_STRING
Dim Attributes As OBJECT_ATTRIBUTES
    RtlInitUnicodeString PhysmemString, StrPtr("\Device\PhysicalMemory")
    Attributes.Length = Len(Attributes)
    Attributes.RootDirectory = 0
    Attributes.ObjectName = VarPtr(PhysmemString)
    Attributes.Attributes = 0
    Attributes.SecurityDeor = 0
    Attributes.SecurityQualityOfService = 0
   
    Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
    If Status = STATUS_ACCESS_DENIED Then
        Status = ZwOpenSection(g_hMPM, READ_CONTROL Or WRITE_DAC, Attributes)
        SetPhyscialMemorySectionCanBeWrited g_hMPM
        CloseHandle g_hMPM
        Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
    End If
   
    Dim lDirectoty As Long
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    If (GetVersionEx(verinfo)) <> 0 Then
        If verinfo.dwPlatformId = 2 Then
            If verinfo.dwMajorVersion = 5 Then
                Select Case verinfo.dwMinorVersion
                    Case 0
                        lDirectoty = &H30000
                    Case 1
                        lDirectoty = &H39000
                End Select
            End If
        End If
    End If
   
    If Status = 0 Then
        g_pMapPhysicalMemory = MapViewOfFile(g_hMPM, 4, 0, lDirectoty, &H1000)
        If g_pMapPhysicalMemory <> 0 Then OpenPhysicalMemory = g_hMPM
    End If
End Function
Public Function LinearToPhys(BaseAddress As Long, addr As Long) As Long
    Dim VAddr As Long, PGDE As Long, PTE As Long, PAddr As Long
    Dim lTemp As Long
   
    VAddr = addr
    CopyMemory aByte(0), VAddr, 4
    lTemp = Fix(ByteArrToLong(aByte) / (2 ^ 22))
   
    PGDE = BaseAddress + lTemp * 4
    CopyMemory PGDE, ByVal PGDE, 4
   
    If (PGDE And 1) <> 0 Then
        lTemp = PGDE And &H80
        If lTemp <> 0 Then
            PAddr = (PGDE And &HFFC00000) + (VAddr And &H3FFFFF)
        Else
            PGDE = MapViewOfFile(g_hMPM, 4, 0, PGDE And &HFFFFF000, &H1000)
            lTemp = (VAddr And &H3FF000) / (2 ^ 12)
            PTE = PGDE + lTemp * 4
            CopyMemory PTE, ByVal PTE, 4
            
            If (PTE And 1) <> 0 Then
                PAddr = (PTE And &HFFFFF000) + (VAddr And &HFFF)
                UnmapViewOfFile PGDE
            End If
        End If
    End If
   
    LinearToPhys = PAddr
End Function
Public Function GetData(addr As Long) As Long
    Dim phys As Long, tmp As Long, ret As Long
   
    phys = LinearToPhys(g_pMapPhysicalMemory, addr)
    tmp = MapViewOfFile(g_hMPM, 4, 0, phys And &HFFFFF000, &H1000)
    If tmp <> 0 Then
        ret = tmp + ((phys And &HFFF) / (2 ^ 2)) * 4
        CopyMemory ret, ByVal ret, 4
        
        UnmapViewOfFile tmp
        GetData = ret
    End If
End Function
Public Function ByteArrToLong(inByte() As Byte) As Double
    Dim i As Integer
    For i = 0 To 3
        ByteArrToLong = ByteArrToLong + inByte(i) * (&H100 ^ i)
    Next i
End Function

