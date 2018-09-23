Attribute VB_Name = "modRemoveUsbDrv"


Option Explicit
'****************************************************************************************************************
'此模块是通过转换C++代码而来
'****************************************************************************************************************
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'typedef struct _SP_DEVICE_INTERFACE_DETAIL_DATA_A {
'    DWORD  cbSize;
'    CHAR   DevicePath[ANYSIZE_ARRAY];
'} SP_DEVICE_INTERFACE_DETAIL_DATA_A, *PSP_DEVICE_INTERFACE_DETAIL_DATA_A;
Private Type SP_DEVICE_INTERFACE_DETAIL_DATA
    cbSize As Long
    strDevicePath As String * 260
End Type
Private Type SP_DEVICE_INTERFACE_DATA
    cbSize As Long 'taille de la structure en octets
    InterfaceClassGuid As GUID 'GUID de la classe d'interface
    flags As Long 'options
    Reserved As Long 'réservé
End Type
Private Type SP_DEVINFO_DATA
    cbSize As Long 'taille de la structure en octets
    ClassGuid As GUID 'GUID de la classe d'installation
    DevInst As Long 'handle utilisable par certaine fonction CM_xxx
    Reserved As Long 'réservé
End Type
'
'typedef struct _STORAGE_DEVICE_NUMBER {
'    //
'    // The FILE_DEVICE_XXX type for this device.
'    //
'    DEVICE_TYPE DeviceType;
'    //
'    // The number of this device
'    //
'    DWORD       DeviceNumber;
'    //
'    // If the device is partitionable, the partition number of the device.
'    // Otherwise -1
'    //
'    DWORD       PartitionNumber;
'} STORAGE_DEVICE_NUMBER, *PSTORAGE_DEVICE_NUMBER;
Private Type STORAGE_DEVICE_NUMBER
    dwDeviceType As Long
    dwDeviceNumber As Long
    dwPartitionNumber As Long
End Type
'typedef enum    _PNP_VETO_TYPE {
'    PNP_VetoTypeUnknown,            // Name is unspecified
'    PNP_VetoLegacyDevice,           // Name is an Instance Path
'    PNP_VetoPendingClose,           // Name is an Instance Path
'    PNP_VetoWindowsApp,             // Name is a Module
'    PNP_VetoWindowsService,         // Name is a Service
'    PNP_VetoOutstandingOpen,        // Name is an Instance Path
'    PNP_VetoDevice,                 // Name is an Instance Path
'    PNP_VetoDriver,                 // Name is a Driver Service Name
'    PNP_VetoIllegalDeviceRequest,   // Name is an Instance Path
'    PNP_VetoInsufficientPower,      // Name is unspecified
'    PNP_VetoNonDisableable,         // Name is an Instance Path
'    PNP_VetoLegacyDriver,           // Name is a Service
'    PNP_VetoInsufficientRights      // Name is unspecified
'}   PNP_VETO_TYPE, *PPNP_VETO_TYPE;
Private Enum PNP_VETO_TYPE
    PNP_VetoTypeUnknown
    PNP_VetoLegacyDevice
    PNP_VetoPendingClose
    PNP_VetoWindowsApp
    PNP_VetoWindowsService
    PNP_VetoOutstandingOpen
    PNP_VetoDevice
    PNP_VetoDriver
    PNP_VetoIllegalDeviceRequest
    PNP_VetoInsufficientPower
    PNP_VetoNonDisableable
    PNP_VetoLegacyDriver
    PNP_VetoInsufficientRights
End Enum
'Private Const DIGCF_DEFAULT = &H1                        ' only valid with DIGCF_DEVICEINTERFACE
Private Const DIGCF_PRESENT = &H2
'Private Const DIGCF_ALLCLASSES = &H4
'Private Const DIGCF_PROFILE = &H8
Private Const DIGCF_DEVICEINTERFACE = &H10
Private Const GENERIC_READ = &H80000000   '允许对设备进行读访问
Private Const FILE_SHARE_READ = &H1       '允许读取共享
Private Const OPEN_EXISTING = 3           '文件必须已经存在。由设备提出要求
Private Const FILE_SHARE_WRITE = &H2      '允许对文件进行共享访问
Private Const IOCTL_STORAGE_BASE As Long = &H2D
Private Const METHOD_BUFFERED = 0
Private Const FILE_ANY_ACCESS = 0
Private Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByVal ClassGuid As Long, ByVal Enumerator As Long, ByVal HwndParent As Long, ByVal flags As Long) As Long
Private Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Private Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, DeviceInterfaceDetailData As Any, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, DeviceInfoData As Any) As Long
Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CM_Get_Parent Lib "cfgmgr32.dll" (pdwDevInst As Long, ByVal dwDevInst As Long, ByVal ulFlags As Long) As Long
Private Declare Function CM_Request_Device_EjectW Lib "setupapi.dll" (ByVal dwDevInst As Long, ByVal pVetoType As Long, ByVal pszVetoName As String, ByVal ulNameLength As Long, ByVal ulFlags As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Function CTL_CODE(ByVal lDeviceType As Long, ByVal lFunction As Long, ByVal lMethod As Long, ByVal lAccess As Long) As Long
    CTL_CODE = (lDeviceType * 2 ^ 16&) Or (lAccess * 2 ^ 14&) Or (lFunction * 2 ^ 2) Or (lMethod)
End Function
'获取设备属性信息，希望得到系统中所安装的各种固定的和可移动的硬盘、优盘和CD/DVD-ROM/R/W的接口类型、序列号、产品ID等信息。
Private Function IOCTL_STORAGE_GET_DEVICE_NUMBER() As Long '2953344
    IOCTL_STORAGE_GET_DEVICE_NUMBER = CTL_CODE(IOCTL_STORAGE_BASE, &H420, METHOD_BUFFERED, FILE_ANY_ACCESS)
End Function
Private Function GetDrivesDevInstByDeviceNumber(ByVal lngDeviceNumber As Long, ByVal uDriveType As Long, ByVal szDosDeviceName As String) As Long
    Dim objGuid As GUID, hDevInfo As Long, dwIndex As Long, lngRes As Long, dwSize As Long
    Dim objSpdid As SP_DEVICE_INTERFACE_DATA, objSpdd As SP_DEVINFO_DATA, objPspdidd As SP_DEVICE_INTERFACE_DETAIL_DATA
    Dim hDrive As Long, objSdn As STORAGE_DEVICE_NUMBER, dwBytesReturned As Long
    Dim dwReturn As Long
    '处理GUID
    With objGuid
        .Data2 = &HB6BF
        .Data3 = &H11D0&
        .Data4(0) = &H94&
        .Data4(1) = &HF2&
        .Data4(2) = &H0&
        .Data4(3) = &HA0&
        .Data4(4) = &HC9&
        .Data4(5) = &H1E&
        .Data4(6) = &HFB&
        .Data4(7) = &H8B&
        Select Case uDriveType
            Case 2
                If InStr(szDosDeviceName, "\Floppy") Then
                    .Data1 = &H53F56311
                Else
                    .Data1 = &H53F56307
                End If
            Case 3
                .Data1 = &H53F56307
            Case 5
                .Data1 = &H53F56308
        End Select
    End With
    'Get device interface info set handle for all devices attached to system
    hDevInfo = SetupDiGetClassDevs(VarPtr(objGuid), 0, 0, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)
    If hDevInfo = -1 Then
        GetDrivesDevInstByDeviceNumber = 0
        Exit Function
    End If
    objSpdid.cbSize = Len(objSpdid)
    Do While 1
        lngRes = SetupDiEnumDeviceInterfaces(hDevInfo, 0, objGuid, dwIndex, objSpdid)
        If lngRes = 0 Then Exit Do
        dwSize = 0
        Call SetupDiGetDeviceInterfaceDetail(hDevInfo, objSpdid, ByVal 0&, 0, dwSize, ByVal 0&)
        If dwSize <> 0 And dwSize <= 1024 Then
            objPspdidd.cbSize = 5 'Len(objPspdidd) '这里十分注意这里必须是5不能用'Len(objPspdidd)
            objSpdd.cbSize = Len(objSpdd)
            lngRes = SetupDiGetDeviceInterfaceDetail(hDevInfo, objSpdid, objPspdidd, ByVal dwSize, dwReturn, objSpdd)
            If lngRes > 0 Then
                '打开设备
                hDrive = CreateFile(objPspdidd.strDevicePath, 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
                If hDrive <> -1 Then
                    '获取设备号
                    lngRes = DeviceIoControl(hDrive, IOCTL_STORAGE_GET_DEVICE_NUMBER, ByVal 0&, 0, objSdn, Len(objSdn), dwBytesReturned, ByVal 0&)
                    If lngRes Then
                        'match the given device number with the one of the current device
                        If lngDeviceNumber = objSdn.dwDeviceNumber Then
                            Call CloseHandle(hDrive)
                            SetupDiDestroyDeviceInfoList hDevInfo
                            GetDrivesDevInstByDeviceNumber = objSpdd.DevInst
                            Exit Function
                        End If
                    End If
                    Call CloseHandle(hDrive)
                End If
            End If
        End If
        dwIndex = dwIndex + 1
    Loop
    Call SetupDiDestroyDeviceInfoList(hDevInfo)
End Function
'************************************************************************************************
'参数为szDosDeviceName为USB的路径格式为"\\.\" & drive & ":"形式，blnIsShowNote参数是是否显示
'消息窗体的着用，这里需要注意的是在9X下只能把blnIsShowNote参数设置为FALSE
'************************************************************************************************
Public Function RemoveUsbDrive(ByVal szDosDeviceName As String, ByVal blnIsShowNote As Boolean) As Boolean
    Dim strDrive As String, dwDeviceNumber As Long, hVolume As Long, objSdn As STORAGE_DEVICE_NUMBER, dwBytesReturned As Long
    Dim lngRes As Long, uDriveType As Long, strDosDriveName As String, hDevInst As Long, uType As PNP_VETO_TYPE
    Dim strVetoName As String, blnSuccess As Boolean, dwDevInstParent As Long, i As Integer, pVetoType As Long
    '获取USB所在盘符
    strDrive = Right(szDosDeviceName, 2)
    dwDeviceNumber = -1
    '打开设备
    hVolume = CreateFile(szDosDeviceName, 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If hVolume = -1 Then
        RemoveUsbDrive = False
        Exit Function
    End If
    '获取设备号
    lngRes = DeviceIoControl(hVolume, IOCTL_STORAGE_GET_DEVICE_NUMBER, ByVal 0&, 0, objSdn, Len(objSdn), dwBytesReturned, ByVal 0&)
    If lngRes Then
        dwDeviceNumber = objSdn.dwDeviceNumber
    End If
    '关闭设备
    Call CloseHandle(hVolume)
    If dwDeviceNumber = -1 Then
        RemoveUsbDrive = False
        Exit Function
    End If
    '获取驱动器类型
    uDriveType = GetDriveType(strDrive)
    strDosDriveName = String(280, Chr(0))
    'get the dos device name (like \device\floppy0) to decide if it's a floppy or not - who knows a better way?
    lngRes = QueryDosDevice(strDrive, strDosDriveName, 280)
    strDosDriveName = Left(strDosDriveName, InStr(strDosDriveName, Chr(0)) - 1)
    If lngRes = 0 Then
        RemoveUsbDrive = False
        Exit Function
    End If
    'get the device instance handle of the storage volume by means of a SetupDi enum and matching the device number
    hDevInst = GetDrivesDevInstByDeviceNumber(dwDeviceNumber, uDriveType, strDosDriveName)
    If hDevInst = 0 Then
        RemoveUsbDrive = False
        Exit Function
    End If
    strVetoName = String(260, Chr(0))
    'get drives's parent, e.g. the USB bridge, the SATA port, an IDE channel with two drives!
    lngRes = CM_Get_Parent(dwDevInstParent, hDevInst, 0)
    For i = 0 To 3
        '卸载UB设备
        If blnIsShowNote Then
            lngRes = CM_Request_Device_EjectW(dwDevInstParent, ByVal VarPtr(pVetoType), vbNullString, 0, 0)
        Else
            lngRes = CM_Request_Device_EjectW(dwDevInstParent, uType, strVetoName, 260, 0)
        End If
        If lngRes = 0 And uType = PNP_VetoTypeUnknown Then
            blnSuccess = True
            Exit For
        End If
        Sleep 300
    Next
    RemoveUsbDrive = blnSuccess
End Function


