Attribute VB_Name = "modRemoveUsbDrive"
Option Explicit
'****************************************************************************************************************
'此模块来自于网络
'****************************************************************************************************************
'判断驱动器的类型
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_UNKNOWN = 0        '驱动器类型无法确定
Private Const DRIVE_NO_ROOT_DIR = 1    '驱动器根目录不存在
Private Const DRIVE_REMOVABLE = 2      '软盘驱动器或可移动盘
Private Const DRIVE_FIXED = 3          '硬盘驱动器
Private Const DRIVE_REMOTE = 4         'Network 驱动器
Private Const DRIVE_CDROM = 5          '光盘驱动器
Private Const DRIVE_RAMDISK = 6        'RAM 存储器
'****************************************************************************************************************
' CreateFile获取设备句柄
'参数
'lpFileName                       文件名
'dwDesiredAccess                  访问方式
'dwShareMode                      共享方式
'ATTRIBUTES lpSecurityAttributes  安全描述符指针
'dwCreationDisposition            创建方式
'dwFlagsAndAttributes             文件属性及标志
' hTemplateFile                   模板文件的句柄
'CreateFile这个函数用处很多，这里我们用它「打开」设备驱动程序，得到设备的句柄。
'操作完成後用CloseHandle关闭设备句柄。
'与普通文件名有所不同，设备驱动的「文件名」形式固定为「\\.\DeviceName」(注意在C程序中该字符串写法为「[url=]\\\\.\\DeviceName[/url]」)
'DeviceName必须与设备驱动程序内规定的设备名称一致。
'一般地，调用CreateFile获得设备句柄时，访问方式参数设置为0或GENERIC_READ|GENERIC_WRITE
'共享方式参数设置为FILE_SHARE_READ|FILE_SHARE_WRITE，创建方式参数设置为OPEN_EXISTING，其它参数设置为0或NULL。
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const GENERIC_READ = &H80000000   '允许对设备进行读访问
Private Const FILE_SHARE_READ = &H1       '允许读取共享
Private Const OPEN_EXISTING = 3           '文件必须已经存在。由设备提出要求
Private Const FILE_SHARE_WRITE = &H2      '允许对文件进行共享访问
'****************************************************************************************************************
'DeviceIoControl说明
'用途              实现对设备的访问—获取信息，发送命令，交换数据等。
'参数
'hDevice           设备句柄
'dwIoControlCode   控制码
'lpInBuffer        输入数据缓冲区指针
'nInBufferSize     输入数据缓冲区长度
'lpOutBuffer       输出数据缓冲区指针
'nOutBufferSize    输出数据缓冲区长度
'lpBytesReturned   输出数据实际长度单元长度
'lpOverlapped      重叠操作结构指针
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long                    '结构体的大小
    lpSecurityDescriptor As Long       '安全描述符（一个C-Style的字符串）。
    bInheritHandle As Long             '所创建出来的东西是可以被其他的子进程使用的
End Type
'查询存储设备还是适配器属性
Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0          '查询设备属性
    StorageAdapterProperty             '查询适配器属性
End Enum
'查询存储设备属性的类型
Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0          '读取描述
    PropertyExistsQuery                '测试是否支持
    PropertyMaskQuery                  '读取指定的描述
    PropertyQueryMaxDefined            '验证数据
End Enum
'查询属性输入的数据结构
Private Type STORAGE_PROPERTY_QUERY
    PropertyId As STORAGE_PROPERTY_ID  '设备/适配器
    QueryType As STORAGE_QUERY_TYPE    '查询类型
    AdditionalParameters(0) As Byte    '额外的数据(仅定义了象徵性的1个字节)
End Type
Private Type OVERLAPPED
    Internal As Long                  '保留给操作系统使用。用于保存系统状态，当GetOverLappedRseult的返回值中没有设置ERROR_IO_PENDING时，本域为有效。
    InternalHigh As Long              '成员保留给操作系统使用。用于保存异步传输数据的长度。当GetOverLappedRseult返回TRUE时，本域为有效。
    offset As Long                    '指定开始进行异步传输的文件的一个位置。该位置是距离文件开头处的偏移值。在调用ReadFile或WriteFile之前，必须设置此分量。
    OffsetHigh As Long                '指定开始异步传输处的字节偏移的高位字部分。
    hEvent As Long                    '指向一个事件的句柄，当传输完后将其设置为信号状态。
End Type
'存储设备的总线类型
Private Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum
'查询属性输出的数据结构
Private Type STORAGE_DEVICE_DESCRIPTOR
    Version As Long                 '版本
    Size As Long                    '结构大小
    DeviceType As Byte              '设备类型
    DeviceTypeModifier As Byte      'SCSI-2额外的设备类型
    RemovableMedia As Byte          '是否可移动
    CommandQueueing As Byte         '是否支持命令队列
    VendorIdOffset As Long          '厂家设定值的偏移
    ProductIdOffset As Long         '产品ID的偏移
    ProductRevisionOffset As Long   '产品版本的偏移
    SerialNumberOffset As Long      '序列号的偏移
    BusType As STORAGE_BUS_TYPE     '总线类型
    RawPropertiesLength As Long     '额外的属性数据长度
    RawDeviceProperties(0) As Byte  '额外的属性数据(仅定义了象徵性的1个字节)
End Type
'计算控制码   IOCTL_STORAGE_QUERY_PROPERTY
Private Const IOCTL_STORAGE_BASE As Long = &H2D
Private Const METHOD_BUFFERED = 0
Private Const FILE_ANY_ACCESS = 0
'获取磁盘属性
Private Function GetDisksProperty(ByVal hDevice As Long, utDevDesc As STORAGE_DEVICE_DESCRIPTOR) As Boolean
    Dim ut As OVERLAPPED
    Dim utQuery As STORAGE_PROPERTY_QUERY
    Dim lOutBytes As Long
    With utQuery
        .PropertyId = StorageDeviceProperty
        .QueryType = PropertyStandardQuery
    End With
    GetDisksProperty = DeviceIoControl(hDevice, IOCTL_STORAGE_QUERY_PROPERTY, utQuery, LenB(utQuery), utDevDesc, LenB(utDevDesc), lOutBytes, ut)
End Function
Private Function CTL_CODE(ByVal lDeviceType As Long, ByVal lFunction As Long, ByVal lMethod As Long, ByVal lAccess As Long) As Long
    CTL_CODE = (lDeviceType * 2 ^ 16&) Or (lAccess * 2 ^ 14&) Or (lFunction * 2 ^ 2) Or (lMethod)
End Function
'获取设备属性信息，希望得到系统中所安装的各种固定的和可移动的硬盘、优盘和CD/DVD-ROM/R/W的接口类型、序列号、产品ID等信息。
Private Function IOCTL_STORAGE_QUERY_PROPERTY() As Long
    IOCTL_STORAGE_QUERY_PROPERTY = CTL_CODE(IOCTL_STORAGE_BASE, &H500, METHOD_BUFFERED, FILE_ANY_ACCESS)
End Function
'获取驱动器总线类型
Public Function GetDriveBusType(ByVal strDriveLetter As String) As String
    Dim hDevice As Long
    Dim utDevDesc As STORAGE_DEVICE_DESCRIPTOR
    hDevice = CreateFile("\\.\" & strDriveLetter, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If hDevice <> -1 Then
        utDevDesc.Size = LenB(utDevDesc)
        Call GetDisksProperty(hDevice, utDevDesc)
        Select Case utDevDesc.BusType
            Case BusType1394
                GetDriveBusType = "1394"
            Case BusTypeAta
                GetDriveBusType = "Ata"
            Case BusTypeAtapi
                GetDriveBusType = "Atapi"
            Case BusTypeFibre
                GetDriveBusType = "Fibre"
            Case BusTypeRAID
                GetDriveBusType = "RAID"
            Case BusTypeScsi
                GetDriveBusType = "Scsi"
            Case BusTypeSsa
                GetDriveBusType = "Ssa"
            Case BusTypeUsb
                GetDriveBusType = "Usb"
            Case BusTypeUnknown
                GetDriveBusType = "未知"
            Case Else
        End Select
        Call CloseHandle(hDevice)
    End If
End Function
