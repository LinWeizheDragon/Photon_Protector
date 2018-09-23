VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSharedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'Class Module Name: CShareString
'Refactored by AdamBear 2002-1-31 22:00
'Originally Created By Bruce Meckinney, HardCore Visual Basic
'Purpose: Demostrate the FileMapping and Inter-Process Community
'************************************************************************

'************************************************************************
'修正了原来代码的一个错误，不过错误原因不知，问题涉及到DBCS和UNICODE转换
'现象：内部自动转换回UNICODE时只返回了该串的DBCS长度部分。
'     如: "Adam苯蛋Bear" 长度为10，返回却是"Adam笨蛋Be"长度为8，是其在DBCS
'中的长度
' 临时解决方案：将原串放大一倍，再重新计算长度，并截取。
' 经验：不要过于相信大师的代码，他也不一定把所有问题都考虑了。
'************************************************************************

Private Const ERROR_ALREADY_EXISTS = 183&
Private Const ERROR_BAD_ARGUMENTS = 160&
Private Const ERROR_INVALID_DATA = 13&




Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal strDst As String, pScr As Long, ByVal length As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (strDst As Long, ByVal pScr As String, ByVal length As Long)


Private Const PAGE_READWRITE = 4
Private Const FILE_MAP_WRITE = 2



Private h As Long                      '内存映射句柄
Private p As Long                      '映射视图指针





Sub Create(strName As String)
    Dim e As Long
    If strName = "" Then ApiRaise ERROR_BAD_ARGUMENTS
    ' 尝试建立一个大小为64K内存文件映射，不要怕设得大了，因为使用时是按4K为单位分配的。
    h = CreateFileMapping(-1, ByVal pNull, PAGE_READWRITE, _
                          0, 65535, strName)
    
    '保存错误描述，因其不一定真是错误，此处可能是正常的文件已存在信息，
    '   具体可参见CreateFileMapping的帮助文件
    e = Err.LastDllError
    If h = 0 Then ApiRaise e    '当然对CreateFileMapping这API，我们可以认为返回0才是错误。
    
    ' 取得映射文件的视图指针p，它指向就是从共享内存文件中映射来的进程可用的内存。
    p = MapViewOfFile(h, FILE_MAP_WRITE, 0, 0, 0)
    If p = pNull Then
        CloseHandle h   '一定要注意
        ApiRaise Err.LastDllError
    End If
    '
    If e <> ERROR_ALREADY_EXISTS Then
        'HACK BSTR 保存p的前四个字节作为BSTR的长度存放处。
        CopyMemory ByVal p, 0, 4
    Else
        '返回的是已经存在的FileMapping的句柄，不是错误
        'ApiRaise ERROR_ALREADY_EXISTS
    End If
End Sub

Private Sub Class_Terminate()
    '清理，重要，否则会有内存漏洞。
    '    当然，注释掉这两句话有时也有奇妙的用处，你可以在内存中留下东西，只要不重新起动Windows，
    '当程序再次起动后，还可以访问上次留下的内容。谁知道有什么用！
    UnmapViewOfFile p
    CloseHandle h
End Sub

'NOTE:为了让本类能做为String类型处理，须将本过程属性设为缺省。
Property Get item() As String
Attribute item.VB_UserMemId = 0
    If h = 0 Then ErrRaise ERROR_INVALID_DATA
    BugAssert p <> pNull
    Dim c As Long
    '取出BSTR的长度
    CopyMemory c, ByVal p, 4
    If c Then
        ' 由于未知原因，此处先将字符串放大一倍
        item = String$(c * 2, " ")
        CopyMemoryToStr item, ByVal (p + 4), c * 2
        Dim pos As Long
        pos = InStr(item, Chr(0))
        If pos Then item = Left(item, pos - 1)
    End If
End Property

Property Let item(s As String)
    If h = 0 Then ErrRaise ERROR_INVALID_DATA
    BugAssert p <> pNull
    Dim c As Long ', ab() As Byte
    c = Len(s)
    'ab = StrConv(s, vbFromUnicode) & vbNullChar
    ' 重新将BSTR放回内存中
    CopyMemory ByVal p, c, 4
    CopyMemoryStr ByVal (p + 4), s, c * 2
End Property
'


Private Sub ErrRaise(e As Long)
    Dim sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".SharedString"
        Err.Raise COMError(e), sSource
    Else
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub




