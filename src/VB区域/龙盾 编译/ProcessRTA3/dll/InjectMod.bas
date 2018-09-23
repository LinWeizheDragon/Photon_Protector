Attribute VB_Name = "InjectMod"
'Download by http://www.codefans.net
Option Explicit
'dll注入程序
'api申明模块
'
'蓝色炫影  制作
'www.rekersoft.cn
'
'最后更新 2008/05/06
'您可以自由用于非商业用途。
'请保留此行版权信息，谢谢。

'菜鸟学飞 作了微小改动 2010/7/19


Public Const MEM_COMMIT = 4096

Public Const MEM_DECOMMIT = &H4000

Public Const PAGE_READWRITE = 4


Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'这两个api的作用是在目标进程中分配一段空白内存供程序使用。在vb的api浏览器中是找不到的。

Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'得到函数地址与dll模块地址

Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'这里注意lpBaseAddress的传送方式是byval，和api浏览器中的声明是不一样的。 _
 byval是传值，默认是byref是传址，也就是传递的是参数在内存中的地址

Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public MyDllFileName As String

Public MyDllFileLength         As Long   'dll文件名长度



Public Sub GetDllfilepath()

        Dim tmp As Long

        MyDllFileName = Space$(255)
        tmp = GetModuleFileName(DLLhandle, MyDllFileName, 255)
        If tmp <> 0 Then
            MyDllFileName = Left$(MyDllFileName, tmp)
            MyDllFileLength = strlen(ByVal MyDllFileName)
        End If

End Sub

Public Sub InjectMyself(ByVal Process As Long)

        '注入子程序

        Dim MyDllFileBuffer         As Long   '写入dll文件名的内存地址

        Dim MyAddr                  As Long   '执行远程线程代码的起始地址。这里等于LoadLibraryA的地址

        Dim MyReturn                As Long
        
        Dim MyResult As Long
        
        '得到进程的句柄
        If Process = 0 Then GoTo errhandle
        
        MyDllFileBuffer = VirtualAllocEx(Process, 0, MyDllFileLength + 1, MEM_COMMIT, PAGE_READWRITE)
        '在目标进程中申请分配一块空白内存区域。内存的起始地址保存在MyDllFileBuffer中。 _
         这块内存区域我们用来存放dll文件路径，并作为参数传递给LoadLibraryA。

        If MyDllFileBuffer = 0 Then GoTo errhandle
        
        MyReturn = WriteProcessMemory(Process, MyDllFileBuffer, ByVal (MyDllFileName), MyDllFileLength + 1, ByVal 0)
        '在分配出来的内存区域中写入dll路径径。注意第二个参数传递的是MyDllFileBuffer的内容， _
         而不是MyDllFileBuffer的内存地址?

        If MyReturn = 0 Then GoTo errhandle

        MyAddr = GetProcAddress(GetModuleHandle("Kernel32"), "LoadLibraryA")
        '得到LoadLibraryA函数的起始地址。他的参数就是我们刚才写入的dll路径。但是LoadLibraryA本身是不知道参数在哪里的。 _
         接下来我们就用CreateRemoteThread函数告诉他参数放在哪里了?

        If MyAddr = 0 Then GoTo errhandle

        MyResult = CreateRemoteThread(Process, 0, 0, MyAddr, MyDllFileBuffer, 0, 0)
        '好了,现在用CreateRemoteThread在目标进程创建一个线程，线程起始地址指向LoadLibraryA， _
         参数就是MyDllFileBuffer中保存的dll路径?


errhandle:
        Exit Sub
End Sub

