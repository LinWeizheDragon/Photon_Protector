Attribute VB_Name = "modFileWork"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4
Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Long
     fAnyOperationsAborted As Long
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Function SuperSleep(DealyTime As Single) '´Ë´¦Ô­Îªlong£¬ÐÞ¸ÄÎªsingle¿ÉÑÓÊ±1ms :SK<2<8h
Dim TimerCount As Single
    TimerCount = Timer + DealyTime 'Ôö¼ÓXÃë ZJ9x6|q
    While TimerCount - Timer > 0
        DoEvents
        Sleep 1
    Wend
End Function
Public Function FileDel(str1 As String) As Long
    Dim result As Long, fileop As SHFILEOPSTRUCT
    With fileop
        .hWnd = 0
        .wFunc = FO_DELETE
        .pFrom = str1 & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    result = SHFileOperation(fileop)
End Function
