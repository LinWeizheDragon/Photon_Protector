Attribute VB_Name = "mdlFunction"
Option Explicit
Dim IsFree As Boolean
Dim CRC32 As New clsGetCRC32

Function GetChecksum(sFile As String) As String
    On Error GoTo ErrHandle
    Dim cb0 As Byte
    Dim cb1 As Byte
    Dim cb2 As Byte
    Dim cb3 As Byte
    Dim cb4 As Byte
    Dim cb5 As Byte
    Dim cb6 As Byte
    Dim cb7 As Byte
    Dim cb8 As Byte
    Dim cb9 As Byte
    Dim cb10 As Byte
    Dim cb11 As Byte
    Dim cb12 As Byte
    Dim cb13 As Byte
    Dim cb14 As Byte
    Dim cb15 As Byte
    Dim cb16 As Byte
    Dim cb17 As Byte
    Dim cb18 As Byte
    Dim cb19 As Byte
    Dim cb20 As Byte
    Dim cb21 As Byte
    Dim cb22 As Byte
    Dim cb23 As Byte
    Dim buff As String
    Open sFile For Binary Access Read As #1
        buff = Space$(1)
        Get #1, , buff
    Close #1
    Open sFile For Binary Access Read As #2
        Get #2, 512, cb0
        Get #2, 1024, cb1
        Get #2, 2048, cb2
        Get #2, 3000, cb3
        Get #2, 4096, cb4
        Get #2, 5000, cb5
        Get #2, 6000, cb6
        Get #2, 7000, cb7
        Get #2, 8192, cb8
        Get #2, 9000, cb9
        Get #2, 10000, cb10
        Get #2, 11000, cb11
        Get #2, 12288, cb12
        Get #2, 13000, cb13
        Get #2, 14000, cb14
        Get #2, 15000, cb15
        Get #2, 16384, cb16
        Get #2, 17000, cb17
        Get #2, 18000, cb18
        Get #2, 19000, cb19
        Get #2, 20480, cb20
        Get #2, 21000, cb21
        Get #2, 22000, cb22
        Get #2, 23000, cb23
    Close #2
    buff = cb0
    buff = buff & cb1
    buff = buff & cb2
    buff = buff & cb3
    buff = buff & cb4
    buff = buff & cb5
    buff = buff & cb6
    buff = buff & cb7
    buff = buff & cb8
    buff = buff & cb9
    buff = buff & cb10
    buff = buff & cb11
    buff = buff & cb12
    buff = buff & cb13
    buff = buff & cb14
    buff = buff & cb15
    buff = buff & cb16
    buff = buff & cb17
    buff = buff & cb18
    buff = buff & cb19
    buff = buff & cb20
    buff = buff & cb21
    buff = buff & cb22
    buff = buff & cb23
    GetChecksum = CRC32.StringChecksum(buff)
    Set CRC32 = Nothing
    Exit Function
ErrHandle:
    Close #2
End Function

Function GetFullCRC(sFile As String) As String
    GetFullCRC = CRC32.FileChecksum(sFile)
End Function
