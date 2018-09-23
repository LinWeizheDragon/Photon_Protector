Attribute VB_Name = "modGetICON"
Public Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Public Type CLSID
    id(16) As Byte
End Type
'Public Const MAX_PATH = 260
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Const SHGFI_ICON = &H100
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1

Public Function IconToPicture(hIcon As Long) As IPictureDisp    'ICON 转 Picture

Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown

    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    Dim CA As ColorConstants
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    If hRes = 0 Then Set IconToPicture = lpUnk
    
End Function


Public Function GetIcon(FileName, Optional ByVal SmallIcon As Boolean = True) As IPictureDisp   '获得文件ICON

Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
    If SmallIcon = True Then
        SHGetFileInfo FileName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_SMALLICON
    Else
        SHGetFileInfo FileName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
    End If
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetIcon = icon_pic
    
End Function

