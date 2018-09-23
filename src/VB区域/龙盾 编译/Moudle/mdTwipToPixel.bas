Attribute VB_Name = "mdTwipToPixel"

Option Explicit
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" _
(ByVal hwnd As Long) As Long
Private Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" _
(ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public Const DIRECTION_VERTICAL = 1
Public Const DIRECTION_HORIZONTAL = 0

'===============================================================================
'-函数名称:         gFunTwipsToPixels
'-功能描述:         转换堤到像素
'-输入参数说明:     参数1:rlngTwips Long 需要转换的堤
'                   参数2:rlngDirection Long DIRECTION_VERTICAL是Y方向 DIRECTION_HORIZONTAL为X方向
'-返回参数说明:     转换后像素值
'-使用语法示例:     gFunTwipsToPixels 50,DIRECTION_VERTICAL
'-参考:
'-使用注意:
'-兼容性:           97,2000,XP compatible
'-作者:             王宇虹（参考微软KB)，改进：王宇虹
'-更新日期：        2002-08-26 ,2002-11-15
'===============================================================================
Function gFunTwipsToPixels(rlngTwips As Long, rlngDirection As Long) As Long
On Error GoTo Err_gFunTwipsToPixels
Dim lngDeviceHandle As Long
Dim lngPixelsPerInch As Long
lngDeviceHandle = apiGetDC(0)
If rlngDirection = DIRECTION_HORIZONTAL Then  '水平X方向
lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
Else       '垂直Y方向
lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
End If
lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
gFunTwipsToPixels = rlngTwips / 1440 * lngPixelsPerInch
Exit_gFunTwipsToPixels:
On Error Resume Next
Exit Function
Err_gFunTwipsToPixels:
MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
Resume Exit_gFunTwipsToPixels
End Function
'===============================================================================
'-函数名称:         gFunPixelsToTwips
'-功能描述:         转换像素到堤
'-输入参数说明:     参数1:rlngPixels Long 需要转换的像素
'                   参数2:rlngDirection Long DIRECTION_VERTICAL是Y方向 DIRECTION_HORIZONTAL为X方向
'-返回参数说明:     转换后堤值
'-使用语法示例:     gFunPixelsToTwips 50,DIRECTION_VERTICAL
'-参考:
'-使用注意:
'-兼容性:           97,2000,XP compatible
'-作者:             王宇虹（参考微软KB)，改进：王宇虹
'-更新日期：        2002-08-26 ,2002-11-15
'===============================================================================
Function gFunPixelsToTwips(rlngPixels As Long, rlngDirection As Long) As Long
On Error GoTo Err_gFunPixelsToTwips
Dim lngDeviceHandle As Long
Dim lngPixelsPerInch As Long
lngDeviceHandle = apiGetDC(0)
If rlngDirection = DIRECTION_HORIZONTAL Then  '水平X方向
lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
Else       '垂直Y方向
lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
End If
lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
gFunPixelsToTwips = rlngPixels * 1440 / lngPixelsPerInch
Exit_gFunPixelsToTwips:
On Error Resume Next
Exit Function
Err_gFunPixelsToTwips:
MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
Resume Exit_gFunPixelsToTwips
End Function


