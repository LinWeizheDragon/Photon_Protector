VERSION 5.00
Begin VB.UserControl jcbutton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   DefaultCancel   =   -1  'True
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "jcButton.ctx":0000
End
Attribute VB_Name = "jcbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'ÈËÈËÎªÎÒ£¬ÎÒÎªÈËÈË
'ÕíÉÆ¾Óºº»¯ÊÕ²ØÕûÀí
'·¢²¼ÈÕÆÚ£º2008/12/20
'Ãè    Êö£ºJCButton°´Å¥¿Ø¼þÕýÊ½°æ
'Íø    Õ¾£ºhttp://www.Mndsoft.com/  (VB6Ô´Âë²©¿Í)
'Íø    Õ¾£ºhttp://www.VbDnet.com/   (VB.NETÔ´Âë²©¿Í,Ö÷Òª»ùÓÚ.NET2005)
'e-mail  £ºMndsoft@163.com
'e-mail  £ºMndsoft@126.com
'OICQ    £º88382850
'          Èç¹ûÄúÓÐÐÂµÄºÃµÄ´úÂë±ðÍü¼Ç¸øÕíÉÆ¾ÓÅ¶!
'****************************************************************************
Option Explicit

'***************************************************************************
'*  Title:      JC button
'*  Function:   An ownerdrawn multistyle button
'*  Author:     Juned Chhipa
'*  Created:    November 2008
'*  Contact me: juned.chhipa@yahoo.com
'*
'*  Copyright © 2008-2009 Juned Chhipa. All rights reserved.
'***************************************************************************
'* This control can be used as an alternative to Command Button. It is
'* a lightweight button control which will emulate new command buttons.
'* Compile to get more faster results
'*
'* This control uses self-subclassing routines of Paul Caton.
'* Feel free to use this control. Please read Licence.txt
'* Please send comments/suggestions/bug reports to juned.chhipa@yahoo.com
'****************************************************************************
'*
'* - CREDITS:
'* - Paul Caton  :-  Self-Subclass Routines
'* - Noel Dacara :-  For helping me (Also, his dcbutton helped me a lot)
'* - Jim Jose    :-  To make grayscale (disabled) bitmap/icon
'* - Carles P.V. :-  For fastest gradient routines
'*   If any bugs found, please report  :- juned.chhipa@yahoo.com
'*
'* I have tested this control many times and I have tried my best to make
'* it work as a real command button. But still, I cannot guarantee that
'* that this is FREE OF BUGS. So please let me know if u find any.

'****************************************************************************
'* This software is provided "as-is" without any express/implied warranty.  *
'* In no event shall the author be held liable for any damages arising      *
'* from the use of this software.                                           *
'* If you do not agree with these terms, do not install "JCButton". Use     *
'* of the program implicitly means you have agreed to these terms.          *        *
'                                                                           *
'* Permission is granted to anyone to use this software for any purpose,    *
'* including commercial use, and to alter and redistribute it, provided     *
'* that the following conditions are met:                                   *
'*                                                                          *
'* 1.All redistributions of source code files must retain all copyright     *
'*   notices that are currently in place, and this list of conditions       *
'*   without any modification.                                              *
'*                                                                          *
'* 2.All redistributions in binary form must retain all occurrences of      *
'*   above copyright notice and web site addresses that are currently in    *
'*   place (for example, in the About boxes).                               *
'*                                                                          *
'* 3.Modified versions in source or binary form must be plainly marked as   *
'*   such, and must not be misrepresented as being the original software.   *
'****************************************************************************

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hicon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

'User32 Declares
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TransparentBlt Lib "MSIMG32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long

Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hicon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'==========================================================================================================================================================================================================================================================================================
' Subclassing Declares
Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

'Windows Messages
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOVING                 As Long = &H216
Private Const WM_NCACTIVATE             As Long = &H86
Private Const WM_ACTIVATE               As Long = &H6

Private Const ALL_MESSAGES              As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED                As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC               As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04                  As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05                  As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08                  As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09                  As Long = 137                                      'Table A (after) entry count patch offset

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                   As Long
    dwFlags                  As TRACKMOUSEEVENT_FLAGS
    hwndTrack                As Long
    dwHoverTime              As Long
End Type

'for subclass
Private Type tSubData                                                            'Subclass data type
    hWnd                      As Long                                            'Handle of the window being subclassed
    nAddrSub                  As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                 As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                  As Long                                            'Msg after table entry count
    nMsgCntB                  As Long                                            'Msg before table entry count
    aMsgTblA()                As Long                                            'Msg after table array
    aMsgTblB()                As Long                                            'Msg Before table array
End Type

'for subclass
Private sc_aSubData()       As tSubData                                        'Subclass data array
Private bTrack              As Boolean
Private bTrackUser32        As Boolean

'Kernel32 declares used by the Subclasser
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  End of Subclassing Declares
'==========================================================================================================================================================================================================================================================================================================

'[Enumerations]
Public Enum enumButtonStlyes
    [eStandard]                 '1) Standard VB Button
    [eFlat]                     '2) Standard Toolbar Button
    [eWindowsXP]                '3) Famous Win XP Button
    [eXPToolbar]                '4) XP Toolbar
    [eVistaAero]                '5) The New Vista Aero Button
    [eAOL]                      '6) AOL Buttons
    [eInstallShield]            '7) InstallShield?!?~?
    [eOutlook2007]              '8) Office 2007 Outlook Button
    [eVistaToolbar]             '9) Vista Toolbar Button
    [eVisualStudio]            '10) Visual Studio 2005 Button
    [eGelButton]               '11) Gel Button
    [e3DHover]                 '13) 3D Hover Button
    [eFlatHover]               '12) Flat Hover Button
    [eOffice2003]              '13) Office 2003 Style
    [eAqua]                    '14) Aqua (Near to MACOSX)
    [eSleek]                   '15) Somewhere I saw this (Not Bad!)
End Enum

#If False Then
    Private eStandard, eFlat, eVistaAero, eVistaToolbar, eInstallShield, eFlatHover, eOffice2003, eSleek
    Private eWindowsXP, eAqua, eXPToolbar, eVisualStudio, e3DHover, eGelButton, eOutlook2007, eAOL
#End If

Public Enum enumButtonStates
    [eStateNormal]              'Normal State
    [eStateOver]                'Hover State
    [eStateDown]                'Down State
End Enum

#If False Then
'A trick to preserve casing when typing in IDE
Private eStateNormal, eStateOver, eStateDown, eStateDisabled, eStateFocus
#End If

Public Enum enumCaptionAlign
    [ecLeftAlign]
    [ecCenterAlign]
    [ecRightAlign]
End Enum

#If False Then
'A trick to preserve casing when typing in IDE
Private ecLeftAlign, ecCenterAlign, ecRightAlign
#End If

Public Enum enumPictureAlign
    [epLeftEdge]
    [epLeftOfCaption]
    [epRightEdge]
    [epRightOfCaption]
    [epCenter]
    [epTopEdge]
    [epTopOfCaption]
    [epBottomEdge]
    [epBottomOfCaption]
End Enum

#If False Then
Private epLeftEdge, epRightEdge, epRightOfCaption, epLeftOfCaption, epCenter
Private epTopEdge, epTopOfCaption, epBottomEdge, epBottomOfCaption
#End If

Public Enum enumPictureSize
    [epsNormal]
    [eps16x16]
    [eps24x24]
    [eps32x32]
End Enum

#If False Then
    Private epsNormal, eps16x16, eps24x24, eps32x32
#End If

Public Enum GradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

#If False Then
Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If

'  used for Button colors
Private Type tButtonColors
    tBackColor      As Long
    tDisabledColor  As Long
    tForeColor      As Long
    tGreyText       As Long
End Type

'  used to define various graphics areas
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINT
    x       As Long
    y       As Long
End Type

'  RGB Colors structure
Private Type RGBColor
    r       As Single
    g       As Single
    b       As Single
End Type

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type ICONINFO
    fIcon       As Long
    xHotspot    As Long
    yHotspot    As Long
    hbmMask     As Long
    hbmColor    As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type
 
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
End Type
 
' --constants for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2
 
' --constants for  Flat Button
Private Const BDR_RAISEDINNER   As Long = &H4

' --constants for Win 98 style buttons
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private Const BF_LEFT       As Long = &H1
Private Const BF_TOP        As Long = &H2
Private Const BF_RIGHT      As Long = &H4
Private Const BF_BOTTOM     As Long = &H8
Private Const BF_RECT       As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' --System Hand Pointer
Private Const IDC_HAND As Long = 32649

' --Color Constant
Private Const CLR_INVALID       As Long = &HFFFF
Private Const DIB_RGB_COLORS    As Long = 0

' --Formatting Text Consts
Private Const DT_SINGLELINE     As Long = &H20

' --for drawing Icon Constants
Private Const DI_NORMAL As Long = &H3

' --Property Variables:
Private m_Picture           As StdPicture           'Icon of button
Private m_PicOver           As StdPicture
Private m_PicSize           As enumPictureSize
Private m_PictureAlign      As enumPictureAlign     'Picture Alignments
Private PicSizeW            As Long                 'Picture's Height
Private PicSizeH            As Long                 'Picture's Width

Private m_ButtonStyle       As enumButtonStlyes     'Choose your Style
Private m_Buttonstate       As enumButtonStates     'Normal / Over / Down

Private m_bIsDown           As Boolean              'Is button is pressed?
Private m_bMouseInCtl       As Boolean              'Is Mouse in Control
Private m_bHasFocus         As Boolean              'Has focus?
Private m_bHandPointer      As Boolean              'Use Hand Pointer
Private m_lCursor           As Long
Private m_bDefault          As Boolean              'Is Default?
Private m_bCheckBoxMode     As Boolean              'Is checkbox?
Private m_bValue            As Boolean              'Value (Checked/Unchekhed)
Private m_bShowFocus        As Boolean              'Bool to show focus
Private m_bParentActive     As Boolean              'Parent form Active or not
Private m_lParenthWnd       As Long                 'Is parent active?
Private m_WindowsNT         As Long                 'OS Supports Unicode?
Private m_bEnabled          As Boolean              'Enabled/Disabled
Private m_Caption           As String               'String to draw caption
Private m_TextRect          As RECT                 'Text Position
Private m_CapRect           As RECT                 'For InstallShield style
Private m_CaptionAlign      As enumCaptionAlign
Private m_bColors           As tButtonColors        'Button Colors
Private m_bUseMaskColor     As Boolean              'Transparent areas
Private m_lMaskColor        As Long                 'Set Transparent color
Private m_lButtonRgn        As Long                 'Button Region
Private m_bIsSpaceBarDown   As Boolean              'Space bar down boolean
Private m_ButtonRect        As RECT                 'Button Position
Private m_FocusRect         As RECT
Private WithEvents mFont    As StdFont
Attribute mFont.VB_VarHelpID = -1

Private m_lDownButton       As Integer              'For click/Dblclick events
Private m_lDShift           As Integer              'A flag for dblClick
Private m_lDX               As Single
Private m_lDY               As Single

Private lh                  As Long                 'ScaleHeight of button
Private lw                  As Long                 'ScaleWidth of button
Private XPos                As Long                 'X position of picture
Private YPos                As Long                 'Y Position of Picture

'  Events
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAcsii As Integer)

'  PRIVATE ROUTINES

Private Function PaintGrayScale(ByVal lhdc As Long, ByVal hPicture As Long, ByVal lLeft As Long, ByVal lTop As Long, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight As Long = -1) As Boolean

'****************************************************************************
'*  Converts an icon/bitmap to grayscale (used for Disabled buttons)        *
'*  Author:  Jim Jose                                                       *
'*  Modified by me for Disabled Bitmaps (for Maskcolor)
'*  All Credits goes to Jim Jose                                            *
'****************************************************************************

Dim BMP        As BITMAP
Dim BMPiH      As BITMAPINFOHEADER
Dim lBits()    As Byte 'Packed DIB
Dim lTrans()   As Byte 'Packed DIB
Dim TmpDC      As Long
Dim x          As Long
Dim xMax       As Long
Dim TmpCol     As Long
Dim r1         As Long
Dim g1         As Long
Dim b1         As Long
Dim bIsIcon    As Boolean

Dim hDCSrc   As Long
Dim hOldob   As Long
Dim PicSize  As Long
Dim oPic     As New StdPicture

    Set oPic = m_Picture

    '  Get the Image format
    If (GetObjectType(hPicture) = 0) Then
        Dim mIcon As ICONINFO
        bIsIcon = True
        GetIconInfo hPicture, mIcon
        hPicture = mIcon.hbmColor
    End If

    '  Get image info
    GetObject hPicture, Len(BMP), BMP

    '  Prepare DIB header and redim. lBits() array
    With BMPiH
        .biSize = Len(BMPiH) '40
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = BMP.bmWidth
        .biHeight = BMP.bmHeight
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        If lWidth = -1 Then lWidth = .biWidth
        If lHeight = -1 Then lHeight = .biHeight
    End With

    ReDim lBits(Len(BMPiH) + BMPiH.biSizeImage)   '[Header + Bits]

    '  Create TemDC and Get the image bits
    TmpDC = CreateCompatibleDC(lhdc)
    GetDIBits TmpDC, hPicture, 0, BMP.bmHeight, lBits(0), BMPiH, 0

    '  Loop through the array... (grayscale - average!!)
    xMax = BMPiH.biSizeImage - 1
    For x = 0 To xMax - 3 Step 3
        r1 = lBits(x)
        g1 = lBits(x + 1)
        b1 = lBits(x + 2)
        TmpCol = (r1 + g1 + b1) \ 3
        lBits(x) = TmpCol
        lBits(x + 1) = TmpCol
        lBits(x + 2) = TmpCol
    Next x

    '  Paint it!
    If bIsIcon Then
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0  ' Get the mask
        StretchDIBits lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd    ' Draw the mask
        PaintGrayScale = StretchDIBits(lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)  'Draw the gray
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    Else
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0  ' Get the mask
        StretchDIBits lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd    ' Draw the mask
        PaintGrayScale = StretchDIBits(lhdc, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    End If

    '   Clear memory
    DeleteDC TmpDC

End Function

Private Sub DrawLineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal color As Long)

'****************************************************************************
'*  draw lines
'****************************************************************************

Dim pt      As POINT
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, 1, color)
    hPenOld = SelectObject(UserControl.hdc, hPen)
    MoveToEx UserControl.hdc, X1, Y1, pt
    LineTo UserControl.hdc, X2, Y2
    SelectObject UserControl.hdc, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld

End Sub


Private Function BlendColorEx(Color1 As Long, Color2 As Long, Optional Percent As Long) As Long

'   Combines two colors together by how many percent.
'   Inspired from dcbutton (honestly not copied!!) hehe

Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim r3 As Long, g3 As Long, b3 As Long

    If Percent <= 0 Then Percent = 0
    If Percent >= 100 Then Percent = 100
    
    r1 = Color1 And 255
    g1 = (Color1 \ 256) And 255
    b1 = (Color1 \ 65536) And 255
    
    r2 = Color2 And 255
    g2 = (Color2 \ 256) And 255
    b2 = (Color2 \ 65536) And 255
    
    r3 = r1 + (r1 - r2) * Percent \ 100
    g3 = g1 + (g1 - g2) * Percent \ 100
    b3 = b1 + (b1 - b2) * Percent \ 100
    
    BlendColorEx = r3 + 256& * g3 + 65536 * b3
    
End Function

Private Function BlendColors(ByVal lBackColorFrom As Long, ByVal lBackColorTo As Long) As Long

'***************************************************************************
'*  Combines (mix) two colors                                              *
'*  This is amother method in which you can't specify percentage
'***************************************************************************

    BlendColors = RGB(((lBackColorFrom And &HFF) + (lBackColorTo And &HFF)) / 2, (((lBackColorFrom \ &H100) And &HFF) + ((lBackColorTo \ &H100) And &HFF)) / 2, (((lBackColorFrom \ &H10000) And &HFF) + ((lBackColorTo \ &H10000) And &HFF)) / 2)

End Function

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal color As Long)

'****************************************************************************
'*  Draws a rectangle specified by coords and color of the rectangle        *
'****************************************************************************

Dim brect As RECT
Dim hBrush As Long
Dim ret As Long

    brect.Left = x
    brect.Top = y
    brect.Right = x + Width
    brect.Bottom = y + Height

    hBrush = CreateSolidBrush(color)

    ret = FrameRect(hdc, brect, hBrush)

    ret = DeleteObject(hBrush)

End Sub

Private Sub DrawFocusRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)

'****************************************************************************
'*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
'****************************************************************************

Dim brect As RECT
Dim RetVal As Long

    brect.Left = x
    brect.Top = y
    brect.Right = x + Width
    brect.Bottom = y + Height

    RetVal = DrawFocusRect(hdc, brect)

End Sub

Private Sub DrawGradientEx(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As GradientDirectionCts)

'****************************************************************************
'* Draws very fast Gradient in four direction.                              *
'* Author: Carles P.V (Gradient Master)                                     *
'* This routine works as a heart for this control.                          *
'* Thank you so much Carles.                                                *
'****************************************************************************

Dim uBIH    As BITMAPINFOHEADER
Dim lBits() As Long
Dim lGrad() As Long

Dim r1      As Long
Dim g1      As Long
Dim b1      As Long
Dim r2      As Long
Dim g2      As Long
Dim b2      As Long
Dim dR      As Long
Dim dG      As Long
Dim dB      As Long

Dim Scan    As Long
Dim i       As Long
Dim iEnd    As Long
Dim iOffset As Long
Dim j       As Long
Dim jEnd    As Long
Dim iGrad   As Long

'-- A minor check

    'If (Width < 1 Or Height < 1) Then Exit Sub
    If (Width < 1 Or Height < 1) Then
        Exit Sub
    End If

    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    r1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    g1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    b1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    r2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    g2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    b2 = Color2 Mod &H100&

    '-- Get color distances
    dR = r2 - r1
    dG = g2 - g1
    dB = b2 - b1

    '-- Size gradient-colors array
    Select Case GradientDirection
    Case [gdHorizontal]
        ReDim lGrad(0 To Width - 1)
    Case [gdVertical]
        ReDim lGrad(0 To Height - 1)
    Case Else
        ReDim lGrad(0 To Width + Height - 2)
    End Select

    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
    Else
        For i = 0 To iEnd
            lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
        Next i
    End If

    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width

    '-- Render gradient DIB
    Select Case GradientDirection

    Case [gdHorizontal]

        For j = 0 To jEnd
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(i - iOffset)
            Next i
            iOffset = iOffset + Scan
        Next j

    Case [gdVertical]

        For j = jEnd To 0 Step -1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(j)
            Next i
            iOffset = iOffset + Scan
        Next j

    Case [gdDownwardDiagonal]

        iOffset = jEnd * Scan
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset - Scan
            iGrad = j
        Next j

    Case [gdUpwardDiagonal]

        iOffset = 0
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset + Scan
            iGrad = j
        Next j
    End Select

    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With

    '-- Paint it!
    StretchDIBits UserControl.hdc, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy

End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long

'****************************************************************************
'*  System color code to long rgb                                           *
'****************************************************************************

    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

Private Sub RedrawButton()

'****************************************************************************
'*  The main routine of this usercontrol. Everything is drawn here.         *
'****************************************************************************

    UserControl.Cls                                'Clears usercontrol
    lh = ScaleHeight
    lw = ScaleWidth

    SetRect m_ButtonRect, 0, 0, lw, lh             'Sets the button rectangle

    If (m_bCheckBoxMode) Then                      'If Checkboxmode True
        If Not (m_ButtonStyle = eStandard Or m_ButtonStyle = eXPToolbar Or m_ButtonStyle = eVisualStudio) Then
            If m_bValue Then m_Buttonstate = eStateDown
        End If
    End If

    Select Case m_ButtonStyle

    Case eStandard
        DrawStandardButton m_Buttonstate
    Case e3DHover
        DrawStandardButton m_Buttonstate
    Case eFlat
        DrawStandardButton m_Buttonstate
    Case eFlatHover
        DrawStandardButton m_Buttonstate
    Case eWindowsXP
        DrawWinXPButton m_Buttonstate
    Case eXPToolbar
        DrawXPToolbar m_Buttonstate
    Case eGelButton
        DrawGelButton m_Buttonstate
    Case eAOL
        DrawAOLButton m_Buttonstate
    Case eInstallShield
        DrawInstallShieldButton m_Buttonstate
    Case eVistaAero
        DrawVistaButton m_Buttonstate
    Case eVistaToolbar
        DrawVistaToolbarStyle m_Buttonstate
    Case eVisualStudio
        DrawVisualStudio2005 m_Buttonstate
    Case eOutlook2007
        DrawOutlook2007 m_Buttonstate
    Case eOffice2003
        DrawOffice2003 m_Buttonstate
    Case eAqua
        DrawAquaButton m_Buttonstate
    Case eSleek
        DrawSleekButton m_Buttonstate
    End Select

    DrawPicwithCaption

End Sub

Private Sub CreateRegion()

'***************************************************************************
'*  Create region everytime you redraw a button.                           *
'*  Because some settings may have changed the button regions              *
'***************************************************************************
    
    If m_lButtonRgn Then DeleteObject m_lButtonRgn
    Select Case m_ButtonStyle
    Case eWindowsXP, eVistaAero, eVistaToolbar, eInstallShield
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 3, 3)
    Case eGelButton, eXPToolbar
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 4, 4)
    Case eAqua
        m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 18, 18)
    Case Else
        m_lButtonRgn = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    End Select
    SetWindowRgn UserControl.hWnd, m_lButtonRgn, True       'Set Button Region
    DeleteObject m_lButtonRgn                               'Free memory

End Sub

Private Sub DrawPicwithCaption()

'****************************************************************************
'* Draws a Picture in Enabled / Disabled mode along with Caption            *
'* Also captions are drawn here calculating all rects                       *
'* Routine to make GrayScale images is the work of Jim Jose.                *
'****************************************************************************

'=========================================================================
'* This routine resulted longer and difficult to read due to 9 picture
'* alignments along with 3 caption alignments.
'* And with Picture Normal and Picture Over features along with
'* UseMaskColor option

'* But it's still fast enough!
'=========================================================================

Dim PicX     As Long                       'X position of picture
Dim PicY     As Long                       'Y Position of Picture
Dim tmpPic      As New StdPicture          'Temp picture (Normal)
Dim tmpPicOver  As New StdPicture          'Over picture
Dim hDCSrc   As Long
Dim hOldob   As Long

Dim lpRect   As RECT                      'RECT to draw caption
Dim CaptionW As Long                      'Width of Caption
Dim CaptionH As Long                      'Height of Caption
Dim CaptionX As Long                      'Left of Caption
Dim CaptionY As Long                      'Top of Caption

    lw = ScaleWidth                          'Height of Button
    lh = ScaleHeight                         'Width of Button

    '  Get the Caption's height and Width
    CaptionW = TextWidth(m_Caption)         'Caption's Width
    CaptionH = TextHeight(m_Caption)        'Caption's Height
      
    '  Copy the original picture into a temp var
    Set tmpPic = m_Picture
    Set tmpPicOver = m_PicOver
    
    ' --Adjust Picture Sizes
    Select Case m_PicSize
    Case epsNormal
        PicSizeH = ScaleX(tmpPic.Height, vbHimetric, vbPixels)
        PicSizeW = ScaleX(tmpPic.Width, vbHimetric, vbPixels)
    Case eps16x16
        PicSizeH = 16
        PicSizeW = 16
    Case eps24x24
        PicSizeH = 24
        PicSizeW = 24
    Case eps32x32
        PicSizeH = 32
        PicSizeW = 32
    End Select
    
    ' --This button not allows different types for different states.
    ' --I could do this, but it makes great overload and waste f time
    ' --See the coding at the end of routine :(
    If tmpPic.Type <> tmpPicOver.Type Then
        Set tmpPicOver = m_Picture
    End If
    
    '=========================================================================
    ' --Set text pos and Pic pos
    Select Case m_PictureAlign
    Case epLeftOfCaption
        PicX = (lw - (PicSizeW + CaptionW)) \ 2
        If PicX < 4 Then PicX = 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2 - CaptionW \ 2) + (PicSizeW \ 2) + 3 'Some distance of 3
        If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8  'Text shouldn't draw over picture
        CaptionY = (lh \ 2 - CaptionH \ 2)

    Case epLeftEdge
        PicX = 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2) - (CaptionW \ 2) + (PicSizeW \ 2)
        If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8  'Text shouldn't draw over picture
        CaptionY = (lh \ 2 - CaptionH \ 2)

    Case epRightEdge

        PicX = lw - PicSizeW - 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw - CaptionW - 4) - PicSizeW
        CaptionY = (lh \ 2 - CaptionH \ 2)

    Case epRightOfCaption

        PicX = (lw - (PicSizeW - CaptionW)) \ 2
        If PicX > (lw - PicSizeW - 4) Then PicX = lw - PicSizeW - 4
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2 - CaptionW \ 2) - (PicSizeW \ 2) - 3
        If CaptionX + CaptionW < CaptionW Then
            CaptionX = (lw - CaptionW - 4) - PicSizeW
        End If
        CaptionY = lh \ 2 - (CaptionH \ 2)

    Case epCenter
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - PicSizeH) \ 2
        CaptionX = (lw \ 2) - (CaptionW \ 2)
        CaptionY = (lh \ 2) - CaptionH \ 2

    Case epBottomEdge
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - PicSizeH) - 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2

    Case epBottomOfCaption
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - (PicSizeH - CaptionH)) \ 2
        If PicY > lh - PicSizeH - 4 Then PicY = lh - PicSizeH - 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2

    Case epTopEdge
        PicX = (lw - PicSizeW) \ 2
        PicY = 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2

    Case epTopOfCaption
        PicX = (lw - PicSizeW) \ 2
        PicY = (lh - (PicSizeH + CaptionH)) \ 2
        If PicY < 4 Then PicY = 4
        CaptionX = (lw \ 2 - CaptionW \ 2)
        CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2

    End Select

    ' --Minor check if picture's size exceeds button size
    If PicX < 1 Then PicX = 1
    If PicY < 1 Then PicY = 1
    If PicX + PicSizeW > ScaleWidth Then PicSizeW = ScaleWidth - 8
    If PicY + PicSizeH > ScaleHeight Then PicSizeH = ScaleHeight - 4

    '========================================================================
    ' --Calculate caption rects with Caption Alignment
    If m_Picture Is Nothing Then
        ' --Calculate caption rects if no picture available
        Select Case m_CaptionAlign
        Case ecLeftAlign
            CaptionX = 4
        Case ecCenterAlign
            CaptionX = (lw \ 2) - (CaptionW \ 2)
        Case ecRightAlign
            CaptionX = (lw - CaptionW - 4)
        End Select
        CaptionY = (lh \ 2) - (CaptionH \ 2)
        PicX = 0
        PicY = 0
    Else
        ' --There is a picture, so calc rects with that too.. (depending on Picture Align)
        Select Case m_CaptionAlign
        Case ecLeftAlign
            If m_PictureAlign = epLeftEdge Then
                CaptionX = PicSizeW + 8
            ElseIf m_PictureAlign = epLeftOfCaption Then
                PicX = 4
                CaptionX = PicX + PicSizeW + 4
            ElseIf m_PictureAlign = epRightEdge Then
                If CaptionX < 4 Then
                    CaptionX = (lw - CaptionW - 4) - PicSizeW
                Else
                    CaptionX = 4
                End If
            ElseIf m_PictureAlign = epRightOfCaption Then
                CaptionX = 4
                PicX = CaptionW + 4
            Else
            CaptionX = 4
            End If
        Case ecRightAlign
            If m_PictureAlign = epRightEdge Then
                CaptionX = (lw - CaptionW - 4) - PicSizeW
            ElseIf m_PictureAlign = epRightOfCaption Then
                PicX = lw - PicSizeW - 4
                CaptionX = (lw - CaptionW - 4) - PicSizeW
            ElseIf m_PictureAlign = epLeftEdge Then
                CaptionX = (lw - CaptionW - 4)
                If CaptionX < PicSizeW + 4 Then
                    CaptionX = PicSizeW + 4
                End If
            ElseIf m_PictureAlign = epLeftOfCaption Then
                CaptionX = (lw - CaptionW - 4)
                PicX = CaptionX - PicSizeW - 4
            Else
                CaptionX = (lw - CaptionW - 4)
            End If
        Case ecCenterAlign
            If m_PictureAlign = epRightEdge Then
                If CaptionX + CaptionW < CaptionW Then
                    CaptionX = (lw - CaptionW - 4) - PicSizeW
                Else
                    CaptionX = (lw \ 2) - (CaptionW \ 2)
                End If
            End If
        End Select
    End If

    ' --Uncomment the below lines and see what happens!! Oops
    ' --The caption draws awkwardly with accesskeys!
    If UserControl.AccessKeys <> vbNullString Then
        CaptionX = CaptionX + 3
    End If

    '=========================================================================
    ' --Adjust Picture Positions
    Select Case m_ButtonStyle
    Case eStandard, eFlat, eVistaToolbar, eXPToolbar
        If m_Buttonstate = eStateDown Then
            PicX = PicX + 1
            PicY = PicY + 1
        End If
    Case eAOL
        If m_Buttonstate = eStateDown Then
            PicX = PicX + 2         'More depth for AOL
            PicY = PicY + 2
        Else
            PicX = PicX - 1         'For AOL
            PicY = PicY - 1
        End If
    End Select

    '=========================================================================
    ' --We have calculated all rects . SO set that rects here
    ' --If picture available, Set text rects with Picture
    If m_Buttonstate = eStateDown Then
        Select Case m_ButtonStyle
        Case eStandard, eFlat, eVistaToolbar, eXPToolbar
            ' --Caption pos for Standard/Flat buttons on down state
            SetRect lpRect, CaptionX + 1, CaptionY + 1, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
        Case eAOL
            ' --Caption RECT for AOL buttons
            SetRect lpRect, CaptionX + 1, CaptionY + 2, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
        Case eAqua
            SetRect lpRect, CaptionX, CaptionY - 1, CaptionW + CaptionX, CaptionH + CaptionY - 1
        Case Else
            ' --for other buttons on down state
            SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
        End Select
    Else
        Select Case m_ButtonStyle
        Case eAOL
            SetRect lpRect, CaptionX - 2, CaptionY - 2, CaptionW + CaptionX - 2, CaptionH + CaptionY - 2
        Case eAqua
            SetRect lpRect, CaptionX, CaptionY - 1, CaptionW + CaptionX, CaptionH + CaptionY - 1
        Case Else
            SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
            ' --For drawing Focus rect exactly around Caption
            SetRect m_CapRect, CaptionX - 2, CaptionY, CaptionW + CaptionX + 1, CaptionH + CaptionY + 1
        End Select
    End If

    '=======================================================================
    ' --Draw Picture Enabled/Disabled depending of Pic type and button state
    Select Case tmpPic.Type
    Case vbPicTypeIcon

        If m_bEnabled Then
            Select Case m_Buttonstate
            Case eStateNormal
                DrawIconEx hdc, PicX, PicY, tmpPic.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
            Case eStateOver
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            Case eStateDown
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            End Select
            
            ' --For checkboxmode
            If m_bCheckBoxMode And m_bValue Then
                If m_PicOver Is Nothing Then
                    DrawIconEx hdc, PicX, PicY, tmpPic.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hdc, PicX, PicY, tmpPicOver.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
                End If
            End If
            
        Else
            ' --Draw grayed picture (Thanks to Jim Jose)
            PaintGrayScale hdc, tmpPic.Handle, PicX, PicY, PicSizeW, PicSizeH
        End If

    Case vbPicTypeBitmap
        If m_bEnabled Then
            ' --Draw picture with Maskcolor
            hDCSrc = CreateCompatibleDC(0)
                
            Select Case m_Buttonstate
            Case eStateNormal
                hOldob = SelectObject(hDCSrc, tmpPic.Handle)
            Case eStateOver
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hDCSrc, tmpPic.Handle)
                Else
                    hOldob = SelectObject(hDCSrc, tmpPicOver.Handle)
                End If
            Case eStateDown
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hDCSrc, tmpPic.Handle)
                Else
                    hOldob = SelectObject(hDCSrc, tmpPicOver.Handle)
                End If
            End Select
                
            If m_bCheckBoxMode And m_bValue Then
                If m_PicOver Is Nothing Then
                    hOldob = SelectObject(hDCSrc, tmpPic.Handle)
                Else
                    hOldob = SelectObject(hDCSrc, tmpPicOver.Handle)
                End If
            End If
                
            If m_bUseMaskColor Then
                ' -Create Trans areas
                TransparentBlt UserControl.hdc, PicX, PicY, PicSizeW, PicSizeH, hDCSrc, 0, 0, PicSizeW, PicSizeH, m_lMaskColor
                SelectObject hDCSrc, hOldob
                DeleteDC hDCSrc
            Else
                ' --Simply draw picture
                StretchBlt hdc, PicX, PicY, PicSizeW, PicSizeH, hDCSrc, 0, 0, PicSizeW, PicSizeH, vbSrcCopy
                SelectObject hDCSrc, hOldob
                DeleteDC hDCSrc
            End If
        Else
            ' --Disabled Bitmap (Thanks to Jim Jose.)
            PaintGrayScale hdc, tmpPic.Handle, PicX, PicY, PicSizeW, PicSizeH
        End If
    End Select
    
    '=========================================================================
    ' --At last, draw text
    SetTextColor hdc, IIf(m_bEnabled, TranslateColor(m_bColors.tForeColor), TranslateColor(vbGrayText))
    If Not m_WindowsNT Then
        ' --Unicode not supported
        DrawText hdc, m_Caption, Len(m_Caption), lpRect, DT_SINGLELINE  'Button looks good in SingleLine!
    Else
        ' --Supports Unicode (i.e above Windows NT)
        DrawTextW hdc, StrPtr(m_Caption), Len(m_Caption), lpRect, DT_SINGLELINE
    End If
    
    ' --Clear memory
    Set tmpPic = Nothing
    Set tmpPicOver = Nothing
    
End Sub

Private Sub SetAccessKey()

Dim i As Long

    UserControl.AccessKeys = ""
    If Len(m_Caption) > 1 Then
        i = InStr(1, m_Caption, "&", vbTextCompare)
        If (i < Len(m_Caption)) And (i > 0) Then
            If Mid$(m_Caption, i + 1, 1) <> "&" Then
                AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
            Else
                i = InStr(i + 2, m_Caption, "&", vbTextCompare)
                If Mid$(m_Caption, i + 1, 1) <> "&" Then
                    AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
                End If
            End If
        End If
    End If

End Sub

Private Sub DrawCorners(color As Long)

'****************************************************************************
'* Draws four Corners of the button specified by Color                      *
'****************************************************************************

    With UserControl
        lh = .ScaleHeight
        lw = .ScaleWidth

        SetPixel .hdc, 1, 1, color
        SetPixel .hdc, 1, lh - 2, color
        SetPixel .hdc, lw - 2, 1, color
        SetPixel .hdc, lw - 2, lh - 2, color

    End With

End Sub

Private Sub DrawStandardButton(ByVal vState As enumButtonStates)

'****************************************************************************
' Draws  four different styles in one procedure                             *
' Makes reading the code difficult, but saves much space!! ;)               *
'****************************************************************************

Dim FocusRect   As RECT
Dim tmpRect     As RECT

    lh = ScaleHeight
    lw = ScaleWidth
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        '     Draws raised edge border
        DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
    End If

    If m_bCheckBoxMode And m_bValue Then
        PaintRect ShiftColor(TranslateColor(m_bColors.tBackColor), 0.02), m_ButtonRect
        If m_ButtonStyle <> eFlatHover Then
            DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
            If m_bShowFocus And m_bHasFocus And m_ButtonStyle = eStandard Then
                DrawRectangle 4, 4, lw - 7, lh - 7, TranslateColor(vbApplicationWorkspace)
            End If
        End If
        Exit Sub
    End If

    Select Case vState
    Case eStateNormal
        CreateRegion
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        ' --Draws flat raised edge border
        Select Case m_ButtonStyle
        Case eStandard
            DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
        Case eFlat
            DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
        End Select
    Case eStateOver
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        Select Case m_ButtonStyle
        Case eFlatHover, eFlat
            ' --Draws flat raised edge border
            DrawEdge hdc, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
        Case Else
            ' --Draws 3d raised edge border
            DrawEdge hdc, m_ButtonRect, BDR_RAISED95, BF_RECT
        End Select

    Case eStateDown
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        Select Case m_ButtonStyle
        Case eStandard
            DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&H99A8AC)
            DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
        Case e3DHover
            DrawEdge hdc, m_ButtonRect, BDR_SUNKEN95, BF_RECT
        Case eFlatHover, eFlat
            ' --Draws flat pressed edge
            DrawRectangle 0, 0, lw, lh, TranslateColor(vbWhite)
            DrawRectangle 0, 0, lw + 1, lh + 1, TranslateColor(vbGrayText)
        End Select
    End Select

    ' --Button has focus but not downstate Or button is Default
        If m_bHasFocus Or m_bDefault Then
            If m_bShowFocus And Ambient.UserMode Then
                If m_ButtonStyle = e3DHover Or m_ButtonStyle = eStandard Then
                    SetRect FocusRect, 4, 4, lw - 4, lh - 4
                Else
                    SetRect FocusRect, 3, 3, lw - 3, lh - 3
                End If
                If m_bParentActive Then
                    DrawFocusRect hdc, FocusRect
                End If
            End If
            If vState <> eStateDown And m_ButtonStyle = eStandard Then
                SetRect tmpRect, 0, 0, lw - 1, lh - 1
                DrawEdge hdc, tmpRect, BDR_RAISED95, BF_RECT
                DrawRectangle 0, 0, lw - 1, lh - 1, TranslateColor(vbApplicationWorkspace)
                DrawRectangle 0, 0, lw, lh, TranslateColor(vbBlack)
            End If
        End If

End Sub

Private Sub DrawXPToolbar(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    UserControl.BackColor = Ambient.BackColor
    bColor = TranslateColor(m_bColors.tBackColor)

    
    If m_bCheckBoxMode And m_bValue Then
        ' --Check with XP Toolbar!
        If m_bIsDown Then vState = eStateDown
    End If
    
    If m_bCheckBoxMode And m_bValue And vState <> eStateDown Then
        SetRect lpRect, 0, 0, lw, lh
        PaintRect TranslateColor(&HFEFEFE), lpRect
        m_bColors.tForeColor = TranslateColor(vbButtonText)
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HAF987A)
        DrawCorners ShiftColor(TranslateColor(&HC1B3A0), -0.2)
        If vState = eStateOver Then
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HEDF0F2)  'Right Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HD8DEE4)   'Bottom
            DrawLineApi 1, lh - 3, lw - 1, lh - 3, TranslateColor(&HE8ECEF)  'Bottom
            DrawLineApi 1, lh - 4, lw - 1, lh - 4, TranslateColor(&HF8F9FA)   'Bottom
        End If
    Exit Sub
    End If

    Select Case vState
    Case eStateNormal
        CreateRegion
        PaintRect bColor, m_ButtonRect
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(&HFDFEFE), TranslateColor(&HEEF4F4), gdVertical
        DrawGradientEx 0, lh / 2, lw, lh / 2, TranslateColor(&HEEF4F4), TranslateColor(&HEAF1F1), gdVertical
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE0E7EA) 'right line
        DrawLineApi lw - 3, 2, lw - 3, lh - 2, TranslateColor(&HEAF0F0)
        DrawLineApi 0, lh - 4, lw, lh - 4, TranslateColor(&HE5EDEE)    'Bottom
        DrawLineApi 0, lh - 3, lw, lh - 3, TranslateColor(&HD6E1E4)    'Bottom
        DrawLineApi 0, lh - 2, lw, lh - 2, TranslateColor(&HC6D2D7)    'Bottom
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HC3CECE)
        DrawCorners ShiftColor(TranslateColor(&HC9D4D4), -0.05)
    Case eStateDown
        PaintRect TranslateColor(&HDDE4E5), m_ButtonRect                 'Paint with Darker color
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(TranslateColor(&HD1DADC), -0.02)          'Topmost Line
        DrawLineApi 1, 2, lw - 2, 2, ShiftColor(TranslateColor(&HDAE1E3), -0.02)          'A lighter top line
        DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(TranslateColor(&HDEE5E6), 0.02) 'Bottom Line
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(TranslateColor(&HE5EAEB), 0.02)
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H929D9D)
        DrawCorners ShiftColor(TranslateColor(&HABB4B5), -0.2)
    End Select

    If vState = eStateDown Then
        m_bColors.tForeColor = TranslateColor(vbWhite)
    Else
        m_bColors.tForeColor = TranslateColor(vbButtonText)
    End If

End Sub

Private Sub DrawWinXPButton(ByVal vState As enumButtonStates)

'****************************************************************************
'* Windows XP Button                                                        *
'* I made this in just 4 hours                                              *
'* Totally written from Scratch and coded by Me!!                           *
'****************************************************************************

Dim lpRect As RECT
Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        CreateRegion
        PaintRect ShiftColor(bColor, 0.03), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.1)
        DrawCorners ShiftColor(bColor, 0.2)
        Exit Sub
    End If

    Select Case vState

    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
        DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.09) 'BottomMost line
        DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(bColor, -0.05) 'Bottom Line
        DrawLineApi 1, lh - 4, lw - 2, lh - 4, ShiftColor(bColor, -0.01) 'Bottom Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, -0.08) 'Right Line
        DrawLineApi 1, 1, 1, lh - 2, BlendColors(TranslateColor(vbWhite), (bColor)) 'Left Line

    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
        DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
        DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&H89D8FD)           'uppermost inner hover
        DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HCFF0FF)           'uppermost outer hover
        DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&H49BDF9)           'Leftmost Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&H49BDF9) 'Rightmost Line
        DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&H7AD2FC)           'Left Line
        DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&H7AD2FC) 'Right Line
        DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&H30B3F8) 'BottomMost Line
        DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&H97E5&)  'Bottom Line

    Case eStateDown
        PaintRect ShiftColor(bColor, -0.05), m_ButtonRect               'Paint with Darker color
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(bColor, -0.16)          'Topmost Line
        DrawLineApi 1, 2, lw - 2, 2, ShiftColor(bColor, -0.1)          'A lighter top line
        DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, 0.07) 'Bottom Line
        DrawLineApi 1, 1, 1, lh - 2, ShiftColor(bColor, -0.16)  'Leftmost Line
        DrawLineApi 2, 2, 2, lh - 2, ShiftColor(bColor, -0.1)   'Left1 Line
        DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, 0.04) 'Right Line

    End Select
    
    If m_bParentActive Then
        If (m_bHasFocus Or m_bDefault) And (vState <> eStateDown And vState <> eStateOver) Then
            DrawLineApi 1, 2, lw - 2, 2, TranslateColor(&HF6D4BC)           'uppermost inner hover
            DrawLineApi 1, 1, lw - 2, 1, TranslateColor(&HFFE7CE)           'uppermost outer hover
            DrawLineApi 1, 1, 1, lh - 2, TranslateColor(&HE6AF8E)           'Leftmost Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, TranslateColor(&HE6AF8E) 'Rightmost Line
            DrawLineApi 2, 2, 2, lh - 3, TranslateColor(&HF4D1B8)           'Left Line
            DrawLineApi lw - 3, 3, lw - 3, lh - 3, TranslateColor(&HF4D1B8) 'Right Line
            DrawLineApi 2, lh - 3, lw - 2, lh - 3, TranslateColor(&HE4AD89) 'BottomMost Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, TranslateColor(&HEE8269) 'Bottom Line
        End If
    End If

    On Error Resume Next
    If m_bParentActive Then
        If m_bShowFocus And m_bParentActive And (m_bHasFocus Or m_bDefault) Then  'show focusrect at runtime only
            SetRect lpRect, 2, 2, lw - 2, lh - 2     'I don't like this ugly focusrect!!
            DrawFocusRect hdc, lpRect
        End If
    End If
    
    DrawRectangle 0, 0, lw, lh, TranslateColor(&H743C00)
    DrawCorners ShiftColor(TranslateColor(&H743C00), 0.3)

End Sub

Private Sub DrawVisualStudio2005(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim bColor As Long

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), TranslateColor(vbWhite)), bColor, gdVertical
    End If
    
    If m_bCheckBoxMode And m_bValue Then
        PaintRect TranslateColor(&HE8E6E1), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(&H6F4B4B), 0.05)
        If vState = eStateOver Then
            PaintRect TranslateColor(&HE2B598), m_ButtonRect
            DrawRectangle 0, 0, lw, lh, TranslateColor(&HC56A31)
        End If
    Exit Sub
    End If

    Select Case vState

    Case eStateNormal
        DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), TranslateColor(vbWhite)), bColor, gdVertical
    Case eStateOver
        PaintRect TranslateColor(&HEED2C1), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HC56A31)
    Case eStateDown
        PaintRect TranslateColor(&HE2B598), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H6F4B4B)
    End Select

End Sub

Private Sub DrawAOLButton(ByVal vState As enumButtonStates)

'****************************************************************************
'* AOL (American Online) buttons.                                           *
'****************************************************************************

Dim lpRect As RECT
Dim FocusRect As RECT
Dim bColor As Long

    bColor = TranslateColor(m_bColors.tBackColor)

    If Not m_bEnabled Then                   'Draw Disabled button
        bColor = ShiftColor(TranslateColor(m_bColors.tBackColor), 0.1)
    End If

    Select Case vState
    Case eStateNormal
        CreateRegion
        On Error GoTo h:
        UserControl.BackColor = Ambient.BackColor  'Transparent?!?
        
        ' --Shadows
        DrawRectangle 6, 6, lw - 9, lh - 9, TranslateColor(&H808080)
        DrawRectangle 5, 5, lw - 7, lh - 7, TranslateColor(&HA0A0A0)
        DrawRectangle 4, 4, lw - 5, lh - 5, TranslateColor(&HC0C0C0)

        SetRect lpRect, 0, 0, lw - 5, lh - 5
        PaintRect bColor, lpRect

        DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)

    Case eStateOver
        UserControl.BackColor = Ambient.BackColor

        ' --Shadows
        DrawRectangle 6, 6, lw - 9, lh - 9, TranslateColor(&H808080)
        DrawRectangle 5, 5, lw - 7, lh - 7, TranslateColor(&HA0A0A0)
        DrawRectangle 4, 4, lw - 5, lh - 5, TranslateColor(&HC0C0C0)

        SetRect lpRect, 0, 0, lw - 5, lh - 5
        PaintRect bColor, lpRect

        DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)

    Case eStateDown
        UserControl.BackColor = Ambient.BackColor

        SetRect lpRect, 3, 3, lw, lh
        PaintRect bColor, lpRect

        DrawRectangle 3, 3, lw - 3, lh - 3, ShiftColor(bColor, 0.3)

    End Select
    
    If m_bParentActive Then
        If m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
            UserControl.DrawMode = 6        'For exact AOL effect
            If vState = eStateDown Then
                SetRect lpRect, 6, 6, lw - 3, lh - 3
            Else
                SetRect lpRect, 3, 3, lw - 6, lh - 6
            End If
            DrawFocusRect hdc, lpRect
        End If
    End If
h:
    'Client Site not available (Error in Ambient.BackColor) rarely occurs

End Sub

Private Sub DrawInstallShieldButton(ByVal vState As enumButtonStates)

'****************************************************************************
'* I saw this style while installing JetAudio in my PC.                     *
'* I liked it, so I implemented and gave it a name 'InstallShield'          *
'* hehe .....
'****************************************************************************

Dim FocusRect As RECT
Dim lpRect As RECT

    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        vState = eStateNormal                 'Simple draw normal state for Disabled
    End If

    Select Case vState
    Case eStateNormal
        CreateRegion
        SetRect m_ButtonRect, 0, 0, lw, lh 'Maybe have changed before!

        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
    Case eStateOver

        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, TranslateColor(m_bColors.tBackColor), TranslateColor(m_bColors.tBackColor), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.2)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.25)
    Case eStateDown

        ' --draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, TranslateColor(vbWhite), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1), ShiftColor(TranslateColor(m_bColors.tBackColor), -0.05), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, TranslateColor(vbWhite)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.23)
        DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), -0.1)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(TranslateColor(m_bColors.tBackColor), -0.4)

    End Select

    DrawCorners ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)

    If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
        DrawFocusRect hdc, m_CapRect
    End If

End Sub

Private Sub DrawGelButton(ByVal vState As enumButtonStates)

'****************************************************************************
' Draws a Gelbutton                                                         *
'****************************************************************************

Dim lpRect    As RECT                              'RECT to fill regions
Dim bColor    As Long                              'Original backcolor

    lh = ScaleHeight
    lw = ScaleWidth

    bColor = TranslateColor(m_bColors.tBackColor)
    Select Case vState

    Case eStateNormal                                'Normal State

        CreateRegion

        ' --Fill the button region with background color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect bColor, lpRect

        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.1)), gdVertical

        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.33)

    Case eStateOver
        ' --Fill the button region with background color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, 0.05), lpRect

        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(ShiftColor(bColor, 0.05), TranslateColor(vbWhite)), 0.15), ShiftColor(bColor, 0.05), gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, bColor, BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.15)), gdVertical

        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.28)

    Case eStateDown

        ' --fill the button region with background color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, -0.03), lpRect

        ' --Make a shining Upper Light
        DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, TranslateColor(vbWhite)), 0.1), bColor, gdVertical
        DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.08), BlendColors(TranslateColor(vbWhite), ShiftColor(bColor, 0.07)), gdVertical

        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.36)

    End Select

    DrawCorners ShiftColor(bColor, -0.5)

End Sub

Private Sub DrawVistaToolbarStyle(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim FocusRect As RECT

    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        ' --Draw Disabled button
        PaintRect TranslateColor(m_bColors.tBackColor), m_ButtonRect
        DrawCorners TranslateColor(m_bColors.tBackColor)
        Exit Sub
    End If

    If vState = eStateNormal Then
        CreateRegion
        ' --Set the rect to fill back color
        SetRect lpRect, 0, 0, lw, lh
        ' --Simply fill the button with one color (No gradient effect here!!)
        PaintRect TranslateColor(m_bColors.tBackColor), lpRect

    ElseIf vState = eStateOver Then

        ' --Draws a gradient effect with the folowing colors
        DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HFDF9F1), TranslateColor(&HF8ECD0), gdVertical

        ' --Draws a gradient in half region to give a Light Effect
        DrawGradientEx 1, lh / 1.7, lw - 2, lh - 2, TranslateColor(&HF8ECD0), TranslateColor(&HF8ECD0), gdVertical

        ' --Draw outside borders
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)

    ElseIf vState = eStateDown Then

        DrawGradientEx 1, 1, lw - 2, lh - 2, TranslateColor(&HF1DEB0), TranslateColor(&HF9F1DB), gdVertical

        ' --Draws outside borders
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HCA9E61)
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite)
    
    End If

    If vState = eStateDown Or vState = eStateOver Then
        DrawCorners ShiftColor(TranslateColor(&HCA9E61), 0.3)
    End If

End Sub


Private Sub DrawVistaButton(ByVal vState As enumButtonStates)

'*************************************************************************
'* Draws a cool Vista Aero Style Button                                  *
'* Use a light background color for best result                          *
'*************************************************************************

Dim lpRect As RECT            'Used to set rect for drawing rectangles
Dim Color1 As Long            'Shifted / Blended color
Dim bColor As Long            'Original back Color

    lh = ScaleHeight
    lw = ScaleWidth
    Color1 = ShiftColor(TranslateColor(m_bColors.tBackColor), 0.05)
    bColor = TranslateColor(m_bColors.tBackColor)

    If Not m_bEnabled Then
        ' --Draw the Disabled Button
        CreateRegion
        ' --Fill the button with disabled color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, 0.03), lpRect

        ' --Draws outside disabled color rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.25)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.25)
        DrawCorners ShiftColor(bColor, -0.03)
    Exit Sub
    End If

    Select Case vState

    Case eStateNormal

        CreateRegion

        ' --Draws a gradient in the full region
        DrawGradientEx 1, 1, lw - 1, lh, Color1, bColor, gdVertical

        ' --Draws a gradient in half region to give a glassy look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, ShiftColor(bColor, -0.02), ShiftColor(bColor, -0.15), gdVertical

        ' --Draws border rectangle
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H707070)   'outer
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite) 'inner

    Case eStateOver

        ' --Make gradient in the upper half region
        DrawGradientEx 1, 1, lw - 2, lh / 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFF3DA), gdVertical
        
        ' --Draw gradient in half button downside to give a glass look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HFFE9C1), TranslateColor(&HFDE1AE), gdVertical
        
        ' --Draws left side gradient effects horizontal
        DrawGradientEx 1, 3, 5, lh / 2 - 2, TranslateColor(&HFFEECD), TranslateColor(&HFFF7E4), gdHorizontal    'Left
        DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, TranslateColor(&HFAD68F), ShiftColor(TranslateColor(&HFDE1AC), 0.01), gdHorizontal   'Left
        
        ' --Draws right side gradient effects horizontal
        DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, TranslateColor(&HFFF7E4), TranslateColor(&HFFEECD), gdHorizontal 'Right
        DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HFDE1AC), 0.01), TranslateColor(&HFAD68F), gdHorizontal 'Right
        
        ' --Draws border rectangle
        DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)   'outer
        DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(vbWhite) 'inner

    Case eStateDown

        ' --Draw a gradent in full region
        DrawGradientEx 1, 1, lw - 1, lh, TranslateColor(&HF6E4C2), TranslateColor(&HF6E4C2), gdVertical
        
        ' --Draw gradient in half button downside to give a glass look
        DrawGradientEx 1, lh / 2, lw - 2, lh - 2, TranslateColor(&HF0D29A), TranslateColor(&HF0D29A), gdVertical
        
        ' --Draws down rectangle
        
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H5C411D)    '
        DrawLineApi 1, 1, lw - 1, 1, TranslateColor(&HB39C71)   '\Top Lines
        DrawLineApi 1, 2, lw - 1, 2, TranslateColor(&HD6C6A9)   '/
        DrawLineApi 1, 3, lw - 1, 3, TranslateColor(&HECD9B9)   '
    
        DrawLineApi 1, 1, 1, lh / 2 - 1, TranslateColor(&HCFB073)   'Left upper
        DrawLineApi 1, lh / 2, 1, lh - (lh / 2) - 1, TranslateColor(&HC5912B)   'Left Bottom
        
        ' --Draws left side gradient effects horizontal
        DrawGradientEx 1, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HE6C891), 0.02), ShiftColor(TranslateColor(&HF6E4C2), -0.01), gdHorizontal   'Left
        DrawGradientEx 1, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HDCAB4E), 0.02), ShiftColor(TranslateColor(&HF0D29A), -0.01), gdHorizontal 'Left
        
        ' --Draws right side gradient effects horizontal
        DrawGradientEx lw - 6, 3, 5, lh / 2 - 2, ShiftColor(TranslateColor(&HF6E4C2), -0.01), ShiftColor(TranslateColor(&HE6C891), 0.02), gdHorizontal 'Right
        DrawGradientEx lw - 6, lh / 2, 5, lh - (lh / 2) - 1, ShiftColor(TranslateColor(&HF0D29A), -0.01), ShiftColor(TranslateColor(&HDCAB4E), 0.02), gdHorizontal 'Right
        
    End Select

    ' --Draw a focus rectangle if button has focus
    
    If m_bParentActive Then
        If (m_bHasFocus Or m_bDefault) And vState = eStateNormal Then
            ' --Draw darker outer rectangle
            DrawRectangle 0, 0, lw, lh, TranslateColor(&HA77532)
            ' --Draw light inner rectangle
            DrawRectangle 1, 1, lw - 2, lh - 2, TranslateColor(&HFBD848)
        End If

        If (m_bShowFocus And m_bHasFocus) Then
            SetRect lpRect, 1.5, 1.5, lw - 2, lh - 2
            DrawFocusRect hdc, lpRect
        End If
    End If

    ' --Create four corners which will be common to all states
    DrawCorners TranslateColor(&HBE965F)

End Sub

Private Sub DrawOutlook2007(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    bColor = TranslateColor(m_bColors.tBackColor)

    If m_bCheckBoxMode And m_bValue Then
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HA9D9FF), TranslateColor(&H6FC0FF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H3FABFF), TranslateColor(&H75E1FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        If m_bMouseInCtl Then
            DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        End If
        Exit Sub
    End If

    Select Case vState
    Case eStateNormal
        PaintRect bColor, m_ButtonRect
        DrawGradientEx 0, 0, lw, lh / 2.7, BlendColors(ShiftColor(bColor, 0.09), TranslateColor(vbWhite)), BlendColors(ShiftColor(bColor, 0.07), bColor), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), bColor, ShiftColor(bColor, 0.03), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&HE1FFFF), TranslateColor(&HACEAFF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H67D7FF), TranslateColor(&H99E4FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh / 2.7, TranslateColor(&H58C1FF), TranslateColor(&H51AFFF), gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), TranslateColor(&H468FFF), TranslateColor(&H5FD3FF), gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    End Select

End Sub

Private Sub DrawAquaButton(ByVal vState As enumButtonStates)

'****************************************************************************
'* This sub is solely written by Fred.CPP                                   *
'* I tried my own aqua button with some other functions                     *
'* But to get accurate results, we have to do pixel to pixel work           *
'* Thank you Fred                                                           *
'****************************************************************************
    
    ' --Checkboxmode ??
    If m_bCheckBoxMode And m_bValue Then
        vState = eStateDown
    End If
    
    If m_bParentActive Then
        If m_bDefault Or m_bHasFocus Then           'MAC OS X draws
            If vState = eStateDown Then             'Hotstate
                vState = eStateDown                 'for focused buttons
            Else
                vState = eStateOver
            End If
        End If
    End If
    
    UserControl.BackColor = Ambient.BackColor
    
    Select Case vState

    Case eStateNormal
        CreateRegion
        DrawAquaNormal
    Case eStateOver
        DrawAquaHot
    Case eStateDown
        DrawAquaDown
    End Select
    
End Sub

Private Sub DrawAquaNormal()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    
    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HEAE7E8, lpRect
    
    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE6E6E6: SetPixel hdc, 8, 0, &HACACAC: SetPixel hdc, 9, 0, &H7A7A7A: SetPixel hdc, 10, 0, &H6C6C6C: SetPixel hdc, 11, 0, &H6B6B6B: SetPixel hdc, 12, 0, &H6F6F6F: SetPixel hdc, 13, 0, &H716F6F: SetPixel hdc, 14, 0, &H727070: SetPixel hdc, 15, 0, &H676866: SetPixel hdc, 16, 0, &H6C6D6B: SetPixel hdc, 17, 0, &H67696A: SetPixel hdc, 5, 1, &HEFEFEF: SetPixel hdc, 6, 1, &H939393: SetPixel hdc, 7, 1, &H676767: SetPixel hdc, 8, 1, &H797979: SetPixel hdc, 9, 1, &HB3B3B3: SetPixel hdc, 10, 1, &HDBDBDB: SetPixel hdc, 11, 1, &HEBEDEE: SetPixel hdc, 12, 1, &HF5F4F6: SetPixel hdc, 13, 1, &HF5F4F6: SetPixel hdc, 14, 1, &HF5F4F6: SetPixel hdc, 15, 1, &HF5F4F6: SetPixel hdc, 16, 1, &HF5F4F6: SetPixel hdc, 17, 1, &HF5F4F6
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE5E5E5: SetPixel hdc, 5, 2, &H737373: SetPixel hdc, 6, 2, &H656565: SetPixel hdc, 7, 2, &H939393: SetPixel hdc, 8, 2, &HDCDCDC: SetPixel hdc, 9, 2, &HE9E9E9: SetPixel hdc, 10, 2, &HF2F1F3: SetPixel hdc, 11, 2, &HF3F2F4: SetPixel hdc, 12, 2, &HF2F1F3: SetPixel hdc, 13, 2, &HF3F2F4: SetPixel hdc, 14, 2, &HF2F1F3: SetPixel hdc, 15, 2, &HF3F2F4: SetPixel hdc, 16, 2, &HF2F1F3: SetPixel hdc, 17, 2, &HF3F2F4: SetPixel hdc, 3, 3, &HEEEEEE: SetPixel hdc, 4, 3, &H717171: SetPixel hdc, 5, 3, &H6C6C6C: SetPixel hdc, 6, 3, &H909090: SetPixel hdc, 7, 3, &HD2D2D2: SetPixel hdc, 8, 3, &HE3E3E3: SetPixel hdc, 9, 3, &HECECEC: SetPixel hdc, 10, 3, &HEDEDED: SetPixel hdc, 11, 3, &HEEEEEE: SetPixel hdc, 12, 3, &HEDEDED: SetPixel hdc, 13, 3, &HEEEEEE: SetPixel hdc, 14, 3, &HEDEDED: SetPixel hdc, 15, 3, &HEEEEEE: SetPixel hdc, 16, 3, &HEDEDED: SetPixel hdc, 17, 3, &HEEEEEE
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H858585: SetPixel hdc, 4, 4, &H686868: SetPixel hdc, 5, 4, &H959595: SetPixel hdc, 6, 4, &HB1B1B1: SetPixel hdc, 7, 4, &HDCDCDC: SetPixel hdc, 8, 4, &HE3E3E3: SetPixel hdc, 9, 4, &HE3E3E3: SetPixel hdc, 10, 4, &HEAEAEA: SetPixel hdc, 11, 4, &HEBEBEB: SetPixel hdc, 12, 4, &HEBEBEB: SetPixel hdc, 13, 4, &HEBEBEB: SetPixel hdc, 14, 4, &HEBEBEB: SetPixel hdc, 15, 4, &HEBEBEB: SetPixel hdc, 16, 4, &HEBEBEB: SetPixel hdc, 17, 4, &HEBEBEB:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCACACA: SetPixel hdc, 3, 5, &H696969: SetPixel hdc, 4, 5, &H949494: SetPixel hdc, 5, 5, &HA6A6A6: SetPixel hdc, 6, 5, &HC5C5C5: SetPixel hdc, 7, 5, &HD8D8D8: SetPixel hdc, 8, 5, &HE0E0E0: SetPixel hdc, 9, 5, &HE1E1E1: SetPixel hdc, 10, 5, &HEAE9EA: SetPixel hdc, 11, 5, &HE7E7E7: SetPixel hdc, 12, 5, &HE9E7E8: SetPixel hdc, 13, 5, &HEBE8EA: SetPixel hdc, 14, 5, &HEAE7E9: SetPixel hdc, 15, 5, &HEBE8EA: SetPixel hdc, 16, 5, &HEAE7E9: SetPixel hdc, 17, 5, &HEBE8EA
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H808080: SetPixel hdc, 3, 6, &H878787: SetPixel hdc, 4, 6, &HA8A8A8: SetPixel hdc, 5, 6, &HB3B3B3: SetPixel hdc, 6, 6, &HC6C6C6: SetPixel hdc, 7, 6, &HDEDEDE: SetPixel hdc, 8, 6, &HE0E0E0: SetPixel hdc, 9, 6, &HE2E2E2: SetPixel hdc, 10, 6, &HE3E2E2: SetPixel hdc, 11, 6, &HE9EAE9: SetPixel hdc, 12, 6, &HE9E8E9: SetPixel hdc, 13, 6, &HEBE8EA: SetPixel hdc, 14, 6, &HEBE8EA: SetPixel hdc, 15, 6, &HEBE8EA: SetPixel hdc, 16, 6, &HEBE8EA: SetPixel hdc, 17, 6, &HEBE8EA
    SetPixel hdc, 1, 7, &HE8E8E8: SetPixel hdc, 2, 7, &H777777: SetPixel hdc, 3, 7, &H9B9B9B: SetPixel hdc, 4, 7, &HB1B1B1: SetPixel hdc, 5, 7, &HB9B9B9: SetPixel hdc, 6, 7, &HC5C5C5: SetPixel hdc, 7, 7, &HD6D6D6: SetPixel hdc, 8, 7, &HE0E0E0: SetPixel hdc, 9, 7, &HE0E0E0: SetPixel hdc, 10, 7, &HE7E7E7: SetPixel hdc, 11, 7, &HE7E7E7: SetPixel hdc, 12, 7, &HE9E9E9: SetPixel hdc, 13, 7, &HEAEAEA: SetPixel hdc, 14, 7, &HEAEAEA: SetPixel hdc, 15, 7, &HEAEAEA: SetPixel hdc, 16, 7, &HEAEAEA: SetPixel hdc, 17, 7, &HEAEAEA
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HC6C6C6: SetPixel hdc, 2, 8, &H7E7E7E: SetPixel hdc, 3, 8, &HABABAB: SetPixel hdc, 4, 8, &HC1C1C1: SetPixel hdc, 5, 8, &HC1C1C1: SetPixel hdc, 6, 8, &HCBCBCB: SetPixel hdc, 7, 8, &HCECECE: SetPixel hdc, 8, 8, &HD5D5D5: SetPixel hdc, 9, 8, &HD8D8D8: SetPixel hdc, 10, 8, &HDADADA: SetPixel hdc, 11, 8, &HDDDDDD: SetPixel hdc, 12, 8, &HDEDEDE: SetPixel hdc, 13, 8, &HE1E1E1: SetPixel hdc, 14, 8, &HE0E0E0: SetPixel hdc, 15, 8, &HE1E1E1: SetPixel hdc, 16, 8, &HE0E0E0: SetPixel hdc, 17, 8, &HE1E1E1
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HAEAEAE: SetPixel hdc, 2, 9, &H919191: SetPixel hdc, 3, 9, &HB9B9B9: SetPixel hdc, 4, 9, &HC4C4C4: SetPixel hdc, 5, 9, &HCECECE: SetPixel hdc, 6, 9, &HD1D1D1: SetPixel hdc, 7, 9, &HDADADA: SetPixel hdc, 8, 9, &HDCDCDC: SetPixel hdc, 9, 9, &HDBDBDB: SetPixel hdc, 10, 9, &HDFDFDF: SetPixel hdc, 11, 9, &HE1E3E1: SetPixel hdc, 12, 9, &HE2E3E2: SetPixel hdc, 13, 9, &HE5E2E3: SetPixel hdc, 14, 9, &HE5E2E3: SetPixel hdc, 15, 9, &HE5E2E3: SetPixel hdc, 16, 9, &HE5E2E3: SetPixel hdc, 17, 9, &HE5E2E3
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HA0A0A0: SetPixel hdc, 2, 10, &H999999: SetPixel hdc, 3, 10, &HC3C3C3: SetPixel hdc, 4, 10, &HC9C9C9: SetPixel hdc, 5, 10, &HD5D5D5: SetPixel hdc, 6, 10, &HD7D7D7: SetPixel hdc, 7, 10, &HDFDFDF: SetPixel hdc, 8, 10, &HE0E0E0: SetPixel hdc, 9, 10, &HE0E0E0: SetPixel hdc, 10, 10, &HE4E4E4: SetPixel hdc, 11, 10, &HE6E8E6: SetPixel hdc, 12, 10, &HE8E7E7: SetPixel hdc, 13, 10, &HEAE7E8: SetPixel hdc, 14, 10, &HEAE7E8: SetPixel hdc, 15, 10, &HEAE7E8: SetPixel hdc, 16, 10, &HEAE7E8: SetPixel hdc, 17, 10, &HEAE7E8
    SetPixel hdc, 0, 11, &HF5F5F5: SetPixel hdc, 1, 11, &HA3A3A3: SetPixel hdc, 2, 11, &H9B9B9B: SetPixel hdc, 3, 11, &HC6C6C6: SetPixel hdc, 4, 11, &HD3D3D3: SetPixel hdc, 5, 11, &HD6D6D6: SetPixel hdc, 6, 11, &HDDDDDD: SetPixel hdc, 7, 11, &HE1E1E1: SetPixel hdc, 8, 11, &HE3E3E3: SetPixel hdc, 9, 11, &HE6E6E6: SetPixel hdc, 10, 11, &HE7E8E7: SetPixel hdc, 11, 11, &HE9EAE9: SetPixel hdc, 12, 11, &HE8EAE9: SetPixel hdc, 13, 11, &HE8EBE9: SetPixel hdc, 14, 11, &HE8EBE9: SetPixel hdc, 15, 11, &HE8EBE9: SetPixel hdc, 16, 11, &HE8EBE9: SetPixel hdc, 17, 11, &HE8EBE9
    SetPixel hdc, 0, 12, &HF5F5F5: SetPixel hdc, 1, 12, &HAAAAAA: SetPixel hdc, 2, 12, &H8E8E8E: SetPixel hdc, 3, 12, &HD0D0D0: SetPixel hdc, 4, 12, &HDADADA: SetPixel hdc, 5, 12, &HDFDFDF: SetPixel hdc, 6, 12, &HE4E4E4: SetPixel hdc, 7, 12, &HE6E6E6: SetPixel hdc, 8, 12, &HE8E8E8: SetPixel hdc, 9, 12, &HECECEC: SetPixel hdc, 10, 12, &HEEEFEE: SetPixel hdc, 11, 12, &HEEF0EF: SetPixel hdc, 12, 12, &HEEF0EF: SetPixel hdc, 13, 12, &HEEF1EF: SetPixel hdc, 14, 12, &HEEF1EF: SetPixel hdc, 15, 12, &HEEF1EF: SetPixel hdc, 16, 12, &HEEF1EF: SetPixel hdc, 17, 12, &HEEF1EF
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HAAAAAA: SetPixel hdc, 2, tmph + 12, &H8E8E8E: SetPixel hdc, 3, tmph + 12, &HD0D0D0: SetPixel hdc, 4, tmph + 12, &HDADADA: SetPixel hdc, 5, tmph + 12, &HDFDFDF: SetPixel hdc, 6, tmph + 12, &HE4E4E4: SetPixel hdc, 7, tmph + 12, &HE6E6E6: SetPixel hdc, 8, tmph + 12, &HE8E8E8: SetPixel hdc, 9, tmph + 12, &HECECEC: SetPixel hdc, 10, tmph + 12, &HEEEFEE: SetPixel hdc, 11, tmph + 12, &HEEF0EF: SetPixel hdc, 12, tmph + 12, &HEEF0EF: SetPixel hdc, 13, tmph + 12, &HEEF1EF: SetPixel hdc, 14, tmph + 12, &HEEF1EF: SetPixel hdc, 15, tmph + 12, &HEEF1EF: SetPixel hdc, 16, tmph + 12, &HEEF1EF: SetPixel hdc, 17, tmph + 12, &HEEF1EF
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC2C2C2: SetPixel hdc, 2, tmph + 13, &H838383: SetPixel hdc, 3, tmph + 13, &HCFCFCF: SetPixel hdc, 4, tmph + 13, &HDEDEDE: SetPixel hdc, 5, tmph + 13, &HE3E3E3: SetPixel hdc, 6, tmph + 13, &HE8E8E8: SetPixel hdc, 7, tmph + 13, &HEAEAEA: SetPixel hdc, 8, tmph + 13, &HEDEDED: SetPixel hdc, 9, tmph + 13, &HF1F1F1: SetPixel hdc, 10, tmph + 13, &HF2F2F2: SetPixel hdc, 11, tmph + 13, &HF2F2F2: SetPixel hdc, 12, tmph + 13, &HF2F2F2: SetPixel hdc, 13, tmph + 13, &HF2F2F2: SetPixel hdc, 14, tmph + 13, &HF2F2F2: SetPixel hdc, 15, tmph + 13, &HF2F2F2: SetPixel hdc, 16, tmph + 13, &HF2F2F2: SetPixel hdc, 17, tmph + 13, &HF2F2F2
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H818181: SetPixel hdc, 3, tmph + 14, &HABABAB: SetPixel hdc, 4, tmph + 14, &HDCDCDC: SetPixel hdc, 5, tmph + 14, &HE5E5E5: SetPixel hdc, 6, tmph + 14, &HEDEDED: SetPixel hdc, 7, tmph + 14, &HEFEFEF: SetPixel hdc, 8, tmph + 14, &HF1F1F1: SetPixel hdc, 9, tmph + 14, &HF4F4F4: SetPixel hdc, 10, tmph + 14, &HF5F5F5: SetPixel hdc, 11, tmph + 14, &HF5F5F5: SetPixel hdc, 12, tmph + 14, &HF5F5F5: SetPixel hdc, 13, tmph + 14, &HF5F5F5: SetPixel hdc, 14, tmph + 14, &HF5F5F5: SetPixel hdc, 15, tmph + 14, &HF5F5F5: SetPixel hdc, 16, tmph + 14, &HF5F5F5: SetPixel hdc, 17, tmph + 14, &HF5F5F5
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &HA0A0A0: SetPixel hdc, 3, tmph + 15, &H898989: SetPixel hdc, 4, tmph + 15, &HDEDEDE: SetPixel hdc, 5, tmph + 15, &HE9E9E9: SetPixel hdc, 6, tmph + 15, &HEEEEEE: SetPixel hdc, 7, tmph + 15, &HF4F4F4: SetPixel hdc, 8, tmph + 15, &HF5F5F5: SetPixel hdc, 9, tmph + 15, &HFAFAFA: SetPixel hdc, 10, tmph + 15, &HFFFDFD: SetPixel hdc, 11, tmph + 15, &HFFFEFE: SetPixel hdc, 12, tmph + 15, &HFFFDFD: SetPixel hdc, 13, tmph + 15, &HFFFEFE: SetPixel hdc, 14, tmph + 15, &HFFFDFD: SetPixel hdc, 15, tmph + 15, &HFFFEFE: SetPixel hdc, 16, tmph + 15, &HFFFDFD: SetPixel hdc, 17, tmph + 15, &HFFFEFE
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD6D6D6: SetPixel hdc, 3, tmph + 16, &H7B7B7B: SetPixel hdc, 4, tmph + 16, &H8D8D8D: SetPixel hdc, 5, tmph + 16, &HE4E4E4: SetPixel hdc, 6, tmph + 16, &HF0F0F0: SetPixel hdc, 7, tmph + 16, &HF6F6F6: SetPixel hdc, 8, tmph + 16, &HFEFEFE: SetPixel hdc, 9, tmph + 16, &HFEFEFE: SetPixel hdc, 10, tmph + 16, &HFFFEFE: SetPixel hdc, 12, tmph + 16, &HFFFEFE: SetPixel hdc, 14, tmph + 16, &HFFFEFE: SetPixel hdc, 16, tmph + 16, &HFFFEFE
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBEBEBE: SetPixel hdc, 4, tmph + 17, &H727272: SetPixel hdc, 5, tmph + 17, &H898989: SetPixel hdc, 6, tmph + 17, &HEBEBEB: SetPixel hdc, 7, tmph + 17, &HF5F5F5: SetPixel hdc, 8, tmph + 17, &HFCFCFC: SetPixel hdc, 10, tmph + 17, &HFDFDFD: SetPixel hdc, 11, tmph + 17, &HFDFDFD: SetPixel hdc, 12, tmph + 17, &HFDFDFD: SetPixel hdc, 13, tmph + 17, &HFDFDFD: SetPixel hdc, 14, tmph + 17, &HFDFDFD: SetPixel hdc, 15, tmph + 17, &HFDFDFD: SetPixel hdc, 16, tmph + 17, &HFDFDFD: SetPixel hdc, 17, tmph + 17, &HFDFDFD
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HB9B9B9: SetPixel hdc, 5, tmph + 18, &H717171: SetPixel hdc, 6, tmph + 18, &H787878: SetPixel hdc, 7, tmph + 18, &HB6B6B6: SetPixel hdc, 8, tmph + 18, &HF7F7F7: SetPixel hdc, 9, tmph + 18, &HFCFCFC: SetPixel hdc, 10, tmph + 18, &HFEFEFE: SetPixel hdc, 11, tmph + 18, &HFEFEFE: SetPixel hdc, 12, tmph + 18, &HFEFEFE: SetPixel hdc, 13, tmph + 18, &HFEFEFE: SetPixel hdc, 14, tmph + 18, &HFEFEFE: SetPixel hdc, 15, tmph + 18, &HFEFEFE: SetPixel hdc, 16, tmph + 18, &HFEFEFE: SetPixel hdc, 17, tmph + 18, &HFEFEFE
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8E8E8E: SetPixel hdc, 7, tmph + 19, &H6C6C6C: SetPixel hdc, 8, tmph + 19, &H757575: SetPixel hdc, 9, tmph + 19, &H9F9F9F: SetPixel hdc, 10, tmph + 19, &HC7C7C7: SetPixel hdc, 11, tmph + 19, &HE9E9E9: SetPixel hdc, 12, tmph + 19, &HFBFBFB: SetPixel hdc, 13, tmph + 19, &HFBFBFB: SetPixel hdc, 14, tmph + 19, &HFBFBFB: SetPixel hdc, 15, tmph + 19, &HFBFBFB: SetPixel hdc, 16, tmph + 19, &HFBFBFB: SetPixel hdc, 17, tmph + 19, &HFBFBFB
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC1C1C1: SetPixel hdc, 8, tmph + 20, &H9D9D9D: SetPixel hdc, 9, tmph + 20, &H7B7B7B: SetPixel hdc, 10, tmph + 20, &H5E5E5E: SetPixel hdc, 11, tmph + 20, &H535353: SetPixel hdc, 12, tmph + 20, &H4D4D4D: SetPixel hdc, 13, tmph + 20, &H4B4B4B: SetPixel hdc, 14, tmph + 20, &H505050: SetPixel hdc, 15, tmph + 20, &H525252: SetPixel hdc, 16, tmph + 20, &H555555: SetPixel hdc, 17, tmph + 20, &H545454
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H67696A: SetPixel hdc, tmpw + 18, 0, &H666869: SetPixel hdc, tmpw + 19, 0, &H716F6F: SetPixel hdc, tmpw + 20, 0, &H6F6D6D: SetPixel hdc, tmpw + 21, 0, &H6F706E: SetPixel hdc, tmpw + 22, 0, &H727371: SetPixel hdc, tmpw + 23, 0, &H6E6E6E: SetPixel hdc, tmpw + 24, 0, &H707070: SetPixel hdc, tmpw + 25, 0, &HA6A6A6: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 1, &HF5F4F6: SetPixel hdc, tmpw + 18, 1, &HF5F4F6: SetPixel hdc, tmpw + 19, 1, &HF5F4F6: SetPixel hdc, tmpw + 20, 1, &HF5F4F6: SetPixel hdc, tmpw + 21, 1, &HF4F3F5: SetPixel hdc, tmpw + 22, 1, &HF1F0F2: SetPixel hdc, tmpw + 23, 1, &HE0E0E0: SetPixel hdc, tmpw + 24, 1, &HC3C3C3: SetPixel hdc, tmpw + 25, 1, &H848484: SetPixel hdc, tmpw + 26, 1, &H6B6B6B: SetPixel hdc, tmpw + 27, 1, &HA0A0A0: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 2, &HF3F2F4: SetPixel hdc, tmpw + 18, 2, &HF2F1F3: SetPixel hdc, tmpw + 19, 2, &HF3F2F4: SetPixel hdc, tmpw + 20, 2, &HF3F2F4: SetPixel hdc, tmpw + 21, 2, &HF0EFF1: SetPixel hdc, tmpw + 22, 2, &HF2F1F3: SetPixel hdc, tmpw + 23, 2, &HF6F6F6: SetPixel hdc, tmpw + 24, 2, &HE8E8E8: SetPixel hdc, tmpw + 25, 2, &HE0E0E0: SetPixel hdc, tmpw + 26, 2, &H999999: SetPixel hdc, tmpw + 27, 2, &H696969: SetPixel hdc, tmpw + 28, 2, &H717171: SetPixel hdc, tmpw + 29, 2, &HEBEBEB: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 3, &HEEEEEE: SetPixel hdc, tmpw + 18, 3, &HEDEDED: SetPixel hdc, tmpw + 19, 3, &HEEEEEE: SetPixel hdc, tmpw + 20, 3, &HEEEEEE: SetPixel hdc, tmpw + 21, 3, &HEEEEEE: SetPixel hdc, tmpw + 22, 3, &HEEEEEE: SetPixel hdc, tmpw + 23, 3, &HE9E9E9: SetPixel hdc, tmpw + 24, 3, &HEAEAEA: SetPixel hdc, tmpw + 25, 3, &HE7E7E7: SetPixel hdc, tmpw + 26, 3, &HD0D0D0: SetPixel hdc, tmpw + 27, 3, &H939393: SetPixel hdc, tmpw + 28, 3, &H727272: SetPixel hdc, tmpw + 29, 3, &H6F6F6F: SetPixel hdc, tmpw + 30, 3, &HEFEFEF: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 4, &HEBEBEB: SetPixel hdc, tmpw + 18, 4, &HEBEBEB: SetPixel hdc, tmpw + 19, 4, &HEBEBEB: SetPixel hdc, tmpw + 20, 4, &HEBEBEB: SetPixel hdc, tmpw + 21, 4, &HEDEDED: SetPixel hdc, tmpw + 22, 4, &HE6E6E6: SetPixel hdc, tmpw + 23, 4, &HE9E9E9: SetPixel hdc, tmpw + 24, 4, &HE6E6E6: SetPixel hdc, tmpw + 25, 4, &HDEDEDE: SetPixel hdc, tmpw + 26, 4, &HDCDCDC: SetPixel hdc, tmpw + 27, 4, &HB2B2B2: SetPixel hdc, tmpw + 28, 4, &H919191: SetPixel hdc, tmpw + 29, 4, &H6E6E6E: SetPixel hdc, tmpw + 30, 4, &H7F7F7F: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 5, &HEBE8EA: SetPixel hdc, tmpw + 18, 5, &HEAE7E9: SetPixel hdc, tmpw + 19, 5, &HEBE8EA: SetPixel hdc, tmpw + 20, 5, &HEBE8EA: SetPixel hdc, tmpw + 21, 5, &HE5E8E6: SetPixel hdc, tmpw + 22, 5, &HE7EAE8: SetPixel hdc, tmpw + 23, 5, &HE5E5E5: SetPixel hdc, tmpw + 24, 5, &HE3E3E3: SetPixel hdc, tmpw + 25, 5, &HDFDFDF: SetPixel hdc, tmpw + 26, 5, &HDCDCDC: SetPixel hdc, tmpw + 27, 5, &HC3C3C3: SetPixel hdc, tmpw + 28, 5, &HA7A7A7: SetPixel hdc, tmpw + 29, 5, &H969696: SetPixel hdc, tmpw + 30, 5, &H717171: SetPixel hdc, tmpw + 31, 5, &HC5C5C5: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 6, &HEBE8EA: SetPixel hdc, tmpw + 18, 6, &HEBE8EA: SetPixel hdc, tmpw + 19, 6, &HEBE8EA: SetPixel hdc, tmpw + 20, 6, &HEBE8EA: SetPixel hdc, tmpw + 21, 6, &HE8EBE9: SetPixel hdc, tmpw + 22, 6, &HE3E6E4: SetPixel hdc, tmpw + 23, 6, &HE5E5E5: SetPixel hdc, tmpw + 24, 6, &HE2E2E2: SetPixel hdc, tmpw + 25, 6, &HE0E0E0: SetPixel hdc, tmpw + 26, 6, &HDADADA: SetPixel hdc, tmpw + 27, 6, &HC7C7C7: SetPixel hdc, tmpw + 28, 6, &HB5B5B5: SetPixel hdc, tmpw + 29, 6, &HA6A6A6: SetPixel hdc, tmpw + 30, 6, &H8C8C8C: SetPixel hdc, tmpw + 31, 6, &H808080: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 7, &HEAEAEA: SetPixel hdc, tmpw + 18, 7, &HEAEAEA: SetPixel hdc, tmpw + 19, 7, &HEAEAEA: SetPixel hdc, tmpw + 20, 7, &HEAEAEA: SetPixel hdc, tmpw + 21, 7, &HE9E6E8: SetPixel hdc, tmpw + 22, 7, &HE9E6E8: SetPixel hdc, tmpw + 23, 7, &HE4E4E4: SetPixel hdc, tmpw + 24, 7, &HE2E2E2: SetPixel hdc, tmpw + 25, 7, &HDFDFDF: SetPixel hdc, tmpw + 26, 7, &HD7D7D7: SetPixel hdc, tmpw + 27, 7, &HC4C4C4: SetPixel hdc, tmpw + 28, 7, &HB7B7B7: SetPixel hdc, tmpw + 29, 7, &HB4B5B3: SetPixel hdc, tmpw + 30, 7, &H9D9E9C: SetPixel hdc, tmpw + 31, 7, &H777777: SetPixel hdc, tmpw + 32, 7, &HE7E7E7: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 8, &HE1E1E1: SetPixel hdc, tmpw + 18, 8, &HE0E0E0: SetPixel hdc, tmpw + 19, 8, &HE1E1E1: SetPixel hdc, tmpw + 20, 8, &HE1E1E1: SetPixel hdc, tmpw + 21, 8, &HDFDCDE: SetPixel hdc, tmpw + 22, 8, &HDDDADC: SetPixel hdc, tmpw + 23, 8, &HDBDBDB: SetPixel hdc, tmpw + 24, 8, &HD6D6D6: SetPixel hdc, tmpw + 25, 8, &HD5D5D5: SetPixel hdc, tmpw + 26, 8, &HD1D1D1: SetPixel hdc, tmpw + 27, 8, &HC9C9C9: SetPixel hdc, tmpw + 28, 8, &HC4C4C4: SetPixel hdc, tmpw + 29, 8, &HC0C1BF: SetPixel hdc, tmpw + 30, 8, &HAFB0AE: SetPixel hdc, tmpw + 31, 8, &H818181: SetPixel hdc, tmpw + 32, 8, &HC3C3C3: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 9, &HE5E2E3: SetPixel hdc, tmpw + 18, 9, &HE5E2E3: SetPixel hdc, tmpw + 19, 9, &HE5E2E3: SetPixel hdc, tmpw + 20, 9, &HE5E2E3: SetPixel hdc, tmpw + 21, 9, &HE1E1E1: SetPixel hdc, tmpw + 22, 9, &HE1E1E1: SetPixel hdc, tmpw + 23, 9, &HE1E1E1: SetPixel hdc, tmpw + 24, 9, &HDDDDDD: SetPixel hdc, tmpw + 25, 9, &HDBDBDB: SetPixel hdc, tmpw + 26, 9, &HD8D8D8: SetPixel hdc, tmpw + 27, 9, &HD2D2D2: SetPixel hdc, tmpw + 28, 9, &HCBCBCB: SetPixel hdc, tmpw + 29, 9, &HC4C4C4: SetPixel hdc, tmpw + 30, 9, &HBABABA: SetPixel hdc, tmpw + 31, 9, &H989898: SetPixel hdc, tmpw + 32, 9, &HA6A6A6: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 10, &HEAE7E8: SetPixel hdc, tmpw + 18, 10, &HEAE7E8: SetPixel hdc, tmpw + 19, 10, &HEAE7E8: SetPixel hdc, tmpw + 20, 10, &HEAE7E8: SetPixel hdc, tmpw + 21, 10, &HE7E7E7: SetPixel hdc, tmpw + 22, 10, &HE6E6E6: SetPixel hdc, tmpw + 23, 10, &HE4E4E4: SetPixel hdc, tmpw + 24, 10, &HE0E0E0: SetPixel hdc, tmpw + 25, 10, &HE0E0E0: SetPixel hdc, tmpw + 26, 10, &HDEDEDE: SetPixel hdc, tmpw + 27, 10, &HD9D9D9: SetPixel hdc, tmpw + 28, 10, &HD3D3D3: SetPixel hdc, tmpw + 29, 10, &HCCCCCC: SetPixel hdc, tmpw + 30, 10, &HC3C3C3: SetPixel hdc, tmpw + 31, 10, &HA3A3A3: SetPixel hdc, tmpw + 32, 10, &H9C9C9C: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 11, &HE8EBE9: SetPixel hdc, tmpw + 18, 11, &HE8EBE9: SetPixel hdc, tmpw + 19, 11, &HE8EBE9: SetPixel hdc, tmpw + 20, 11, &HE8EBE9: SetPixel hdc, tmpw + 21, 11, &HE9EAE8: SetPixel hdc, tmpw + 22, 11, &HE8E9E7: SetPixel hdc, tmpw + 23, 11, &HE9E9E9: SetPixel hdc, tmpw + 24, 11, &HE5E5E5: SetPixel hdc, tmpw + 25, 11, &HE4E4E4: SetPixel hdc, tmpw + 26, 11, &HE2E2E2: SetPixel hdc, tmpw + 27, 11, &HDBDBDB: SetPixel hdc, tmpw + 28, 11, &HD9D9D9: SetPixel hdc, tmpw + 29, 11, &HD1D1D1: SetPixel hdc, tmpw + 30, 11, &HC8C8C8: SetPixel hdc, tmpw + 31, 11, &HA4A4A4: SetPixel hdc, tmpw + 32, 11, &HA2A2A2: SetPixel hdc, tmpw + 33, 11, &HF4F4F4: SetPixel hdc, tmpw + 34, 11, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 12, &HEEF1EF: SetPixel hdc, tmpw + 18, 12, &HEEF1EF: SetPixel hdc, tmpw + 19, 12, &HEEF1EF: SetPixel hdc, tmpw + 20, 12, &HEEF1EF: SetPixel hdc, tmpw + 21, 12, &HEEEFED: SetPixel hdc, tmpw + 22, 12, &HEFF0EE: SetPixel hdc, tmpw + 23, 12, &HEEEEEE: SetPixel hdc, tmpw + 24, 12, &HECECEC: SetPixel hdc, tmpw + 25, 12, &HEAEAEA: SetPixel hdc, tmpw + 26, 12, &HE7E7E7: SetPixel hdc, tmpw + 27, 12, &HE2E2E2: SetPixel hdc, tmpw + 28, 12, &HDFDFDF: SetPixel hdc, tmpw + 29, 12, &HD8D8D8: SetPixel hdc, tmpw + 30, 12, &HD4D4D4: SetPixel hdc, tmpw + 31, 12, &H999999: SetPixel hdc, tmpw + 32, 12, &HAFAFAF: SetPixel hdc, tmpw + 33, 12, &HF5F5F5: SetPixel hdc, tmpw + 34, 12, &HFFFFFFFF
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 18, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 19, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 20, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 21, tmph + 12, &HEEEFED: SetPixel hdc, tmpw + 22, tmph + 12, &HEFF0EE: SetPixel hdc, tmpw + 23, tmph + 12, &HEEEEEE: SetPixel hdc, tmpw + 24, tmph + 12, &HECECEC: SetPixel hdc, tmpw + 25, tmph + 12, &HEAEAEA: SetPixel hdc, tmpw + 26, tmph + 12, &HE7E7E7: SetPixel hdc, tmpw + 27, tmph + 12, &HE2E2E2: SetPixel hdc, tmpw + 28, tmph + 12, &HDFDFDF: SetPixel hdc, tmpw + 29, tmph + 12, &HD8D8D8: SetPixel hdc, tmpw + 30, tmph + 12, &HD4D4D4: SetPixel hdc, tmpw + 31, tmph + 12, &H999999: SetPixel hdc, tmpw + 32, tmph + 12, &HAFAFAF: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5
    SetPixel hdc, tmpw + 17, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 18, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 19, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 20, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 21, tmph + 13, &HF5F4F6: SetPixel hdc, tmpw + 22, tmph + 13, &HF0EFF1: SetPixel hdc, tmpw + 23, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 24, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 25, tmph + 13, &HECECEC: SetPixel hdc, tmpw + 26, tmph + 13, &HEAEAEA: SetPixel hdc, tmpw + 27, tmph + 13, &HEBEBEB: SetPixel hdc, tmpw + 28, tmph + 13, &HE3E3E3: SetPixel hdc, tmpw + 29, tmph + 13, &HDEDEDE: SetPixel hdc, tmpw + 30, tmph + 13, &HD1D1D1: SetPixel hdc, tmpw + 31, tmph + 13, &H8A8A8A: SetPixel hdc, tmpw + 32, tmph + 13, &HD5D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8
    SetPixel hdc, tmpw + 17, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 18, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 19, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 20, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 21, tmph + 14, &HF8F7F9: SetPixel hdc, tmpw + 22, tmph + 14, &HF7F6F8: SetPixel hdc, tmpw + 23, tmph + 14, &HF7F7F7: SetPixel hdc, tmpw + 24, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 25, tmph + 14, &HEFEFEF: SetPixel hdc, tmpw + 26, tmph + 14, &HEEEEEE: SetPixel hdc, tmpw + 27, tmph + 14, &HECECEC: SetPixel hdc, tmpw + 28, tmph + 14, &HE5E5E5: SetPixel hdc, tmpw + 29, tmph + 14, &HDEDEDE: SetPixel hdc, tmpw + 30, tmph + 14, &HB3B3B3: SetPixel hdc, tmpw + 31, tmph + 14, &H808080: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD
    SetPixel hdc, tmpw + 17, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 18, tmph + 15, &HFFFDFD: SetPixel hdc, tmpw + 19, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 20, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 21, tmph + 15, &HFBFBFB: SetPixel hdc, tmpw + 22, tmph + 15, &HFCFCFC: SetPixel hdc, tmpw + 23, tmph + 15, &HFEFEFE: SetPixel hdc, tmpw + 24, tmph + 15, &HF8F8F8: SetPixel hdc, tmpw + 25, tmph + 15, &HF7F7F7: SetPixel hdc, tmpw + 26, tmph + 15, &HF5F5F5: SetPixel hdc, tmpw + 27, tmph + 15, &HEDEDED: SetPixel hdc, tmpw + 28, tmph + 15, &HEAEAEA: SetPixel hdc, tmpw + 29, tmph + 15, &HE0E0E0: SetPixel hdc, tmpw + 30, tmph + 15, &H8D8D8D: SetPixel hdc, tmpw + 31, tmph + 15, &HBABABA: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1
    SetPixel hdc, tmpw + 18, tmph + 16, &HFFFEFE: SetPixel hdc, tmpw + 22, tmph + 16, &HFEFEFE: SetPixel hdc, tmpw + 23, tmph + 16, &HFEFEFE: SetPixel hdc, tmpw + 25, tmph + 16, &HFCFCFC: SetPixel hdc, tmpw + 26, tmph + 16, &HF6F6F6: SetPixel hdc, tmpw + 27, tmph + 16, &HF2F2F2: SetPixel hdc, tmpw + 28, tmph + 16, &HE7E7E7: SetPixel hdc, tmpw + 29, tmph + 16, &H989898: SetPixel hdc, tmpw + 30, tmph + 16, &H828282: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9
    SetPixel hdc, tmpw + 17, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 18, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 19, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 20, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 21, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 23, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 25, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 26, tmph + 17, &HF6F6F6: SetPixel hdc, tmpw + 27, tmph + 17, &HF1F1F1: SetPixel hdc, tmpw + 28, tmph + 17, &H979797: SetPixel hdc, tmpw + 29, tmph + 17, &H6F6F6F: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE
    SetPixel hdc, tmpw + 17, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 18, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 19, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 20, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 22, tmph + 18, &HFDFDFD: SetPixel hdc, tmpw + 23, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 24, tmph + 18, &HFDFDFD: SetPixel hdc, tmpw + 25, tmph + 18, &HFCFCFC: SetPixel hdc, tmpw + 26, tmph + 18, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 18, &H838383: SetPixel hdc, tmpw + 28, tmph + 18, &H6F6F6F: SetPixel hdc, tmpw + 29, tmph + 18, &HC8C8C8: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC
    SetPixel hdc, tmpw + 17, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 18, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 19, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 20, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 21, tmph + 19, &HFAFAFA: SetPixel hdc, tmpw + 22, tmph + 19, &HEFEFEF: SetPixel hdc, tmpw + 23, tmph + 19, &HD0D0D0: SetPixel hdc, tmpw + 24, tmph + 19, &HA3A3A3: SetPixel hdc, tmpw + 25, tmph + 19, &H7E7E7E: SetPixel hdc, tmpw + 26, tmph + 19, &H6A6A6A: SetPixel hdc, tmpw + 27, tmph + 19, &H8F8F8F: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA
    SetPixel hdc, tmpw + 17, tmph + 20, &H545454: SetPixel hdc, tmpw + 18, tmph + 20, &H555555: SetPixel hdc, tmpw + 19, tmph + 20, &H525252: SetPixel hdc, tmpw + 20, tmph + 20, &H505050: SetPixel hdc, tmpw + 21, tmph + 20, &H535353: SetPixel hdc, tmpw + 22, tmph + 20, &H525252: SetPixel hdc, tmpw + 23, tmph + 20, &H616161: SetPixel hdc, tmpw + 24, tmph + 20, &H7A7A7A: SetPixel hdc, tmpw + 25, tmph + 20, &HA3A3A3: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD
    'Vlines
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HA0A0A0: DrawLineApi 2, tmph, 2, tmph1, &H999999: DrawLineApi 3, tmph, 3, tmph1, &HC3C3C3
    DrawLineApi 4, tmph, 4, tmph1, &HC9C9C9: DrawLineApi 5, tmph, 5, tmph1, &HD5D5D5: DrawLineApi 6, tmph, 6, tmph1, &HD7D7D7: DrawLineApi 7, tmph, 7, tmph1, &HDFDFDF
    DrawLineApi 8, tmph, 8, tmph1, &HE0E0E0: DrawLineApi 9, tmph, 9, tmph1, &HE0E0E0: DrawLineApi 10, tmph, 10, tmph1, &HE4E4E4: DrawLineApi 11, tmph, 11, tmph1, &HE6E8E6
    DrawLineApi 12, tmph, 12, tmph1, &HE8E7E7: DrawLineApi 13, tmph, 13, tmph1, &HEAE7E8: DrawLineApi 14, tmph, 14, tmph1, &HEAE7E8: DrawLineApi 15, tmph, 15, tmph1, &HEAE7E8
    DrawLineApi 16, tmph, 16, tmph1, &HEAE7E8: DrawLineApi 17, tmph, 17, tmph1, &HEAE7E8: DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HEAE7E8: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HEAE7E8
    DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HEAE7E8: DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HEAE7E8: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HE7E7E7
    DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HE6E6E6: DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HE4E4E4: DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HE0E0E0
    DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HE0E0E0: DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HDEDEDE: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HD9D9D9
    DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HD3D3D3: DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HCCCCCC: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HC3C3C3
    DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &HA3A3A3: DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &H9C9C9C: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
    'HLines
    DrawLineApi 17, 0, lw - 17, 0, &H67696A
    DrawLineApi 17, 1, lw - 17, 1, &HF5F4F6
    DrawLineApi 17, 2, lw - 17, 2, &HF3F2F4
    DrawLineApi 17, 3, lw - 17, 3, &HEEEEEE
    DrawLineApi 17, 4, lw - 17, 4, &HEBEBEB
    DrawLineApi 17, 5, lw - 17, 5, &HEBE8EA
    DrawLineApi 17, 6, lw - 17, 6, &HEBE8EA
    DrawLineApi 17, 7, lw - 17, 7, &HEAEAEA
    DrawLineApi 17, 8, lw - 17, 8, &HE1E1E1
    DrawLineApi 17, 9, lw - 17, 9, &HE5E2E3
    DrawLineApi 17, 10, lw - 17, 10, &HEAE7E8
    DrawLineApi 17, 11, lw - 17, 11, &HE8EBE9
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HE8EBE9
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HEEF1EF
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HF2F2F2
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HF5F5F5
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HFFFEFE
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HFFFFFF
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HFDFDFD
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HFEFEFE
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HFBFBFB
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H545454
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC
    
End Sub

Private Sub DrawAquaHot()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    
    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HE2A66A, lpRect
    
    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE6E5E5: SetPixel hdc, 8, 0, &HA9A5A5: SetPixel hdc, 9, 0, &H6C5E5E: SetPixel hdc, 10, 0, &H482729: SetPixel hdc, 11, 0, &H370D0C: SetPixel hdc, 12, 0, &H370706: SetPixel hdc, 13, 0, &H360605: SetPixel hdc, 14, 0, &H3A0606: SetPixel hdc, 15, 0, &H410807: SetPixel hdc, 16, 0, &H450707: SetPixel hdc, 17, 0, &H450608:
    SetPixel hdc, 5, 1, &HF0EFEF: SetPixel hdc, 6, 1, &HA38A8C: SetPixel hdc, 7, 1, &H6E342F: SetPixel hdc, 8, 1, &H661F1A: SetPixel hdc, 9, 1, &H9B6A63: SetPixel hdc, 10, 1, &HC9A29D: SetPixel hdc, 11, 1, &HE2BFBD: SetPixel hdc, 12, 1, &HE8C9C6: SetPixel hdc, 13, 1, &HEFD3CC: SetPixel hdc, 14, 1, &HEFD3CC: SetPixel hdc, 15, 1, &HF0D5C9: SetPixel hdc, 16, 1, &HF0D5C9: SetPixel hdc, 17, 1, &HF1D4C9:
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE5E5E5: SetPixel hdc, 5, 2, &H755E5E: SetPixel hdc, 6, 2, &H41070C: SetPixel hdc, 7, 2, &H7F2D28: SetPixel hdc, 8, 2, &HEC9892: SetPixel hdc, 9, 2, &HECB6AF: SetPixel hdc, 10, 2, &HE3BBB6: SetPixel hdc, 11, 2, &HE3C0BD: SetPixel hdc, 12, 2, &HE1C2BF: SetPixel hdc, 13, 2, &HDFC3BC: SetPixel hdc, 14, 2, &HDFC3BC: SetPixel hdc, 15, 2, &HE4C9BD: SetPixel hdc, 16, 2, &HE4C9BD: SetPixel hdc, 17, 2, &HE5C8BD:
    SetPixel hdc, 3, 3, &HEEEEEE: SetPixel hdc, 4, 3, &H8A5A5A: SetPixel hdc, 5, 3, &H7A0702: SetPixel hdc, 6, 3, &H901501: SetPixel hdc, 7, 3, &HC38365: SetPixel hdc, 8, 3, &HE3B08F: SetPixel hdc, 9, 3, &HE1B394: SetPixel hdc, 10, 3, &HE5B798: SetPixel hdc, 11, 3, &HE6BC99: SetPixel hdc, 12, 3, &HE7BD9A: SetPixel hdc, 13, 3, &HE4BC99: SetPixel hdc, 14, 3, &HE7BF9C: SetPixel hdc, 15, 3, &HE9C1A1: SetPixel hdc, 16, 3, &HE8C0A1: SetPixel hdc, 17, 3, &HE8C0A1:
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H897879: SetPixel hdc, 4, 4, &H4D0909: SetPixel hdc, 5, 4, &H951905: SetPixel hdc, 6, 4, &HBF422E: SetPixel hdc, 7, 4, &HD49475: SetPixel hdc, 8, 4, &HD7A483: SetPixel hdc, 9, 4, &HDAAC8D: SetPixel hdc, 10, 4, &HDBAD8E: SetPixel hdc, 11, 4, &HD9AF8C: SetPixel hdc, 12, 4, &HDCB28F: SetPixel hdc, 13, 4, &HDDB592: SetPixel hdc, 14, 4, &HDCB491: SetPixel hdc, 15, 4, &HDFB797: SetPixel hdc, 16, 4, &HE0B898: SetPixel hdc, 17, 4, &HE0B898:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCDC9C9: SetPixel hdc, 3, 5, &H882517: SetPixel hdc, 4, 5, &H922100: SetPixel hdc, 5, 5, &HA13A00: SetPixel hdc, 6, 5, &HD57333: SetPixel hdc, 7, 5, &HDFA36F: SetPixel hdc, 8, 5, &HDDA876: SetPixel hdc, 9, 5, &HD8A573: SetPixel hdc, 10, 5, &HDFAE80: SetPixel hdc, 11, 5, &HDBAD7D: SetPixel hdc, 12, 5, &HDFB084: SetPixel hdc, 13, 5, &HDFB286: SetPixel hdc, 14, 5, &HDFB188: SetPixel hdc, 15, 5, &HE1B58D: SetPixel hdc, 16, 5, &HE3B58E: SetPixel hdc, 17, 5, &HE3B48E:
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H7B706E: SetPixel hdc, 3, 6, &H871405: SetPixel hdc, 4, 6, &HA5330E: SetPixel hdc, 5, 6, &HB34C0D: SetPixel hdc, 6, 6, &HD27030: SetPixel hdc, 7, 6, &HD89C68: SetPixel hdc, 8, 6, &HDAA573: SetPixel hdc, 9, 6, &HD9A674: SetPixel hdc, 10, 6, &HD9A87A: SetPixel hdc, 11, 6, &HDBAD7D: SetPixel hdc, 12, 6, &HDBAC80: SetPixel hdc, 13, 6, &HDCAF83: SetPixel hdc, 14, 6, &HDFB188: SetPixel hdc, 15, 6, &HDEB28A: SetPixel hdc, 16, 6, &HDFB18A: SetPixel hdc, 17, 6, &HE0B18B:
    SetPixel hdc, 1, 7, &HE8E8E7: SetPixel hdc, 2, 7, &H773F34: SetPixel hdc, 3, 7, &H9F2C00: SetPixel hdc, 4, 7, &HBA4B07: SetPixel hdc, 5, 7, &HC35E10: SetPixel hdc, 6, 7, &HCC7323: SetPixel hdc, 7, 7, &HDB8F46: SetPixel hdc, 8, 7, &HE8A763: SetPixel hdc, 9, 7, &HE3A76C: SetPixel hdc, 10, 7, &HE7AB70: SetPixel hdc, 11, 7, &HE8AE73: SetPixel hdc, 12, 7, &HE8AE73: SetPixel hdc, 13, 7, &HEDB17B: SetPixel hdc, 14, 7, &HEFB37D: SetPixel hdc, 15, 7, &HE9B57E: SetPixel hdc, 16, 7, &HE9B57E: SetPixel hdc, 17, 7, &HE9B47F:
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HCAC5C5: SetPixel hdc, 2, 8, &H682A1F: SetPixel hdc, 3, 8, &HB23E0C: SetPixel hdc, 4, 8, &HCC5D19: SetPixel hdc, 5, 8, &HCE691B: SetPixel hdc, 6, 8, &HCE7525: SetPixel hdc, 7, 8, &HCD8138: SetPixel hdc, 8, 8, &HC58440: SetPixel hdc, 9, 8, &HC5894E: SetPixel hdc, 10, 8, &HC98D52: SetPixel hdc, 11, 8, &HC88E53: SetPixel hdc, 12, 8, &HCC9257: SetPixel hdc, 13, 8, &HCF935D: SetPixel hdc, 14, 8, &HD0945E: SetPixel hdc, 15, 8, &HCE9963: SetPixel hdc, 16, 8, &HCE9963: SetPixel hdc, 17, 8, &HCE9963:
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HB9ADAB: SetPixel hdc, 2, 9, &H6E2B10: SetPixel hdc, 3, 9, &HB6580D: SetPixel hdc, 4, 9, &HCA6C20: SetPixel hdc, 5, 9, &HCE792B: SetPixel hdc, 6, 9, &HCE8132: SetPixel hdc, 7, 9, &HD08B42: SetPixel hdc, 8, 9, &HD3904B: SetPixel hdc, 9, 9, &HD3934C: SetPixel hdc, 10, 9, &HD89753: SetPixel hdc, 11, 9, &HDB9B5A: SetPixel hdc, 12, 9, &HDC9B5E: SetPixel hdc, 13, 9, &HDB9C60: SetPixel hdc, 14, 9, &HDB9C60: SetPixel hdc, 15, 9, &HDDA164: SetPixel hdc, 16, 9, &HDDA164: SetPixel hdc, 17, 9, &HDDA064:
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HB0A09E: SetPixel hdc, 2, 10, &H712E13: SetPixel hdc, 3, 10, &HBD5F14: SetPixel hdc, 4, 10, &HD17327: SetPixel hdc, 5, 10, &HD47F31: SetPixel hdc, 6, 10, &HD98C3D: SetPixel hdc, 7, 10, &HD9944B: SetPixel hdc, 8, 10, &HD7944F: SetPixel hdc, 9, 10, &HDC9C55: SetPixel hdc, 10, 10, &HDC9B57: SetPixel hdc, 11, 10, &HE3A362: SetPixel hdc, 12, 10, &HE3A265: SetPixel hdc, 13, 10, &HE2A367: SetPixel hdc, 14, 10, &HE0A165: SetPixel hdc, 15, 10, &HE3A66A: SetPixel hdc, 16, 10, &HE3A66A: SetPixel hdc, 17, 10, &HE2A66A
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 10, &HF7F7F7: SetPixel hdc, 1, tmph + 10, &HB0A09E: SetPixel hdc, 2, tmph + 10, &H712E13: SetPixel hdc, 3, tmph + 10, &HBD5F14: SetPixel hdc, 4, tmph + 10, &HD17327: SetPixel hdc, 5, tmph + 10, &HD47F31: SetPixel hdc, 6, tmph + 10, &HD98C3D: SetPixel hdc, 7, tmph + 10, &HD9944B: SetPixel hdc, 8, tmph + 10, &HD7944F: SetPixel hdc, 9, tmph + 10, &HDC9C55: SetPixel hdc, 10, tmph + 10, &HDC9B57: SetPixel hdc, 11, tmph + 10, &HE3A362: SetPixel hdc, 12, tmph + 10, &HE3A265: SetPixel hdc, 13, tmph + 10, &HE2A367: SetPixel hdc, 14, tmph + 10, &HE0A165: SetPixel hdc, 15, tmph + 10, &HE3A66A: SetPixel hdc, 16, tmph + 10, &HE3A66A: SetPixel hdc, 17, tmph + 10, &HE2A66A:
    SetPixel hdc, 0, tmph + 11, &HF5F5F5: SetPixel hdc, 1, tmph + 11, &HACA39E: SetPixel hdc, 2, tmph + 11, &H744421: SetPixel hdc, 3, tmph + 11, &HC56F1F: SetPixel hdc, 4, tmph + 11, &HD17A2A: SetPixel hdc, 5, tmph + 11, &HD58C42: SetPixel hdc, 6, tmph + 11, &HD7914B: SetPixel hdc, 7, tmph + 11, &HDF9854: SetPixel hdc, 8, tmph + 11, &HE4A05F: SetPixel hdc, 9, tmph + 11, &HE29F66: SetPixel hdc, 10, tmph + 11, &HE4A56B: SetPixel hdc, 11, tmph + 11, &HDDA467: SetPixel hdc, 12, tmph + 11, &HE0A76A: SetPixel hdc, 13, tmph + 11, &HE2A96C: SetPixel hdc, 14, tmph + 11, &HE3A870: SetPixel hdc, 15, tmph + 11, &HE6AC76: SetPixel hdc, 16, tmph + 11, &HE6AC76: SetPixel hdc, 17, tmph + 11, &HE6AC76:
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HB1AAA7: SetPixel hdc, 2, tmph + 12, &H825533: SetPixel hdc, 3, tmph + 12, &HCF792A: SetPixel hdc, 4, tmph + 12, &HE48D3D: SetPixel hdc, 5, tmph + 12, &HDD944A: SetPixel hdc, 6, tmph + 12, &HE49E58: SetPixel hdc, 7, tmph + 12, &HEBA460: SetPixel hdc, 8, tmph + 12, &HEEAA69: SetPixel hdc, 9, tmph + 12, &HF3B077: SetPixel hdc, 10, tmph + 12, &HEEAF75: SetPixel hdc, 11, tmph + 12, &HEBB275: SetPixel hdc, 12, tmph + 12, &HEFB679: SetPixel hdc, 13, tmph + 12, &HF1B87B: SetPixel hdc, 14, tmph + 12, &HF1B67E: SetPixel hdc, 15, tmph + 12, &HF2B781: SetPixel hdc, 16, tmph + 12, &HF1B681: SetPixel hdc, 17, tmph + 12, &HF1B681:
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC2C2C1: SetPixel hdc, 2, tmph + 13, &H6B5D4E: SetPixel hdc, 3, tmph + 13, &HC27831: SetPixel hdc, 4, tmph + 13, &HDA8E46: SetPixel hdc, 5, tmph + 13, &HE7A05C: SetPixel hdc, 6, tmph + 13, &HEAA665: SetPixel hdc, 7, tmph + 13, &HE9AF6E: SetPixel hdc, 8, tmph + 13, &HEFB377: SetPixel hdc, 9, tmph + 13, &HF3B579: SetPixel hdc, 10, tmph + 13, &HF7B97D: SetPixel hdc, 11, tmph + 13, &HF2BB7E: SetPixel hdc, 12, tmph + 13, &HF4BB83: SetPixel hdc, 13, tmph + 13, &HF5BE85: SetPixel hdc, 14, tmph + 13, &HF4BB87: SetPixel hdc, 15, tmph + 13, &HF5BE8A: SetPixel hdc, 16, tmph + 13, &HF5BD8A: SetPixel hdc, 17, tmph + 13, &HF3BD8A:
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H85796E: SetPixel hdc, 3, tmph + 14, &HB76F2B: SetPixel hdc, 4, tmph + 14, &HDE924A: SetPixel hdc, 5, tmph + 14, &HE8A15D: SetPixel hdc, 6, tmph + 14, &HF2AE6D: SetPixel hdc, 7, tmph + 14, &HF1B776: SetPixel hdc, 8, tmph + 14, &HF2B67A: SetPixel hdc, 9, tmph + 14, &HFBBD81: SetPixel hdc, 10, tmph + 14, &HFFC286: SetPixel hdc, 11, tmph + 14, &HFAC386: SetPixel hdc, 12, tmph + 14, &HFBC28A: SetPixel hdc, 13, tmph + 14, &HFAC38A: SetPixel hdc, 14, tmph + 14, &HFAC18D: SetPixel hdc, 15, tmph + 14, &HFDC592: SetPixel hdc, 16, tmph + 14, &HFDC592: SetPixel hdc, 17, tmph + 14, &HFCC592:
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &HA2A0A0: SetPixel hdc, 3, tmph + 15, &H816753: SetPixel hdc, 4, tmph + 15, &HC09068: SetPixel hdc, 5, tmph + 15, &HEDA55F: SetPixel hdc, 6, tmph + 15, &HFAB26C: SetPixel hdc, 7, tmph + 15, &HFCBF7D: SetPixel hdc, 8, tmph + 15, &HF7C182: SetPixel hdc, 9, tmph + 15, &HF8C38A: SetPixel hdc, 10, tmph + 15, &HFACA90: SetPixel hdc, 11, tmph + 15, &HF7CB8E: SetPixel hdc, 12, tmph + 15, &HF8CC8F: SetPixel hdc, 13, tmph + 15, &HFACC96: SetPixel hdc, 14, tmph + 15, &HF9CB95: SetPixel hdc, 15, tmph + 15, &HF9CE97: SetPixel hdc, 16, tmph + 15, &HF8CD97: SetPixel hdc, 17, tmph + 15, &HF8CE97:
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD6D6D6: SetPixel hdc, 3, tmph + 16, &H8E7C6F: SetPixel hdc, 4, tmph + 16, &H946843: SetPixel hdc, 5, tmph + 16, &HEEA762: SetPixel hdc, 6, tmph + 16, &HFFB771: SetPixel hdc, 7, tmph + 16, &HFEC17F: SetPixel hdc, 8, tmph + 16, &HFFC98A: SetPixel hdc, 9, tmph + 16, &HFFCE95: SetPixel hdc, 10, tmph + 16, &HFBCB91: SetPixel hdc, 11, tmph + 16, &HFFD396: SetPixel hdc, 12, tmph + 16, &HFFD396: SetPixel hdc, 13, tmph + 16, &HFFD29C: SetPixel hdc, 14, tmph + 16, &HFFD39D: SetPixel hdc, 15, tmph + 16, &HFFD49E: SetPixel hdc, 16, tmph + 16, &HFFD49E: SetPixel hdc, 17, tmph + 16, &HFED59E:
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBEBEBE: SetPixel hdc, 4, tmph + 17, &H6C6C6C: SetPixel hdc, 5, tmph + 17, &H7C684F: SetPixel hdc, 6, tmph + 17, &HD1AE81: SetPixel hdc, 7, tmph + 17, &HF1C284: SetPixel hdc, 8, tmph + 17, &HFDCE90: SetPixel hdc, 9, tmph + 17, &HF8D193: SetPixel hdc, 10, tmph + 17, &HFBD899: SetPixel hdc, 11, tmph + 17, &HF5DC9E: SetPixel hdc, 12, tmph + 17, &HF8DFA1: SetPixel hdc, 13, tmph + 17, &HF8DFA1: SetPixel hdc, 14, tmph + 17, &HF8DFA1: SetPixel hdc, 15, tmph + 17, &HF8DEA3: SetPixel hdc, 16, tmph + 17, &HF7DDA3: SetPixel hdc, 17, tmph + 17, &HF7DDA3:
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HBABABA: SetPixel hdc, 5, tmph + 18, &H827666: SetPixel hdc, 6, tmph + 18, &H836743: SetPixel hdc, 7, tmph + 18, &HBE935B: SetPixel hdc, 8, tmph + 18, &HF4C78B: SetPixel hdc, 9, tmph + 18, &HFDD79A: SetPixel hdc, 10, tmph + 18, &HFFDFA0: SetPixel hdc, 11, tmph + 18, &HFBE2A4: SetPixel hdc, 12, tmph + 18, &HFFE7A9: SetPixel hdc, 13, tmph + 18, &HFFE9AB: SetPixel hdc, 14, tmph + 18, &HFFE7A9: SetPixel hdc, 15, tmph + 18, &HFFE6AC: SetPixel hdc, 16, tmph + 18, &HFFE6AD: SetPixel hdc, 17, tmph + 18, &HFFE6AD:
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8F8F8F: SetPixel hdc, 7, tmph + 19, &H686462: SetPixel hdc, 8, tmph + 19, &H6D655E: SetPixel hdc, 9, tmph + 19, &H918472: SetPixel hdc, 10, tmph + 19, &HB3A88E: SetPixel hdc, 11, tmph + 19, &HDAD1B2: SetPixel hdc, 12, tmph + 19, &HE3DBBA: SetPixel hdc, 13, tmph + 19, &HE7E0C0: SetPixel hdc, 14, tmph + 19, &HE9E2C1: SetPixel hdc, 15, tmph + 19, &HE9E2C5: SetPixel hdc, 16, tmph + 19, &HE9E1C5: SetPixel hdc, 17, tmph + 19, &HE9E2C5:
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC2C2C1: SetPixel hdc, 8, tmph + 20, &H9F9D9B: SetPixel hdc, 9, tmph + 20, &H827D75: SetPixel hdc, 10, tmph + 20, &H6A6353: SetPixel hdc, 11, tmph + 20, &H5F5941: SetPixel hdc, 12, tmph + 20, &H5D553B: SetPixel hdc, 13, tmph + 20, &H595338: SetPixel hdc, 14, tmph + 20, &H5E5739: SetPixel hdc, 15, tmph + 20, &H5F5A3C: SetPixel hdc, 16, tmph + 20, &H635E3F: SetPixel hdc, 17, tmph + 20, &H635D40:
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5:
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC:
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H450608: SetPixel hdc, tmpw + 18, 0, &H450608: SetPixel hdc, tmpw + 19, 0, &H3B0707: SetPixel hdc, tmpw + 20, 0, &H370706: SetPixel hdc, tmpw + 21, 0, &H360507: SetPixel hdc, tmpw + 22, 0, &H3B0F10: SetPixel hdc, tmpw + 23, 0, &H442526: SetPixel hdc, tmpw + 24, 0, &H604E4E: SetPixel hdc, tmpw + 25, 0, &HA29D9E: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 1, &HF1D4C9: SetPixel hdc, tmpw + 18, 1, &HF1D4C9: SetPixel hdc, tmpw + 19, 1, &HEDD3CD: SetPixel hdc, tmpw + 20, 1, &HEBD1CB: SetPixel hdc, tmpw + 21, 1, &HE9CEC4: SetPixel hdc, tmpw + 22, 1, &HE5C1B9: SetPixel hdc, tmpw + 23, 1, &HCFA89F: SetPixel hdc, tmpw + 24, 1, &HAA6E68: SetPixel hdc, tmpw + 25, 1, &H73211B: SetPixel hdc, tmpw + 26, 1, &H702924: SetPixel hdc, tmpw + 27, 1, &HAA9897: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 2, &HE5C8BD: SetPixel hdc, tmpw + 18, 2, &HE5C8BD: SetPixel hdc, tmpw + 19, 2, &HDEC4BE: SetPixel hdc, tmpw + 20, 2, &HDCC2BC: SetPixel hdc, tmpw + 21, 2, &HE2C7BD: SetPixel hdc, tmpw + 22, 2, &HE2BEB6: SetPixel hdc, tmpw + 23, 2, &HE8C1B8: SetPixel hdc, tmpw + 24, 2, &HF0B4AE: SetPixel hdc, tmpw + 25, 2, &HF29C96: SetPixel hdc, tmpw + 26, 2, &H822D27: SetPixel hdc, tmpw + 27, 2, &H400807: SetPixel hdc, tmpw + 28, 2, &H71585A: SetPixel hdc, tmpw + 29, 2, &HEBEBEB: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 3, &HE8C0A1: SetPixel hdc, tmpw + 18, 3, &HE8C0A1: SetPixel hdc, tmpw + 19, 3, &HE5C09A: SetPixel hdc, tmpw + 20, 3, &HE4BF99: SetPixel hdc, tmpw + 21, 3, &HE4BA97: SetPixel hdc, tmpw + 22, 3, &HE9BF9C: SetPixel hdc, tmpw + 23, 3, &HDFB695: SetPixel hdc, tmpw + 24, 3, &HDFB695: SetPixel hdc, tmpw + 25, 3, &HE0AE90: SetPixel hdc, tmpw + 26, 3, &HCB8469: SetPixel hdc, tmpw + 27, 3, &H941600: SetPixel hdc, tmpw + 28, 3, &H830800: SetPixel hdc, tmpw + 29, 3, &H895253: SetPixel hdc, tmpw + 30, 3, &HF0EFEF: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 4, &HE0B898: SetPixel hdc, tmpw + 18, 4, &HE0B897: SetPixel hdc, tmpw + 19, 4, &HDAB58F: SetPixel hdc, tmpw + 20, 4, &HDBB690: SetPixel hdc, tmpw + 21, 4, &HDBB18E: SetPixel hdc, tmpw + 22, 4, &HD7AD8A: SetPixel hdc, tmpw + 23, 4, &HDAB190: SetPixel hdc, tmpw + 24, 4, &HD2A988: SetPixel hdc, tmpw + 25, 4, &HD6A486: SetPixel hdc, tmpw + 26, 4, &HDA9378: SetPixel hdc, tmpw + 27, 4, &HBF4129: SetPixel hdc, tmpw + 28, 4, &H991B03: SetPixel hdc, tmpw + 29, 4, &H500709: SetPixel hdc, tmpw + 30, 4, &H826F70: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 5, &HE3B48E: SetPixel hdc, tmpw + 18, 5, &HE3B48D: SetPixel hdc, tmpw + 19, 5, &HE0B387: SetPixel hdc, tmpw + 20, 5, &HDEB185: SetPixel hdc, tmpw + 21, 5, &HE1B084: SetPixel hdc, tmpw + 22, 5, &HE3AE83: SetPixel hdc, tmpw + 23, 5, &HE1AF7B: SetPixel hdc, tmpw + 24, 5, &HE0A976: SetPixel hdc, tmpw + 25, 5, &HDCA473: SetPixel hdc, tmpw + 26, 5, &HDEA372: SetPixel hdc, tmpw + 27, 5, &HCC712E: SetPixel hdc, tmpw + 28, 5, &HA53900: SetPixel hdc, tmpw + 29, 5, &H9D2200: SetPixel hdc, tmpw + 30, 5, &H9E2114: SetPixel hdc, tmpw + 31, 5, &HC7C5C4: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 6, &HE0B18B: SetPixel hdc, tmpw + 18, 6, &HE0B18A: SetPixel hdc, tmpw + 19, 6, &HDEB185: SetPixel hdc, tmpw + 20, 6, &HDEB185: SetPixel hdc, tmpw + 21, 6, &HDCAB7F: SetPixel hdc, tmpw + 22, 6, &HE1AC81: SetPixel hdc, tmpw + 23, 6, &HDCAA76: SetPixel hdc, tmpw + 24, 6, &HDCA572: SetPixel hdc, tmpw + 25, 6, &HDBA372: SetPixel hdc, tmpw + 26, 6, &HD79C6B: SetPixel hdc, tmpw + 27, 6, &HD17633: SetPixel hdc, tmpw + 28, 6, &HB74B0B: SetPixel hdc, tmpw + 29, 6, &HAC310D: SetPixel hdc, tmpw + 30, 6, &H961507: SetPixel hdc, tmpw + 31, 6, &H736D6A: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 7, &HE9B47F: SetPixel hdc, tmpw + 18, 7, &HEAB47E: SetPixel hdc, tmpw + 19, 7, &HEFB67E: SetPixel hdc, tmpw + 20, 7, &HE8AF77: SetPixel hdc, tmpw + 21, 7, &HE7AF74: SetPixel hdc, tmpw + 22, 7, &HE4AC71: SetPixel hdc, tmpw + 23, 7, &HEAAD6F: SetPixel hdc, tmpw + 24, 7, &HE9A968: SetPixel hdc, tmpw + 25, 7, &HE7A564: SetPixel hdc, tmpw + 26, 7, &HD9904C: SetPixel hdc, tmpw + 27, 7, &HC5711F: SetPixel hdc, tmpw + 28, 7, &HC16010: SetPixel hdc, tmpw + 29, 7, &HBB4D05: SetPixel hdc, tmpw + 30, 7, &HA02D00: SetPixel hdc, tmpw + 31, 7, &H774033: SetPixel hdc, tmpw + 32, 7, &HE7E6E6: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 8, &HCE9963: SetPixel hdc, tmpw + 18, 8, &HCF9963: SetPixel hdc, tmpw + 19, 8, &HCE955D: SetPixel hdc, tmpw + 20, 8, &HCE955D: SetPixel hdc, tmpw + 21, 8, &HCA9257: SetPixel hdc, tmpw + 22, 8, &HC89055: SetPixel hdc, tmpw + 23, 8, &HCB8E50: SetPixel hdc, tmpw + 24, 8, &HCB8B4A: SetPixel hdc, tmpw + 25, 8, &HC58342: SetPixel hdc, tmpw + 26, 8, &HC87F3B: SetPixel hdc, tmpw + 27, 8, &HCA7624: SetPixel hdc, tmpw + 28, 8, &HCA6919: SetPixel hdc, tmpw + 29, 8, &HCC5E16: SetPixel hdc, tmpw + 30, 8, &HB23E07: SetPixel hdc, tmpw + 31, 8, &H682B1D: SetPixel hdc, tmpw + 32, 8, &HC7C2C2: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 9, &HDDA064: SetPixel hdc, tmpw + 18, 9, &HDCA064: SetPixel hdc, tmpw + 19, 9, &HDA9D5D: SetPixel hdc, tmpw + 20, 9, &HD99C5C: SetPixel hdc, tmpw + 21, 9, &HDA9D5D: SetPixel hdc, tmpw + 22, 9, &HDA9A5A: SetPixel hdc, tmpw + 23, 9, &HD89753: SetPixel hdc, tmpw + 24, 9, &HD7914E: SetPixel hdc, tmpw + 25, 9, &HD38E49: SetPixel hdc, tmpw + 26, 9, &HD38B43: SetPixel hdc, tmpw + 27, 9, &HCD8430: SetPixel hdc, tmpw + 28, 9, &HCA7826: SetPixel hdc, tmpw + 29, 9, &HCE6C1E: SetPixel hdc, tmpw + 30, 9, &HB9560C: SetPixel hdc, tmpw + 31, 9, &H742E0D: SetPixel hdc, tmpw + 32, 9, &HB3A6A4: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 10, &HE2A66A: SetPixel hdc, tmpw + 18, 10, &HE2A66A: SetPixel hdc, tmpw + 19, 10, &HE1A464: SetPixel hdc, tmpw + 20, 10, &HE0A363: SetPixel hdc, tmpw + 21, 10, &HE0A363: SetPixel hdc, tmpw + 22, 10, &HE1A161: SetPixel hdc, tmpw + 23, 10, &HE09F5B: SetPixel hdc, tmpw + 24, 10, &HDE9855: SetPixel hdc, tmpw + 25, 10, &HDC9752: SetPixel hdc, tmpw + 26, 10, &HDB934B: SetPixel hdc, tmpw + 27, 10, &HD68D39: SetPixel hdc, tmpw + 28, 10, &HD17F2D: SetPixel hdc, tmpw + 29, 10, &HD67426: SetPixel hdc, tmpw + 30, 10, &HC05D13: SetPixel hdc, tmpw + 31, 10, &H7C3514: SetPixel hdc, tmpw + 32, 10, &HAB9B98: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF:
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 10, &HE2A66A: SetPixel hdc, tmpw + 18, tmph + 10, &HE2A66A: SetPixel hdc, tmpw + 19, tmph + 10, &HE1A464: SetPixel hdc, tmpw + 20, tmph + 10, &HE0A363: SetPixel hdc, tmpw + 21, tmph + 10, &HE0A363: SetPixel hdc, tmpw + 22, tmph + 10, &HE1A161: SetPixel hdc, tmpw + 23, tmph + 10, &HE09F5B: SetPixel hdc, tmpw + 24, tmph + 10, &HDE9855: SetPixel hdc, tmpw + 25, tmph + 10, &HDC9752: SetPixel hdc, tmpw + 26, tmph + 10, &HDB934B: SetPixel hdc, tmpw + 27, tmph + 10, &HD68D39: SetPixel hdc, tmpw + 28, tmph + 10, &HD17F2D: SetPixel hdc, tmpw + 29, tmph + 10, &HD67426: SetPixel hdc, tmpw + 30, tmph + 10, &HC05D13: SetPixel hdc, tmpw + 31, tmph + 10, &H7C3514: SetPixel hdc, tmpw + 32, tmph + 10, &HAB9B98: SetPixel hdc, tmpw + 33, tmph + 10, &HF6F6F6:
    SetPixel hdc, tmpw + 17, tmph + 11, &HE6AC76: SetPixel hdc, tmpw + 18, tmph + 11, &HE6AC76: SetPixel hdc, tmpw + 19, tmph + 11, &HE2A86D: SetPixel hdc, tmpw + 20, tmph + 11, &HE5A66C: SetPixel hdc, tmpw + 21, tmph + 11, &HE1A56A: SetPixel hdc, tmpw + 22, tmph + 11, &HE4A46A: SetPixel hdc, tmpw + 23, tmph + 11, &HE1A266: SetPixel hdc, tmpw + 24, tmph + 11, &HE6A364: SetPixel hdc, tmpw + 25, tmph + 11, &HE19F5E: SetPixel hdc, tmpw + 26, tmph + 11, &HDF9A55: SetPixel hdc, tmpw + 27, tmph + 11, &HD89048: SetPixel hdc, tmpw + 28, tmph + 11, &HD88A3E: SetPixel hdc, tmpw + 29, tmph + 11, &HCF7927: SetPixel hdc, tmpw + 30, tmph + 11, &HC87220: SetPixel hdc, tmpw + 31, tmph + 11, &H77481E: SetPixel hdc, tmpw + 32, tmph + 11, &HABA39E: SetPixel hdc, tmpw + 33, tmph + 11, &HF4F4F4:
    SetPixel hdc, tmpw + 17, tmph + 12, &HF1B681: SetPixel hdc, tmpw + 18, tmph + 12, &HF0B780: SetPixel hdc, tmpw + 19, tmph + 12, &HF2B87D: SetPixel hdc, tmpw + 20, tmph + 12, &HF5B67C: SetPixel hdc, tmpw + 21, tmph + 12, &HF1B57A: SetPixel hdc, tmpw + 22, tmph + 12, &HF2B278: SetPixel hdc, tmpw + 23, tmph + 12, &HF0B175: SetPixel hdc, tmpw + 24, tmph + 12, &HF3B071: SetPixel hdc, tmpw + 25, tmph + 12, &HECAA69: SetPixel hdc, tmpw + 26, tmph + 12, &HE9A45F: SetPixel hdc, tmpw + 27, tmph + 12, &HE8A058: SetPixel hdc, tmpw + 28, tmph + 12, &HE5974B: SetPixel hdc, tmpw + 29, tmph + 12, &HE38D3B: SetPixel hdc, tmpw + 30, tmph + 12, &HD37D2B: SetPixel hdc, tmpw + 31, tmph + 12, &H895A32: SetPixel hdc, tmpw + 32, tmph + 12, &HB4AFAC: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5:
    SetPixel hdc, tmpw + 17, tmph + 13, &HF3BD8A: SetPixel hdc, tmpw + 18, tmph + 13, &HF3BD8A: SetPixel hdc, tmpw + 19, tmph + 13, &HF2BD84: SetPixel hdc, tmpw + 20, tmph + 13, &HF5BC84: SetPixel hdc, tmpw + 21, tmph + 13, &HF3BC83: SetPixel hdc, tmpw + 22, tmph + 13, &HF4B981: SetPixel hdc, tmpw + 23, tmph + 13, &HF2B97C: SetPixel hdc, tmpw + 24, tmph + 13, &HF5B77B: SetPixel hdc, tmpw + 25, tmph + 13, &HF1B476: SetPixel hdc, tmpw + 26, tmph + 13, &HEFAF6E: SetPixel hdc, tmpw + 27, tmph + 13, &HE5A45F: SetPixel hdc, tmpw + 28, tmph + 13, &HE49F5A: SetPixel hdc, tmpw + 29, tmph + 13, &HDA8F4A: SetPixel hdc, tmpw + 30, tmph + 13, &HC57A35: SetPixel hdc, tmpw + 31, tmph + 13, &H736353: SetPixel hdc, tmpw + 32, tmph + 13, &HD6D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8:
    SetPixel hdc, tmpw + 17, tmph + 14, &HFCC592: SetPixel hdc, tmpw + 18, tmph + 14, &HFBC592: SetPixel hdc, tmpw + 19, tmph + 14, &HF7C289: SetPixel hdc, tmpw + 20, tmph + 14, &HFCC38B: SetPixel hdc, tmpw + 21, tmph + 14, &HFAC38A: SetPixel hdc, tmpw + 22, tmph + 14, &HFDC28A: SetPixel hdc, tmpw + 23, tmph + 14, &HFBC285: SetPixel hdc, tmpw + 24, tmph + 14, &HFBBD81: SetPixel hdc, tmpw + 25, tmph + 14, &HF6B97B: SetPixel hdc, tmpw + 26, tmph + 14, &HF6B675: SetPixel hdc, tmpw + 27, tmph + 14, &HF0AF6A: SetPixel hdc, tmpw + 28, tmph + 14, &HE8A35E: SetPixel hdc, tmpw + 29, tmph + 14, &HDD924D: SetPixel hdc, tmpw + 30, tmph + 14, &HBA702B: SetPixel hdc, tmpw + 31, tmph + 14, &H847A70: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 15, &HF8CE97: SetPixel hdc, tmpw + 18, tmph + 15, &HF9CD97: SetPixel hdc, tmpw + 19, tmph + 15, &HF9CE95: SetPixel hdc, tmpw + 20, tmph + 15, &HF7CC93: SetPixel hdc, tmpw + 21, tmph + 15, &HF6CB92: SetPixel hdc, tmpw + 22, tmph + 15, &HF9CA92: SetPixel hdc, tmpw + 23, tmph + 15, &HFCCD90: SetPixel hdc, tmpw + 24, tmph + 15, &HF8C488: SetPixel hdc, tmpw + 25, tmph + 15, &HF3BD80: SetPixel hdc, tmpw + 26, tmph + 15, &HFABD7D: SetPixel hdc, tmpw + 27, tmph + 15, &HF7B26D: SetPixel hdc, tmpw + 28, tmph + 15, &HEAA560: SetPixel hdc, tmpw + 29, tmph + 15, &HC0925D: SetPixel hdc, tmpw + 30, tmph + 15, &H896F54: SetPixel hdc, tmpw + 31, tmph + 15, &HBABABB: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1:
    SetPixel hdc, tmpw + 17, tmph + 16, &HFED59E: SetPixel hdc, tmpw + 18, tmph + 16, &HFFD59F: SetPixel hdc, tmpw + 19, tmph + 16, &HFED39A: SetPixel hdc, tmpw + 20, tmph + 16, &HFFD49B: SetPixel hdc, tmpw + 21, tmph + 16, &HFCD198: SetPixel hdc, tmpw + 22, tmph + 16, &HFFD098: SetPixel hdc, tmpw + 23, tmph + 16, &HFECF92: SetPixel hdc, tmpw + 24, tmph + 16, &HFFCB8F: SetPixel hdc, tmpw + 25, tmph + 16, &HFFC98C: SetPixel hdc, tmpw + 26, tmph + 16, &HFEC181: SetPixel hdc, tmpw + 27, tmph + 16, &HFBB671: SetPixel hdc, tmpw + 28, tmph + 16, &HF0AB66: SetPixel hdc, tmpw + 29, tmph + 16, &H9F733E: SetPixel hdc, tmpw + 30, tmph + 16, &H918478: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9:
    SetPixel hdc, tmpw + 17, tmph + 17, &HF7DDA3: SetPixel hdc, tmpw + 18, tmph + 17, &HF8DDA4: SetPixel hdc, tmpw + 19, tmph + 17, &HF9E0A2: SetPixel hdc, tmpw + 20, tmph + 17, &HF5DC9E: SetPixel hdc, tmpw + 21, tmph + 17, &HF8DEA2: SetPixel hdc, tmpw + 22, tmph + 17, &HFBDDA2: SetPixel hdc, tmpw + 23, tmph + 17, &HF7D495: SetPixel hdc, tmpw + 24, tmph + 17, &HF8D193: SetPixel hdc, tmpw + 25, tmph + 17, &HFCCD90: SetPixel hdc, tmpw + 26, tmph + 17, &HF1C088: SetPixel hdc, tmpw + 27, tmph + 17, &HDBB186: SetPixel hdc, tmpw + 28, tmph + 17, &H8C7259: SetPixel hdc, tmpw + 29, tmph + 17, &H6D6B6B: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE:
    SetPixel hdc, tmpw + 17, tmph + 18, &HFFE6AD: SetPixel hdc, tmpw + 18, tmph + 18, &HFFE6AD: SetPixel hdc, tmpw + 19, tmph + 18, &HFFE7A9: SetPixel hdc, tmpw + 20, tmph + 18, &HFFEAAC: SetPixel hdc, tmpw + 21, tmph + 18, &HF7DDA1: SetPixel hdc, tmpw + 22, tmph + 18, &HFFE1A6: SetPixel hdc, tmpw + 23, tmph + 18, &HFFE1A2: SetPixel hdc, tmpw + 24, tmph + 18, &HFED799: SetPixel hdc, tmpw + 25, tmph + 18, &HFACC8F: SetPixel hdc, tmpw + 26, tmph + 18, &HC99A64: SetPixel hdc, tmpw + 27, tmph + 18, &H977048: SetPixel hdc, tmpw + 28, tmph + 18, &H817060: SetPixel hdc, tmpw + 29, tmph + 18, &HC8C8C8: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC:
    SetPixel hdc, tmpw + 17, tmph + 19, &HE9E2C5: SetPixel hdc, tmpw + 18, tmph + 19, &HE9E2C5: SetPixel hdc, tmpw + 19, tmph + 19, &HEAE2C4: SetPixel hdc, tmpw + 20, tmph + 19, &HE7DFC1: SetPixel hdc, tmpw + 21, tmph + 19, &HEEE4C6: SetPixel hdc, tmpw + 22, tmph + 19, &HDBD1B4: SetPixel hdc, tmpw + 23, tmph + 19, &HB7AF93: SetPixel hdc, tmpw + 24, tmph + 19, &H8D8973: SetPixel hdc, tmpw + 25, tmph + 19, &H736D60: SetPixel hdc, tmpw + 26, tmph + 19, &H6A6660: SetPixel hdc, tmpw + 27, tmph + 19, &H8E9090: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 20, &H635D40: SetPixel hdc, tmpw + 18, tmph + 20, &H615B3F: SetPixel hdc, tmpw + 19, tmph + 20, &H60583C: SetPixel hdc, tmpw + 20, tmph + 20, &H5D563A: SetPixel hdc, tmpw + 21, tmph + 20, &H61583D: SetPixel hdc, tmpw + 22, tmph + 20, &H605840: SetPixel hdc, tmpw + 23, tmph + 20, &H6A6556: SetPixel hdc, tmpw + 24, tmph + 20, &H7F7D75: SetPixel hdc, tmpw + 25, tmph + 20, &HA4A3A1: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD:
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    'Generar lineas intermedias
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HB0A09E: DrawLineApi 2, tmph, 2, tmph1, &H712E13: DrawLineApi 3, tmph, 3, tmph1, &HBD5F14:
    DrawLineApi 4, tmph, 4, tmph1, &HD17327: DrawLineApi 5, tmph, 5, tmph1, &HD47F31: DrawLineApi 6, tmph, 6, tmph1, &HD98C3D: DrawLineApi 7, tmph, 7, tmph1, &HD9944B:
    DrawLineApi 8, tmph, 8, tmph1, &HD7944F: DrawLineApi 9, tmph, 9, tmph1, &HDC9C55: DrawLineApi 10, tmph, 10, tmph1, &HDC9B57: DrawLineApi 11, tmph, 11, tmph1, &HE3A362:
    DrawLineApi 12, tmph, 12, tmph1, &HE3A265: DrawLineApi 13, tmph, 13, tmph1, &HE2A367: DrawLineApi 14, tmph, 14, tmph1, &HE0A165: DrawLineApi 15, tmph, 15, tmph1, &HE3A66A:
    DrawLineApi 16, tmph, 16, tmph1, &HE3A66A: DrawLineApi 17, tmph, 17, tmph1, &HE2A66A:
    DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HE2A66A: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HE2A66A: DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HE1A464:
    DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HE0A363: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HE0A363: DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HE1A161
    DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HE09F5B: DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HDE9855: DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HDC9752:
    DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HDB934B: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HD68D39: DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HD17F2D:
    DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HD67426: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HC05D13: DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &H7C3514:
    DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &HAB9B98: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6:
    'Lineas verticales
    DrawLineApi 17, 0, lw - 17, 0, &H450608
    DrawLineApi 17, 1, lw - 17, 1, &HF1D4C9
    DrawLineApi 17, 2, lw - 17, 2, &HE5C8BD
    DrawLineApi 17, 3, lw - 17, 3, &HE8C0A1
    DrawLineApi 17, 4, lw - 17, 4, &HE0B898
    DrawLineApi 17, 5, lw - 17, 5, &HE3B48E
    DrawLineApi 17, 6, lw - 17, 6, &HE0B18B
    DrawLineApi 17, 7, lw - 17, 7, &HE9B47F
    DrawLineApi 17, 8, lw - 17, 8, &HCE9963
    DrawLineApi 17, 9, lw - 17, 9, &HDDA064
    DrawLineApi 17, 10, lw - 17, 10, &HE2A66A
    DrawLineApi 17, 11, lw - 17, 11, &HE6AC76
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HE6AC76
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HF1B681
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HF3BD8A
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HFCC592
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HF8CE97
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HFED59E
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HF7DDA3
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HFFE6AD
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HE9E2C5
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H635D40
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub
Private Sub DrawAquaDown()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
        
    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HCC9B6A, lpRect
    
    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE5E4E4: SetPixel hdc, 8, 0, &HA5A2A2: SetPixel hdc, 9, 0, &H675C5C: SetPixel hdc, 10, 0, &H422729: SetPixel hdc, 11, 0, &H300E0D: SetPixel hdc, 12, 0, &H300A09: SetPixel hdc, 13, 0, &H2F0908: SetPixel hdc, 14, 0, &H330909: SetPixel hdc, 15, 0, &H390A0A: SetPixel hdc, 16, 0, &H3C0A0A: SetPixel hdc, 17, 0, &H3C090A:
    SetPixel hdc, 5, 1, &HF0EEEE: SetPixel hdc, 6, 1, &H9D888A: SetPixel hdc, 7, 1, &H653531: SetPixel hdc, 8, 1, &H5A201D: SetPixel hdc, 9, 1, &H8D655F: SetPixel hdc, 10, 1, &HB99995: SetPixel hdc, 11, 1, &HD0B4B2: SetPixel hdc, 12, 1, &HD7BEBB: SetPixel hdc, 13, 1, &HDDC6C0: SetPixel hdc, 14, 1, &HDDC6C0: SetPixel hdc, 15, 1, &HDDC7BE: SetPixel hdc, 16, 1, &HDDC7BE: SetPixel hdc, 17, 1, &HDEC7BE:
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE4E4E4: SetPixel hdc, 5, 2, &H6F5C5C: SetPixel hdc, 6, 2, &H390A0E: SetPixel hdc, 7, 2, &H712E2A: SetPixel hdc, 8, 2, &HD6928D: SetPixel hdc, 9, 2, &HD8ACA6: SetPixel hdc, 10, 2, &HD1B0AC: SetPixel hdc, 11, 2, &HD1B5B2: SetPixel hdc, 12, 2, &HD0B7B4: SetPixel hdc, 13, 2, &HCEB7B1: SetPixel hdc, 14, 2, &HCEB7B1: SetPixel hdc, 15, 2, &HD2BCB2: SetPixel hdc, 16, 2, &HD2BCB2: SetPixel hdc, 17, 2, &HD3BCB2:
    SetPixel hdc, 3, 3, &HEEEDED: SetPixel hdc, 4, 3, &H805858: SetPixel hdc, 5, 3, &H6A0D08: SetPixel hdc, 6, 3, &H7D1909: SetPixel hdc, 7, 3, &HB07B63: SetPixel hdc, 8, 3, &HCFA58A: SetPixel hdc, 9, 3, &HCDA78E: SetPixel hdc, 10, 3, &HD1AB92: SetPixel hdc, 11, 3, &HD2AF93: SetPixel hdc, 12, 3, &HD3B094: SetPixel hdc, 13, 3, &HD0AF93: SetPixel hdc, 14, 3, &HD3B296: SetPixel hdc, 15, 3, &HD4B49A: SetPixel hdc, 16, 3, &HD4B39A: SetPixel hdc, 17, 3, &HD4B39A:
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H837576: SetPixel hdc, 4, 4, &H440C0C: SetPixel hdc, 5, 4, &H821D0D: SetPixel hdc, 6, 4, &HA94433: SetPixel hdc, 7, 4, &HC08B72: SetPixel hdc, 8, 4, &HC49A7F: SetPixel hdc, 9, 4, &HC6A188: SetPixel hdc, 10, 4, &HC7A189: SetPixel hdc, 11, 4, &HC6A387: SetPixel hdc, 12, 4, &HC8A689: SetPixel hdc, 13, 4, &HC9A98C: SetPixel hdc, 14, 4, &HC8A88B: SetPixel hdc, 15, 4, &HCBAB91: SetPixel hdc, 16, 4, &HCCAC92: SetPixel hdc, 17, 4, &HCCAC92:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCAC8C7: SetPixel hdc, 3, 5, &H79281D: SetPixel hdc, 4, 5, &H7F2409: SetPixel hdc, 5, 5, &H8C3809: SetPixel hdc, 6, 5, &HBD6D39: SetPixel hdc, 7, 5, &HC9986E: SetPixel hdc, 8, 5, &HC89D74: SetPixel hdc, 9, 5, &HC49A71: SetPixel hdc, 10, 5, &HCAA17C: SetPixel hdc, 11, 5, &HC6A07A: SetPixel hdc, 12, 5, &HCAA480: SetPixel hdc, 13, 5, &HCAA582: SetPixel hdc, 14, 5, &HCBA584: SetPixel hdc, 15, 5, &HCDA989: SetPixel hdc, 16, 5, &HCFA98A: SetPixel hdc, 17, 5, &HCFA88A:
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H756C6B: SetPixel hdc, 3, 6, &H76190D: SetPixel hdc, 4, 6, &H913416: SetPixel hdc, 5, 6, &H9D4916: SetPixel hdc, 6, 6, &HBA6A36: SetPixel hdc, 7, 6, &HC39268: SetPixel hdc, 8, 6, &HC59A71: SetPixel hdc, 9, 6, &HC59B72: SetPixel hdc, 10, 6, &HC59C77: SetPixel hdc, 11, 6, &HC6A07A: SetPixel hdc, 12, 6, &HC6A07C: SetPixel hdc, 13, 6, &HC7A27F: SetPixel hdc, 14, 6, &HCBA584: SetPixel hdc, 15, 6, &HCAA686: SetPixel hdc, 16, 6, &HCBA586: SetPixel hdc, 17, 6, &HCCA587:
    SetPixel hdc, 1, 7, &HE8E7E7: SetPixel hdc, 2, 7, &H6C3E35: SetPixel hdc, 3, 7, &H8A2D09: SetPixel hdc, 4, 7, &HA34812: SetPixel hdc, 5, 7, &HAB591A: SetPixel hdc, 6, 7, &HB46B2B: SetPixel hdc, 7, 7, &HC3854A: SetPixel hdc, 8, 7, &HD19C64: SetPixel hdc, 9, 7, &HCD9C6C: SetPixel hdc, 10, 7, &HD1A070: SetPixel hdc, 11, 7, &HD2A272: SetPixel hdc, 12, 7, &HD2A272: SetPixel hdc, 13, 7, &HD6A57A: SetPixel hdc, 14, 7, &HD8A77C: SetPixel hdc, 15, 7, &HD2A87C: SetPixel hdc, 16, 7, &HD2A87C: SetPixel hdc, 17, 7, &HD2A77D:
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HC7C3C3: SetPixel hdc, 2, 8, &H5C2A21: SetPixel hdc, 3, 8, &H9C3E15: SetPixel hdc, 4, 8, &HB35A22: SetPixel hdc, 5, 8, &HB56324: SetPixel hdc, 6, 8, &HB66D2D: SetPixel hdc, 7, 8, &HB6783D: SetPixel hdc, 8, 8, &HB07B44: SetPixel hdc, 9, 8, &HB18050: SetPixel hdc, 10, 8, &HB58454: SetPixel hdc, 11, 8, &HB48554: SetPixel hdc, 12, 8, &HB78858: SetPixel hdc, 13, 8, &HBA895E: SetPixel hdc, 14, 8, &HBB8A5F: SetPixel hdc, 15, 8, &HB98E62: SetPixel hdc, 16, 8, &HB98E62: SetPixel hdc, 17, 8, &HB98E62:
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HB4ABA9: SetPixel hdc, 2, 9, &H612A14: SetPixel hdc, 3, 9, &HA05316: SetPixel hdc, 4, 9, &HB36628: SetPixel hdc, 5, 9, &HB67132: SetPixel hdc, 6, 9, &HB67738: SetPixel hdc, 7, 9, &HB98146: SetPixel hdc, 8, 9, &HBD864E: SetPixel hdc, 9, 9, &HBD894F: SetPixel hdc, 10, 9, &HC28D55: SetPixel hdc, 11, 9, &HC4905B: SetPixel hdc, 12, 9, &HC5905F: SetPixel hdc, 13, 9, &HC49161: SetPixel hdc, 14, 9, &HC49161: SetPixel hdc, 15, 9, &HC69564: SetPixel hdc, 16, 9, &HC69564: SetPixel hdc, 17, 9, &HC69464:
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HA99D9B: SetPixel hdc, 2, 10, &H632D17: SetPixel hdc, 3, 10, &HA65A1D: SetPixel hdc, 4, 10, &HB96C2E: SetPixel hdc, 5, 10, &HBC7738: SetPixel hdc, 6, 10, &HC18242: SetPixel hdc, 7, 10, &HC2894E: SetPixel hdc, 8, 10, &HC18A52: SetPixel hdc, 9, 10, &HC59157: SetPixel hdc, 10, 10, &HC59159: SetPixel hdc, 11, 10, &HCC9863: SetPixel hdc, 12, 10, &HCC9665: SetPixel hdc, 13, 10, &HCB9767: SetPixel hdc, 14, 10, &HC99565: SetPixel hdc, 15, 10, &HCC9A6A: SetPixel hdc, 16, 10, &HCC9A6A: SetPixel hdc, 17, 10, &HCC9B6A:
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 10, &HF7F7F7: SetPixel hdc, 1, tmph + 10, &HA99D9B: SetPixel hdc, 2, tmph + 10, &H632D17: SetPixel hdc, 3, tmph + 10, &HA65A1D: SetPixel hdc, 4, tmph + 10, &HB96C2E: SetPixel hdc, 5, tmph + 10, &HBC7738: SetPixel hdc, 6, tmph + 10, &HC18242: SetPixel hdc, 7, tmph + 10, &HC2894E: SetPixel hdc, 8, tmph + 10, &HC18A52: SetPixel hdc, 9, tmph + 10, &HC59157: SetPixel hdc, 10, tmph + 10, &HC59159: SetPixel hdc, 11, tmph + 10, &HCC9863: SetPixel hdc, 12, tmph + 10, &HCC9665: SetPixel hdc, 13, tmph + 10, &HCB9767: SetPixel hdc, 14, tmph + 10, &HC99565: SetPixel hdc, 15, tmph + 10, &HCC9A6A: SetPixel hdc, 16, tmph + 10, &HCC9A6A: SetPixel hdc, 17, tmph + 10, &HCC9B6A:
    SetPixel hdc, 0, tmph + 11, &HF5F5F5: SetPixel hdc, 1, tmph + 11, &HA59F9A: SetPixel hdc, 2, tmph + 11, &H674024: SetPixel hdc, 3, tmph + 11, &HAE6827: SetPixel hdc, 4, tmph + 11, &HB97231: SetPixel hdc, 5, tmph + 11, &HBE8247: SetPixel hdc, 6, tmph + 11, &HC0874E: SetPixel hdc, 7, tmph + 11, &HC78E56: SetPixel hdc, 8, tmph + 11, &HCD9561: SetPixel hdc, 9, tmph + 11, &HCB9466: SetPixel hdc, 10, tmph + 11, &HCD9A6B: SetPixel hdc, 11, tmph + 11, &HC79867: SetPixel hdc, 12, tmph + 11, &HCA9B6A: SetPixel hdc, 13, tmph + 11, &HCC9D6C: SetPixel hdc, 14, tmph + 11, &HCD9D70: SetPixel hdc, 15, tmph + 11, &HD0A175: SetPixel hdc, 16, tmph + 11, &HD0A175: SetPixel hdc, 17, tmph + 11, &HD0A175:
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HACA7A4: SetPixel hdc, 2, tmph + 12, &H755035: SetPixel hdc, 3, tmph + 12, &HB77131: SetPixel hdc, 4, tmph + 12, &HCB8443: SetPixel hdc, 5, tmph + 12, &HC5894E: SetPixel hdc, 6, tmph + 12, &HCC935A: SetPixel hdc, 7, tmph + 12, &HD29962: SetPixel hdc, 8, tmph + 12, &HD69F6A: SetPixel hdc, 9, tmph + 12, &HDBA476: SetPixel hdc, 10, tmph + 12, &HD6A374: SetPixel hdc, 11, tmph + 12, &HD4A574: SetPixel hdc, 12, tmph + 12, &HD8A978: SetPixel hdc, 13, tmph + 12, &HDAAB7A: SetPixel hdc, 14, tmph + 12, &HDAAA7D: SetPixel hdc, 15, tmph + 12, &HDBAB7F: SetPixel hdc, 16, tmph + 12, &HDAAA7F: SetPixel hdc, 17, tmph + 12, &HDAAA7F:
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC0C0BF: SetPixel hdc, 2, tmph + 13, &H63574B: SetPixel hdc, 3, tmph + 13, &HAC7036: SetPixel hdc, 4, tmph + 13, &HC2854A: SetPixel hdc, 5, tmph + 13, &HCF955E: SetPixel hdc, 6, tmph + 13, &HD29B66: SetPixel hdc, 7, tmph + 13, &HD1A26E: SetPixel hdc, 8, tmph + 13, &HD8A776: SetPixel hdc, 9, tmph + 13, &HDBA878: SetPixel hdc, 10, tmph + 13, &HDFAC7C: SetPixel hdc, 11, tmph + 13, &HDBAF7D: SetPixel hdc, 12, tmph + 13, &HDDAF81: SetPixel hdc, 13, tmph + 13, &HDEB183: SetPixel hdc, 14, tmph + 13, &HDDAF84: SetPixel hdc, 15, tmph + 13, &HDEB087: SetPixel hdc, 16, tmph + 13, &HDEB087: SetPixel hdc, 17, tmph + 13, &HDCB087:
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H7C7269: SetPixel hdc, 3, tmph + 14, &HA26830: SetPixel hdc, 4, tmph + 14, &HC6884E: SetPixel hdc, 5, tmph + 14, &HD0965F: SetPixel hdc, 6, tmph + 14, &HDAA26E: SetPixel hdc, 7, tmph + 14, &HD9AA75: SetPixel hdc, 8, tmph + 14, &HDBAA79: SetPixel hdc, 9, tmph + 14, &HE2AF7F: SetPixel hdc, 10, tmph + 14, &HE6B484: SetPixel hdc, 11, tmph + 14, &HE2B684: SetPixel hdc, 12, tmph + 14, &HE3B588: SetPixel hdc, 13, tmph + 14, &HE2B587: SetPixel hdc, 14, tmph + 14, &HE2B48A: SetPixel hdc, 15, tmph + 14, &HE5B78E: SetPixel hdc, 16, tmph + 14, &HE5B78E: SetPixel hdc, 17, tmph + 14, &HE4B88E:
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &H9E9C9C: SetPixel hdc, 3, tmph + 15, &H766051: SetPixel hdc, 4, tmph + 15, &HAD8666: SetPixel hdc, 5, tmph + 15, &HD49A61: SetPixel hdc, 6, tmph + 15, &HE0A66D: SetPixel hdc, 7, tmph + 15, &HE3B17C: SetPixel hdc, 8, tmph + 15, &HE0B380: SetPixel hdc, 9, tmph + 15, &HE0B587: SetPixel hdc, 10, tmph + 15, &HE2BC8C: SetPixel hdc, 11, tmph + 15, &HE0BB8B: SetPixel hdc, 12, tmph + 15, &HE0BC8B: SetPixel hdc, 13, tmph + 15, &HE3BD92: SetPixel hdc, 14, tmph + 15, &HE2BC91: SetPixel hdc, 15, tmph + 15, &HE2BF93: SetPixel hdc, 16, tmph + 15, &HE1BE93: SetPixel hdc, 17, tmph + 15, &HE1BF93:
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD5D5D5: SetPixel hdc, 3, tmph + 16, &H86766C: SetPixel hdc, 4, tmph + 16, &H856144: SetPixel hdc, 5, tmph + 16, &HD59C63: SetPixel hdc, 6, tmph + 16, &HE5AB71: SetPixel hdc, 7, tmph + 16, &HE5B37E: SetPixel hdc, 8, tmph + 16, &HE7BB88: SetPixel hdc, 9, tmph + 16, &HE7BF91: SetPixel hdc, 10, tmph + 16, &HE3BC8D: SetPixel hdc, 11, tmph + 16, &HE7C392: SetPixel hdc, 12, tmph + 16, &HE7C392: SetPixel hdc, 13, tmph + 16, &HE8C398: SetPixel hdc, 14, tmph + 16, &HE8C499: SetPixel hdc, 15, tmph + 16, &HE8C599: SetPixel hdc, 16, tmph + 16, &HE8C599: SetPixel hdc, 17, tmph + 16, &HE7C699:
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBDBDBD: SetPixel hdc, 4, tmph + 17, &H676767: SetPixel hdc, 5, tmph + 17, &H71604C: SetPixel hdc, 6, tmph + 17, &HBEA17D: SetPixel hdc, 7, tmph + 17, &HDAB381: SetPixel hdc, 8, tmph + 17, &HE5BE8C: SetPixel hdc, 9, tmph + 17, &HE1C18F: SetPixel hdc, 10, tmph + 17, &HE4C895: SetPixel hdc, 11, tmph + 17, &HDFCA98: SetPixel hdc, 12, tmph + 17, &HE2CE9B: SetPixel hdc, 13, tmph + 17, &HE2CE9B: SetPixel hdc, 14, tmph + 17, &HE2CE9B: SetPixel hdc, 15, tmph + 17, &HE2CD9D: SetPixel hdc, 16, tmph + 17, &HE2CC9D: SetPixel hdc, 17, tmph + 17, &HE2CC9D:
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HB9B9B9: SetPixel hdc, 5, tmph + 18, &H7A7163: SetPixel hdc, 6, tmph + 18, &H776043: SetPixel hdc, 7, tmph + 18, &HAB885B: SetPixel hdc, 8, tmph + 18, &HDDB888: SetPixel hdc, 9, tmph + 18, &HE6C796: SetPixel hdc, 10, tmph + 18, &HE8CD9A: SetPixel hdc, 11, tmph + 18, &HE5D19E: SetPixel hdc, 12, tmph + 18, &HE9D6A3: SetPixel hdc, 13, tmph + 18, &HE9D6A5: SetPixel hdc, 14, tmph + 18, &HE9D6A3: SetPixel hdc, 15, tmph + 18, &HE9D5A6: SetPixel hdc, 16, tmph + 18, &HE9D5A6: SetPixel hdc, 17, tmph + 18, &HE9D5A6:
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8C8C8C: SetPixel hdc, 7, tmph + 19, &H61605E: SetPixel hdc, 8, tmph + 19, &H656059: SetPixel hdc, 9, tmph + 19, &H857C6D: SetPixel hdc, 10, tmph + 19, &HA59C87: SetPixel hdc, 11, tmph + 19, &HC8C1A8: SetPixel hdc, 12, tmph + 19, &HD1CAB0: SetPixel hdc, 13, tmph + 19, &HD5CFB5: SetPixel hdc, 14, tmph + 19, &HD6D1B6: SetPixel hdc, 15, tmph + 19, &HD7D2BA: SetPixel hdc, 16, tmph + 19, &HD7D1BA: SetPixel hdc, 17, tmph + 19, &HD7D2BA:
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC1C1C1: SetPixel hdc, 8, tmph + 20, &H9C9B99: SetPixel hdc, 9, tmph + 20, &H7D7A73: SetPixel hdc, 10, tmph + 20, &H635E50: SetPixel hdc, 11, tmph + 20, &H58533F: SetPixel hdc, 12, tmph + 20, &H554F39: SetPixel hdc, 13, tmph + 20, &H514D36: SetPixel hdc, 14, tmph + 20, &H554F37: SetPixel hdc, 15, tmph + 20, &H57523A: SetPixel hdc, 16, tmph + 20, &H5A563D: SetPixel hdc, 17, tmph + 20, &H5A563E:
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5:
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC:
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H3C090A: SetPixel hdc, tmpw + 18, 0, &H3C090A: SetPixel hdc, tmpw + 19, 0, &H340A0A: SetPixel hdc, tmpw + 20, 0, &H300A09: SetPixel hdc, tmpw + 21, 0, &H2F080A: SetPixel hdc, tmpw + 22, 0, &H341011: SetPixel hdc, tmpw + 23, 0, &H3E2526: SetPixel hdc, tmpw + 24, 0, &H5A4C4C: SetPixel hdc, tmpw + 25, 0, &H9E9B9B: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 1, &HDEC7BE: SetPixel hdc, tmpw + 18, 1, &HDEC7BE: SetPixel hdc, tmpw + 19, 1, &HDBC6C1: SetPixel hdc, tmpw + 20, 1, &HD9C4BF: SetPixel hdc, tmpw + 21, 1, &HD7C1B9: SetPixel hdc, tmpw + 22, 1, &HD3B5AF: SetPixel hdc, tmpw + 23, 1, &HBE9F97: SetPixel hdc, tmpw + 24, 1, &H9B6A65: SetPixel hdc, tmpw + 25, 1, &H65231E: SetPixel hdc, tmpw + 26, 1, &H642A26: SetPixel hdc, tmpw + 27, 1, &HA59696: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 2, &HD3BCB2: SetPixel hdc, tmpw + 18, 2, &HD3BCB2: SetPixel hdc, tmpw + 19, 2, &HCDB8B3: SetPixel hdc, tmpw + 20, 2, &HCBB6B1: SetPixel hdc, tmpw + 21, 2, &HD0BBB2: SetPixel hdc, tmpw + 22, 2, &HD0B2AC: SetPixel hdc, tmpw + 23, 2, &HD6B6AF: SetPixel hdc, tmpw + 24, 2, &HDCABA6: SetPixel hdc, tmpw + 25, 2, &HDC9691: SetPixel hdc, tmpw + 26, 2, &H732E29: SetPixel hdc, tmpw + 27, 2, &H380A0A: SetPixel hdc, tmpw + 28, 2, &H6A5556: SetPixel hdc, tmpw + 29, 2, &HEAEBEA: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 3, &HD4B39A: SetPixel hdc, tmpw + 18, 3, &HD4B39A: SetPixel hdc, tmpw + 19, 3, &HD1B294: SetPixel hdc, tmpw + 20, 3, &HD0B193: SetPixel hdc, tmpw + 21, 3, &HD0AE91: SetPixel hdc, tmpw + 22, 3, &HD4B296: SetPixel hdc, tmpw + 23, 3, &HCBAA8F: SetPixel hdc, tmpw + 24, 3, &HCBAA8F: SetPixel hdc, tmpw + 25, 3, &HCCA38B: SetPixel hdc, tmpw + 26, 3, &HB77E68: SetPixel hdc, tmpw + 27, 3, &H811B09: SetPixel hdc, tmpw + 28, 3, &H720E08: SetPixel hdc, tmpw + 29, 3, &H7D5051: SetPixel hdc, tmpw + 30, 3, &HEFEEEE: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 4, &HCCAC92: SetPixel hdc, tmpw + 18, 4, &HCCAC91: SetPixel hdc, tmpw + 19, 4, &HC6A889: SetPixel hdc, tmpw + 20, 4, &HC7A98A: SetPixel hdc, tmpw + 21, 4, &HC7A589: SetPixel hdc, tmpw + 22, 4, &HC4A185: SetPixel hdc, tmpw + 23, 4, &HC6A58A: SetPixel hdc, tmpw + 24, 4, &HBF9E83: SetPixel hdc, tmpw + 25, 4, &HC39A82: SetPixel hdc, tmpw + 26, 4, &HC58C76: SetPixel hdc, tmpw + 27, 4, &HA9432F: SetPixel hdc, tmpw + 28, 4, &H861F0C: SetPixel hdc, tmpw + 29, 4, &H460B0C: SetPixel hdc, tmpw + 30, 4, &H7B6B6C: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 5, &HCFA88A: SetPixel hdc, tmpw + 18, 5, &HCFA889: SetPixel hdc, tmpw + 19, 5, &HCBA683: SetPixel hdc, tmpw + 20, 5, &HC9A481: SetPixel hdc, tmpw + 21, 5, &HCCA480: SetPixel hdc, tmpw + 22, 5, &HCEA280: SetPixel hdc, tmpw + 23, 5, &HCCA379: SetPixel hdc, tmpw + 24, 5, &HCA9E74: SetPixel hdc, tmpw + 25, 5, &HC69971: SetPixel hdc, tmpw + 26, 5, &HC89870: SetPixel hdc, tmpw + 27, 5, &HB46A34: SetPixel hdc, tmpw + 28, 5, &H90380A: SetPixel hdc, tmpw + 29, 5, &H892509: SetPixel hdc, tmpw + 30, 5, &H8A251B: SetPixel hdc, tmpw + 31, 5, &HC4C2C2: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 6, &HCCA587: SetPixel hdc, tmpw + 18, 6, &HCCA586: SetPixel hdc, tmpw + 19, 6, &HC9A481: SetPixel hdc, tmpw + 20, 6, &HC9A481: SetPixel hdc, tmpw + 21, 6, &HC7A07C: SetPixel hdc, tmpw + 22, 6, &HCCA17E: SetPixel hdc, tmpw + 23, 6, &HC79F74: SetPixel hdc, tmpw + 24, 6, &HC69A70: SetPixel hdc, tmpw + 25, 6, &HC59870: SetPixel hdc, tmpw + 26, 6, &HC2926A: SetPixel hdc, tmpw + 27, 6, &HB96F39: SetPixel hdc, tmpw + 28, 6, &HA04814: SetPixel hdc, tmpw + 29, 6, &H973215: SetPixel hdc, tmpw + 30, 6, &H831A0F: SetPixel hdc, tmpw + 31, 6, &H6E6966: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 7, &HD2A77D: SetPixel hdc, tmpw + 18, 7, &HD3A77C: SetPixel hdc, tmpw + 19, 7, &HD8AA7D: SetPixel hdc, tmpw + 20, 7, &HD2A376: SetPixel hdc, tmpw + 21, 7, &HD1A373: SetPixel hdc, tmpw + 22, 7, &HCEA070: SetPixel hdc, tmpw + 23, 7, &HD2A06F: SetPixel hdc, tmpw + 24, 7, &HD19D68: SetPixel hdc, tmpw + 25, 7, &HD09A65: SetPixel hdc, tmpw + 26, 7, &HC2864F: SetPixel hdc, tmpw + 27, 7, &HAE6927: SetPixel hdc, tmpw + 28, 7, &HA95A19: SetPixel hdc, tmpw + 29, 7, &HA44A10: SetPixel hdc, tmpw + 30, 7, &H8B2E09: SetPixel hdc, tmpw + 31, 7, &H6B3E34: SetPixel hdc, tmpw + 32, 7, &HE7E6E6: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 8, &HB98E62: SetPixel hdc, tmpw + 18, 8, &HBA8E62: SetPixel hdc, tmpw + 19, 8, &HB98B5E: SetPixel hdc, tmpw + 20, 8, &HB98B5E: SetPixel hdc, tmpw + 21, 8, &HB68858: SetPixel hdc, tmpw + 22, 8, &HB48656: SetPixel hdc, tmpw + 23, 8, &HB58452: SetPixel hdc, tmpw + 24, 8, &HB5814C: SetPixel hdc, tmpw + 25, 8, &HB07A46: SetPixel hdc, tmpw + 26, 8, &HB2773F: SetPixel hdc, tmpw + 27, 8, &HB36E2C: SetPixel hdc, tmpw + 28, 8, &HB26221: SetPixel hdc, tmpw + 29, 8, &HB35A20: SetPixel hdc, tmpw + 30, 8, &H9C3E11: SetPixel hdc, tmpw + 31, 8, &H5C2A1F: SetPixel hdc, tmpw + 32, 8, &HC4C1C0: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 9, &HC69464: SetPixel hdc, tmpw + 18, 9, &HC69564: SetPixel hdc, tmpw + 19, 9, &HC3925E: SetPixel hdc, tmpw + 20, 9, &HC3915D: SetPixel hdc, tmpw + 21, 9, &HC3925E: SetPixel hdc, tmpw + 22, 9, &HC38F5B: SetPixel hdc, tmpw + 23, 9, &HC28D55: SetPixel hdc, tmpw + 24, 9, &HC08751: SetPixel hdc, tmpw + 25, 9, &HBC844C: SetPixel hdc, tmpw + 26, 9, &HBC8147: SetPixel hdc, tmpw + 27, 9, &HB57936: SetPixel hdc, tmpw + 28, 9, &HB3702D: SetPixel hdc, tmpw + 29, 9, &HB56626: SetPixel hdc, tmpw + 30, 9, &HA25115: SetPixel hdc, tmpw + 31, 9, &H662D12: SetPixel hdc, tmpw + 32, 9, &HAEA3A1: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 10, &HCC9B6A: SetPixel hdc, tmpw + 18, 10, &HCC9B6A: SetPixel hdc, tmpw + 19, 10, &HCA9864: SetPixel hdc, tmpw + 20, 10, &HC99763: SetPixel hdc, tmpw + 21, 10, &HC99763: SetPixel hdc, tmpw + 22, 10, &HCA9562: SetPixel hdc, tmpw + 23, 10, &HC9945D: SetPixel hdc, tmpw + 24, 10, &HC68E57: SetPixel hdc, tmpw + 25, 10, &HC48C55: SetPixel hdc, tmpw + 26, 10, &HC3884E: SetPixel hdc, tmpw + 27, 10, &HBE823E: SetPixel hdc, tmpw + 28, 10, &HB97634: SetPixel hdc, tmpw + 29, 10, &HBD6D2D: SetPixel hdc, tmpw + 30, 10, &HA8581C: SetPixel hdc, tmpw + 31, 10, &H6D3319: SetPixel hdc, tmpw + 32, 10, &HA49794: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF:
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 10, &HCC9B6A: SetPixel hdc, tmpw + 18, tmph + 10, &HCC9B6A: SetPixel hdc, tmpw + 19, tmph + 10, &HCA9864: SetPixel hdc, tmpw + 20, tmph + 10, &HC99763: SetPixel hdc, tmpw + 21, tmph + 10, &HC99763: SetPixel hdc, tmpw + 22, tmph + 10, &HCA9562: SetPixel hdc, tmpw + 23, tmph + 10, &HC9945D: SetPixel hdc, tmpw + 24, tmph + 10, &HC68E57: SetPixel hdc, tmpw + 25, tmph + 10, &HC48C55: SetPixel hdc, tmpw + 26, tmph + 10, &HC3884E: SetPixel hdc, tmpw + 27, tmph + 10, &HBE823E: SetPixel hdc, tmpw + 28, tmph + 10, &HB97634: SetPixel hdc, tmpw + 29, tmph + 10, &HBD6D2D: SetPixel hdc, tmpw + 30, tmph + 10, &HA8581C: SetPixel hdc, tmpw + 31, tmph + 10, &H6D3319: SetPixel hdc, tmpw + 32, tmph + 10, &HA49794: SetPixel hdc, tmpw + 33, tmph + 10, &HF6F6F6:
    SetPixel hdc, tmpw + 17, tmph + 11, &HD0A175: SetPixel hdc, tmpw + 18, tmph + 11, &HD0A175: SetPixel hdc, tmpw + 19, tmph + 11, &HCC9D6D: SetPixel hdc, tmpw + 20, tmph + 11, &HCE9B6C: SetPixel hdc, tmpw + 21, tmph + 11, &HCB9A6A: SetPixel hdc, tmpw + 22, tmph + 11, &HCD996A: SetPixel hdc, tmpw + 23, tmph + 11, &HCA9666: SetPixel hdc, tmpw + 24, tmph + 11, &HCF9865: SetPixel hdc, tmpw + 25, tmph + 11, &HCA9460: SetPixel hdc, tmpw + 26, tmph + 11, &HC78F57: SetPixel hdc, tmpw + 27, tmph + 11, &HC1864B: SetPixel hdc, tmpw + 28, tmph + 11, &HC08143: SetPixel hdc, tmpw + 29, tmph + 11, &HB7712E: SetPixel hdc, tmpw + 30, tmph + 11, &HB16A28: SetPixel hdc, tmpw + 31, tmph + 11, &H694321: SetPixel hdc, tmpw + 32, tmph + 11, &HA59F9B: SetPixel hdc, tmpw + 33, tmph + 11, &HF4F4F4:
    SetPixel hdc, tmpw + 17, tmph + 12, &HDAAA7F: SetPixel hdc, tmpw + 18, tmph + 12, &HD9AB7E: SetPixel hdc, tmpw + 19, tmph + 12, &HDBAC7C: SetPixel hdc, tmpw + 20, tmph + 12, &HDDAA7B: SetPixel hdc, tmpw + 21, tmph + 12, &HDAA979: SetPixel hdc, tmpw + 22, tmph + 12, &HDAA677: SetPixel hdc, tmpw + 23, tmph + 12, &HD8A474: SetPixel hdc, tmpw + 24, tmph + 12, &HDBA471: SetPixel hdc, tmpw + 25, tmph + 12, &HD49F6A: SetPixel hdc, tmpw + 26, tmph + 12, &HD09861: SetPixel hdc, tmpw + 27, tmph + 12, &HD0955A: SetPixel hdc, tmpw + 28, tmph + 12, &HCC8D4F: SetPixel hdc, tmpw + 29, tmph + 12, &HCA8441: SetPixel hdc, tmpw + 30, tmph + 12, &HBB7532: SetPixel hdc, tmpw + 31, tmph + 12, &H7B5434: SetPixel hdc, tmpw + 32, tmph + 12, &HB1ACAA: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5:
    SetPixel hdc, tmpw + 17, tmph + 13, &HDCB087: SetPixel hdc, tmpw + 18, tmph + 13, &HDCB087: SetPixel hdc, tmpw + 19, tmph + 13, &HDBAF81: SetPixel hdc, tmpw + 20, tmph + 13, &HDEAF82: SetPixel hdc, tmpw + 21, tmph + 13, &HDCAF81: SetPixel hdc, tmpw + 22, tmph + 13, &HDDAD7F: SetPixel hdc, tmpw + 23, tmph + 13, &HDBAC7B: SetPixel hdc, tmpw + 24, tmph + 13, &HDDAA7A: SetPixel hdc, tmpw + 25, tmph + 13, &HD9A775: SetPixel hdc, tmpw + 26, tmph + 13, &HD7A26E: SetPixel hdc, tmpw + 27, tmph + 13, &HCE9961: SetPixel hdc, tmpw + 28, tmph + 13, &HCC945C: SetPixel hdc, tmpw + 29, tmph + 13, &HC2854D: SetPixel hdc, tmpw + 30, tmph + 13, &HAF7239: SetPixel hdc, tmpw + 31, tmph + 13, &H695C4F: SetPixel hdc, tmpw + 32, tmph + 13, &HD5D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8:
    SetPixel hdc, tmpw + 17, tmph + 14, &HE4B88E: SetPixel hdc, tmpw + 18, tmph + 14, &HE3B88E: SetPixel hdc, tmpw + 19, tmph + 14, &HE0B486: SetPixel hdc, tmpw + 20, tmph + 14, &HE4B689: SetPixel hdc, tmpw + 21, tmph + 14, &HE2B587: SetPixel hdc, tmpw + 22, tmph + 14, &HE5B588: SetPixel hdc, tmpw + 23, tmph + 14, &HE3B483: SetPixel hdc, tmpw + 24, tmph + 14, &HE2AF7F: SetPixel hdc, tmpw + 25, tmph + 14, &HDEAC7A: SetPixel hdc, tmpw + 26, tmph + 14, &HDEAA75: SetPixel hdc, tmpw + 27, tmph + 14, &HD8A36B: SetPixel hdc, tmpw + 28, tmph + 14, &HD09860: SetPixel hdc, tmpw + 29, tmph + 14, &HC58850: SetPixel hdc, tmpw + 30, tmph + 14, &HA56930: SetPixel hdc, tmpw + 31, tmph + 14, &H7B746C: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 15, &HE1BF93: SetPixel hdc, tmpw + 18, tmph + 15, &HE2BE93: SetPixel hdc, tmpw + 19, tmph + 15, &HE2BF91: SetPixel hdc, tmpw + 20, tmph + 15, &HE1BD8F: SetPixel hdc, tmpw + 21, tmph + 15, &HE0BC8E: SetPixel hdc, tmpw + 22, tmph + 15, &HE2BC8E: SetPixel hdc, tmpw + 23, tmph + 15, &HE4BD8C: SetPixel hdc, tmpw + 24, tmph + 15, &HE0B685: SetPixel hdc, tmpw + 25, tmph + 15, &HDCB07E: SetPixel hdc, tmpw + 26, tmph + 15, &HE1AF7C: SetPixel hdc, tmpw + 27, tmph + 15, &HDEA66E: SetPixel hdc, tmpw + 28, tmph + 15, &HD19962: SetPixel hdc, tmpw + 29, tmph + 15, &HAD875D: SetPixel hdc, tmpw + 30, tmph + 15, &H7D6851: SetPixel hdc, tmpw + 31, tmph + 15, &HB9B9B9: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1:
    SetPixel hdc, tmpw + 17, tmph + 16, &HE7C699: SetPixel hdc, tmpw + 18, tmph + 16, &HE8C69A: SetPixel hdc, tmpw + 19, tmph + 16, &HE7C496: SetPixel hdc, tmpw + 20, tmph + 16, &HE8C597: SetPixel hdc, tmpw + 21, tmph + 16, &HE5C294: SetPixel hdc, tmpw + 22, tmph + 16, &HE8C194: SetPixel hdc, tmpw + 23, tmph + 16, &HE6BF8E: SetPixel hdc, tmpw + 24, tmph + 16, &HE7BC8C: SetPixel hdc, tmpw + 25, tmph + 16, &HE7BB8A: SetPixel hdc, tmpw + 26, tmph + 16, &HE5B37F: SetPixel hdc, tmpw + 27, tmph + 16, &HE1A971: SetPixel hdc, tmpw + 28, tmph + 16, &HD79F67: SetPixel hdc, tmpw + 29, tmph + 16, &H8E6A40: SetPixel hdc, tmpw + 30, tmph + 16, &H8A8076: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9:
    SetPixel hdc, tmpw + 17, tmph + 17, &HE2CC9D: SetPixel hdc, tmpw + 18, tmph + 17, &HE2CC9E: SetPixel hdc, tmpw + 19, tmph + 17, &HE3CF9C: SetPixel hdc, tmpw + 20, tmph + 17, &HDFCA98: SetPixel hdc, tmpw + 21, tmph + 17, &HE2CD9C: SetPixel hdc, tmpw + 22, tmph + 17, &HE4CC9C: SetPixel hdc, tmpw + 23, tmph + 17, &HE1C491: SetPixel hdc, tmpw + 24, tmph + 17, &HE1C18F: SetPixel hdc, tmpw + 25, tmph + 17, &HE4BD8C: SetPixel hdc, tmpw + 26, tmph + 17, &HDAB285: SetPixel hdc, tmpw + 27, tmph + 17, &HC7A582: SetPixel hdc, tmpw + 28, tmph + 17, &H806A56: SetPixel hdc, tmpw + 29, tmph + 17, &H676565: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE:
    SetPixel hdc, tmpw + 17, tmph + 18, &HE9D5A6: SetPixel hdc, tmpw + 18, tmph + 18, &HE9D5A6: SetPixel hdc, tmpw + 19, tmph + 18, &HE9D6A3: SetPixel hdc, tmpw + 20, tmph + 18, &HE9D7A6: SetPixel hdc, tmpw + 21, tmph + 18, &HE2CC9B: SetPixel hdc, tmpw + 22, tmph + 18, &HE8D0A0: SetPixel hdc, tmpw + 23, tmph + 18, &HE8CF9C: SetPixel hdc, tmpw + 24, tmph + 18, &HE7C795: SetPixel hdc, tmpw + 25, tmph + 18, &HE2BC8B: SetPixel hdc, tmpw + 26, tmph + 18, &HB68F63: SetPixel hdc, tmpw + 27, tmph + 18, &H886948: SetPixel hdc, tmpw + 28, tmph + 18, &H786A5D: SetPixel hdc, tmpw + 29, tmph + 18, &HC7C7C7: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC:
    SetPixel hdc, tmpw + 17, tmph + 19, &HD7D2BA: SetPixel hdc, tmpw + 18, tmph + 19, &HD7D2BA: SetPixel hdc, tmpw + 19, tmph + 19, &HD7D1B9: SetPixel hdc, tmpw + 20, tmph + 19, &HD5CEB6: SetPixel hdc, tmpw + 21, tmph + 19, &HDBD3BB: SetPixel hdc, tmpw + 22, tmph + 19, &HC9C1AA: SetPixel hdc, tmpw + 23, tmph + 19, &HA9A28B: SetPixel hdc, tmpw + 24, tmph + 19, &H827E6C: SetPixel hdc, tmpw + 25, tmph + 19, &H6A665B: SetPixel hdc, tmpw + 26, tmph + 19, &H625F5A: SetPixel hdc, tmpw + 27, tmph + 19, &H8B8C8C: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 20, &H5A563E: SetPixel hdc, tmpw + 18, tmph + 20, &H59543D: SetPixel hdc, tmpw + 19, tmph + 20, &H58513A: SetPixel hdc, tmpw + 20, tmph + 20, &H554F38: SetPixel hdc, tmpw + 21, tmph + 20, &H59513B: SetPixel hdc, tmpw + 22, tmph + 20, &H58513E: SetPixel hdc, tmpw + 23, tmph + 20, &H646053: SetPixel hdc, tmpw + 24, tmph + 20, &H7B7973: SetPixel hdc, tmpw + 25, tmph + 20, &HA2A19F: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD:
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    'Generar lineas intermedias
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HA99D9B: DrawLineApi 2, tmph, 2, tmph1, &H632D17: DrawLineApi 3, tmph, 3, tmph1, &HA65A1D: DrawLineApi 4, tmph, 4, tmph1, &HB96C2E
    DrawLineApi 5, tmph, 5, tmph1, &HBC7738: DrawLineApi 6, tmph, 6, tmph1, &HC18242: DrawLineApi 7, tmph, 7, tmph1, &HC2894E: DrawLineApi 8, tmph, 8, tmph1, &HC18A52: DrawLineApi 9, tmph, 9, tmph1, &HC59157
    DrawLineApi 10, tmph, 10, tmph1, &HC59159: DrawLineApi 11, tmph, 11, tmph1, &HCC9863: DrawLineApi 12, tmph, 12, tmph1, &HCC9665: DrawLineApi 13, tmph, 13, tmph1, &HCB9767: DrawLineApi 14, tmph, 14, tmph1, &HC99565
    DrawLineApi 15, tmph, 15, tmph1, &HCC9A6A: DrawLineApi 16, tmph, 16, tmph1, &HCC9A6A: DrawLineApi 17, tmph, 17, tmph1, &HCC9B6A: DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HCC9B6A: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HCC9B6A:
    DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HCA9864: DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HC99763: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HC99763: DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HCA9562: DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HC9945D
    DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HC68E57: DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HC48C55: DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HC3884E: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HBE823E: DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HB97634
    DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HBD6D2D: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HA8581C: DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &H6D3319: DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &HA49794: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
    'Lineas verticales
    DrawLineApi 17, 0, lw - 17, 0, &H3C090A
    DrawLineApi 17, 1, lw - 17, 1, &HDEC7BE
    DrawLineApi 17, 2, lw - 17, 2, &HD3BCB2
    DrawLineApi 17, 3, lw - 17, 3, &HD4B39A
    DrawLineApi 17, 4, lw - 17, 4, &HCCAC92
    DrawLineApi 17, 5, lw - 17, 5, &HCFA88A
    DrawLineApi 17, 6, lw - 17, 6, &HCCA587
    DrawLineApi 17, 7, lw - 17, 7, &HD2A77D
    DrawLineApi 17, 8, lw - 17, 8, &HB98E62
    DrawLineApi 17, 9, lw - 17, 9, &HC69464
    DrawLineApi 17, 10, lw - 17, 10, &HCC9B6A
    DrawLineApi 17, 11, lw - 17, 11, &HD0A175
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HD0A175
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HDAAA7F
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HDCB087
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HE4B88E
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HE1BF93
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HE7C699
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HE2CC9D
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HE9D5A6
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HD7D2BA
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H5A563E
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawOffice2003(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim bColor As Long

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    bColor = TranslateColor(m_bColors.tBackColor)
    SetRect m_ButtonRect, 0, 0, lw, lh
    
    If m_bCheckBoxMode And m_bValue Then
        If m_bMouseInCtl Then
            DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
        Else
            DrawGradientEx 0, 0, lw, lh, TranslateColor(&H8CD5FF), TranslateColor(&H55ADFF), gdVertical
        End If
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
    Exit Sub
    End If

    Select Case vState

    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(bColor, 0.05), bColor, gdVertical
        DrawGradientEx 0, lh / 2, lw, lh / 2 + 1, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.13), gdVertical
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, TranslateColor(&HCCF4FF), TranslateColor(&H91D0FF), gdVertical
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh, TranslateColor(&H4E91FE), TranslateColor(&H8ED3FF), gdVertical
    End Select

    If m_Buttonstate <> eStateNormal Then
        DrawRectangle 0, 0, lw, lh, TranslateColor(&H800000)
    End If
    
End Sub

Private Sub DrawSleekButton(ByVal vState As enumButtonStates)

Dim lpRect As RECT
Dim bColor As Long

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    bColor = TranslateColor(m_bColors.tBackColor)
    
    If m_bCheckBoxMode And m_bValue Then
        vState = eStateDown
    End If

    Select Case vState

    Case eStateNormal
        CreateRegion
        DrawGradientEx 0, 0, lw, lh, bColor, ShiftColor(bColor, 0.08), gdHorizontal
    Case eStateOver
        DrawGradientEx 0, 0, lw, lh, bColor, ShiftColor(bColor, 0.08), gdHorizontal
    Case eStateDown
        DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, -0.05), ShiftColor(bColor, 0.03), gdHorizontal
    End Select
    
    DrawRectangle 0, 0, lw, lh, TranslateColor(&H7D807E)
    ' --Button has focus  button is Default
    If m_bHasFocus Or m_bDefault Then
        If m_bShowFocus And Ambient.UserMode Then
            SetRect lpRect, 3, 3, lw - 3, lh - 3
            If m_bParentActive Then
                DrawFocusRect hdc, lpRect
            End If
        End If
    End If


End Sub

Private Sub PaintRegion(ByVal lRgn As Long, ByVal lColor As Long)

'Fills a specified region with specified color

Dim hBrush As Long
Dim hOldBrush As Long

    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(hdc, hBrush)
    
    FillRgn hdc, lRgn, hBrush
    
    SelectObject hdc, hOldBrush
    DeleteObject hBrush
    
End Sub

Private Sub PaintRect(ByVal lColor As Long, lpRect As RECT)

'Fills a region with specified color

Dim hOldBrush   As Long
Dim hBrush      As Long

    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(UserControl.hdc, hBrush)

    FillRect UserControl.hdc, lpRect, hBrush

    SelectObject UserControl.hdc, hOldBrush
    DeleteObject hBrush

End Sub

Private Function ShiftColor(color As Long, PercentInDecimal As Single) As Long

'****************************************************************************
'* This routine shifts a color value specified by PercentInDecimal          *
'* Function inspired from DCbutton                                          *
'* All Credits goes to Noel Dacara                                          *
'* A Littlebit modified by me                                               *
'****************************************************************************

Dim r As Long
Dim g As Long
Dim b As Long

'  Add or remove a certain color quantity by how many percent.

    r = color And 255
    g = (color \ 256) And 255
    b = (color \ 65536) And 255

    r = r + PercentInDecimal * 255       ' Percent should already
    g = g + PercentInDecimal * 255       ' be translated.
    b = b + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5

    '  When overflow occurs, ....
    If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
        If (r > 255) Then r = 255
        If (g > 255) Then g = 255
        If (b > 255) Then b = 255
    Else
        If (r < 0) Then r = 0
        If (g < 0) Then g = 0
        If (b < 0) Then b = 0
    End If

    ShiftColor = r + 256& * g + 65536 * b ' Return shifted color value

End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If m_bEnabled Then                           'Disabled?? get out!!
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_bIsDown = False
        End If
        If m_bCheckBoxMode Then                'Checkbox Mode?
            If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
            m_bValue = Not m_bValue             'Change Value (Checked/Unchecked)
            If Not m_bValue Then                'If value unchecked then
                m_Buttonstate = eStateNormal     'Normal State
            End If
            RedrawButton
        End If
        DoEvents                               'To remove focus from other button and Do events before click event
        RaiseEvent Click                       'Now Raiseevent
    End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    m_bDefault = Ambient.DisplayAsDefault
    If PropertyName = "DisplayAsDefault" Then
        RedrawButton
    End If

    If PropertyName = "BackColor" Then
        RedrawButton
    End If

End Sub

Private Sub UserControl_DblClick()

    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
    
    If m_lDownButton = 1 Then                    'React to only Left button
        
        SetCapture (hWnd)                         'Preserve Hwnd on DoubleClick
        If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        RedrawButton
        UserControl_MouseDown m_lDownButton, m_lDShift, m_lDX, m_lDY
        RaiseEvent DblClick
    End If

End Sub

Private Sub UserControl_GotFocus()

    m_bHasFocus = True
    If m_bMouseInCtl Then
        If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
    Else
        If Not m_bIsDown Then m_Buttonstate = eStateNormal
    End If

End Sub

Private Sub UserControl_Initialize()

Dim OS As OSVERSIONINFO

    ' --Get the operating system version for text drawing purposes.
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    m_WindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()

'Initialize Properties for User Control
'Called on designtime everytime a control is added

    m_ButtonStyle = eStandard
    m_bShowFocus = True
    m_bEnabled = True
    m_Caption = Ambient.DisplayName
    UserControl.FontName = "Tahoma"  'Tahoma
    m_PictureAlign = epLeftOfCaption
    m_bUseMaskColor = False
    m_lMaskColor = &HE0E0E0
    m_CaptionAlign = ecCenterAlign
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    SetThemeColors

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 13                                    'Enter Key
        RaiseEvent Click
    Case 37, 38                                'Left and Up Arrows
        SendKeys "+{TAB}"                      'Button should transfer focus to other ctl
    Case 39, 40                                'Right and Down Arrows
        SendKeys "{TAB}"                       'Button should transfer focus to other ctl
    Case 32                                    'SpaceBar held down
        If Not m_bIsDown Then
            If Shift = 4 Then Exit Sub         'System Menu Should pop up
            m_bIsSpaceBarDown = True           'Set space bar as pressed
            If (m_bCheckBoxMode) Then          'Is CheckBoxMode??
                m_bValue = Not m_bValue        'Toggle Check Value
            Else
                If m_Buttonstate <> eStateDown Then
                    m_Buttonstate = eStateDown 'Button state should be down
                    RedrawButton
                End If
            End If
        End If

        If (Not GetCapture = UserControl.hWnd) Then
            ReleaseCapture
            SetCapture UserControl.hWnd     'No other processing until spacebar is released
        End If                              'Thanks to APIGuide
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeySpace Then
        If m_bMouseInCtl And m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
            RedrawButton
        ElseIf m_bMouseInCtl And Not m_bIsDown Then   'If spacebar released over ctl
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver 'Draw Hover State
            RedrawButton
            RaiseEvent Click
        Else                                         'If Spacebar released outside ctl
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            RedrawButton
            RaiseEvent Click
        End If

        RaiseEvent KeyUp(KeyCode, Shift)
        m_bIsSpaceBarDown = False
        m_bIsDown = False
    End If

End Sub

Private Sub UserControl_LostFocus()

    m_bHasFocus = False                                 'No focus
    m_bIsDown = False                                   'No down state
    m_bIsSpaceBarDown = False                           'No spacebar held
    If Not m_bParentActive Then
        If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
    ElseIf m_bMouseInCtl Then
        If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
    Else
        If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
    End If
    RedrawButton

    If m_bDefault Then                                  'If default button,
        RedrawButton                                    'Show Focus
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_lDownButton = Button                       'Button pressed for Dblclick
    m_lDX = x
    m_lDY = y
    m_lDShift = Shift
        
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
        
    If Button = 1 Then
        m_bHasFocus = True
        m_bIsDown = True
        If m_bMouseInCtl Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        End If
        RedrawButton
        RaiseEvent MouseDown(Button, Shift, x, y)
    End If

End Sub

Private Sub SetThemeColors()

'Sets a style colors to default colors when button initialized
'or whenever you change the style of Button

    With m_bColors

        Select Case m_ButtonStyle

        Case eStandard, eFlat, eVistaToolbar, e3DHover, eFlatHover
            .tBackColor = TranslateColor(vbButtonFace)
        Case eWindowsXP
            .tBackColor = TranslateColor(&HE7EBEC)
        Case eOutlook2007, eGelButton
            .tBackColor = TranslateColor(&HFFD1AD)
            .tForeColor = TranslateColor(&H8B4215)
        Case eXPToolbar
            .tBackColor = TranslateColor(vbButtonFace)
        Case eAOL
            .tBackColor = TranslateColor(&HAA6D00)
            .tForeColor = TranslateColor(vbWhite)
        Case eAqua
            .tBackColor = TranslateColor(&HD0D0D0)
        Case eVistaAero
            .tBackColor = ShiftColor(TranslateColor(&HD4D4D4), 0.06)
        Case eInstallShield
            .tBackColor = TranslateColor(&HE1D6D5)
        Case eVisualStudio
            .tBackColor = TranslateColor(vbButtonFace)
        Case eOffice2003
            .tBackColor = TranslateColor(&HFCE1CA)
        Case eSleek
            .tBackColor = TranslateColor(&HE6CAA7)
        End Select

        If m_ButtonStyle <> eAOL Then .tForeColor = TranslateColor(vbButtonText)
        If m_ButtonStyle = eFlat Or m_ButtonStyle = eSleek Or m_ButtonStyle = eInstallShield Or m_ButtonStyle = eStandard Or m_ButtonStyle = eAOL Then
            m_bShowFocus = True
        Else
            m_bShowFocus = False
        End If

    End With

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim p As POINT

    GetCursorPos p

    If (Not WindowFromPoint(p.x, p.y) = UserControl.hWnd) Then
        m_bMouseInCtl = False
        RaiseEvent MouseLeave
    End If

    TrackMouseLeave UserControl.hWnd

    If m_bMouseInCtl Then
        
        If m_bHandPointer Then
            SetCursor m_lCursor
        End If
        
        If m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        ElseIf Not m_bIsDown And Not m_bIsSpaceBarDown Then
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
        End If
        RedrawButton
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

            
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If
    
    If Button = vbLeftButton Then
        m_bIsDown = False
        
        If (x > 0 And y > 0) And (x < ScaleWidth And y < ScaleHeight) Then
            If m_bCheckBoxMode Then m_bValue = Not m_bValue
            RedrawButton
            RaiseEvent Click
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_Resize()

    ' --At least, a checkbox will also need this much of size!!!!
    If Height < 220 Then Height = 220
    If Width < 220 Then Width = 220
    
    ' --On resize, create button region again
    CreateRegion
    RedrawButton
    
End Sub

Private Sub UserControl_Paint()

    RedrawButton

End Sub

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_ButtonStyle = .ReadProperty("ButtonStyle", eFlat)
        m_bShowFocus = .ReadProperty("ShowFocusRect", False) 'for eFlat style only
        Set mFont = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = mFont
        m_bColors.tBackColor = .ReadProperty("BackColor", TranslateColor(vbButtonFace))
        m_bEnabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", "jcbutton")
        m_bValue = .ReadProperty("Value", False)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0) 'vbdefault
        m_bHandPointer = .ReadProperty("HandPointer", False)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_PicOver = .ReadProperty("PictureHover", Nothing)
        m_PicSize = .ReadProperty("PictureSize", epsNormal)
        m_lMaskColor = .ReadProperty("MaskColor", &HE0E0E0)
        m_bUseMaskColor = .ReadProperty("UseMaskCOlor", False)
        m_bCheckBoxMode = .ReadProperty("CheckBoxMode", False)
        m_PictureAlign = .ReadProperty("PictureAlign", epLeftOfCaption)
        m_CaptionAlign = .ReadProperty("CaptionAlign", ecCenterAlign)
        m_bColors.tForeColor = .ReadProperty("ForeColor", TranslateColor(vbButtonText))
        UserControl.ForeColor = m_bColors.tForeColor
        UserControl.Enabled = m_bEnabled
        SetAccessKey
        lh = UserControl.ScaleHeight
        lw = UserControl.ScaleWidth
        m_lParenthWnd = UserControl.Parent.hWnd
    End With

    UserControl_Resize

    If Ambient.UserMode Then                                                              'If we're not in design mode
        
        If m_bHandPointer Then
            m_lCursor = LoadCursor(0, IDC_HAND)     'Load System Hand pointer
            m_bHandPointer = (Not m_lCursor = 0)
        End If
        
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If

        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Subclass_Start .hWnd
                Subclass_Start m_lParenthWnd
                Subclass_AddMsg .hWnd, WM_MOUSEMOVE, MSG_AFTER
                Subclass_AddMsg .hWnd, WM_MOUSELEAVE, MSG_AFTER
                On Error Resume Next
                If UserControl.Parent.MDIChild Then
                    Call Subclass_AddMsg(m_lParenthWnd, WM_NCACTIVATE, MSG_AFTER)
                Else
                    Call Subclass_AddMsg(m_lParenthWnd, WM_ACTIVATE, MSG_AFTER)
                End If
            End With
        End If
    End If

End Sub

'A nice place to stop subclasser

Private Sub UserControl_Terminate()

On Error GoTo Crash:
    If m_lButtonRgn Then DeleteObject m_lButtonRgn
    If Ambient.UserMode Then
    Subclass_Stop m_lParenthWnd
    Subclass_Stop UserControl.hWnd
    Subclass_StopAll                                                   'Terminate all subclassing
    End If
Crash:

End Sub

'Write property values to storage

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "ButtonStyle", m_ButtonStyle, eFlat
        .WriteProperty "ShowFocusRect", m_bShowFocus, False
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "Font", mFont, Ambient.Font
        .WriteProperty "BackColor", m_bColors.tBackColor, TranslateColor(vbButtonFace)
        .WriteProperty "Caption", m_Caption, "jcbutton1"
        .WriteProperty "ForeColor", m_bColors.tForeColor, TranslateColor(vbButtonText)
        .WriteProperty "CheckBoxMode", m_bCheckBoxMode, False
        .WriteProperty "Value", m_bValue, False
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
        .WriteProperty "HandPointer", m_bHandPointer, False
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "Picture", m_Picture, Nothing
        .WriteProperty "PictureHover", m_PicOver, Nothing
        .WriteProperty "PictureAlign", m_PictureAlign, epLeftOfCaption
        .WriteProperty "pictureSize", m_PicSize, epsNormal
        .WriteProperty "UseMaskCOlor", m_bUseMaskColor, False
        .WriteProperty "MaskColor", m_lMaskColor, &HE0E0E0
        .WriteProperty "CaptionAlign", m_CaptionAlign, ecCenterAlign
    End With

End Sub

'Determine if the passed function is supported

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim hMod        As Long
Dim bLibLoaded  As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        FreeLibrary hMod
    End If

End Function

'Track the mouse leaving the indicated window

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

Dim tme As TRACKMOUSEEVENT_STRUCT

    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If

End Sub

'=========================================================================
'PUBLIC ROUTINES including subclassing & public button properties

' CREDITS: Paul Caton
'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'Parameters:
'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'hWnd     - The window handle
'uMsg     - The message number
'wParam   - Message related data
'lParam   - Message related data
'Notes:
'If you really know what you're doing, it's possible to change the values of the
'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'values get passed to the default handler.. and optionaly, the 'after' callback

Static bMoving As Boolean

    Select Case uMsg
    Case WM_MOUSEMOVE
        If Not m_bMouseInCtl Then
            m_bMouseInCtl = True
            TrackMouseLeave lng_hWnd
            If m_bMouseInCtl Then
                'If Not m_bIsSpaceBarDown Then m_Buttonstate = eStateOver
                If Not m_bIsSpaceBarDown Then
                    m_Buttonstate = eStateOver
                End If
            End If
            RedrawButton
            'CreateToopTip
            RaiseEvent MouseEnter
        End If

    Case WM_MOUSELEAVE

        m_bMouseInCtl = False
        If m_bIsSpaceBarDown Then Exit Sub
        If m_bEnabled Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        'DestroyToolTip
        RaiseEvent MouseLeave
    
    Case WM_NCACTIVATE, WM_ACTIVATE
        If wParam Then
            m_bParentActive = True
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            If m_bDefault Then
                RedrawButton
            End If
            RedrawButton
        Else
            m_bIsDown = False
            m_bIsSpaceBarDown = False
            m_bHasFocus = False
            m_bParentActive = False
            If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
            RedrawButton
        End If
    End Select

End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552

    MsgBox "JCButton" & vbNewLine & _
           "A Multistyle Button Control" & vbNewLine & vbNewLine & _
           "Created by: Juned S. Chhipa", vbInformation + vbOKOnly, "About"

End Sub

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_bColors.tBackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    m_bColors.tBackColor = New_BackColor
    RedrawButton
    PropertyChanged "BackColor"

End Property

Public Property Get ButtonStyle() As enumButtonStlyes

    ButtonStyle = m_ButtonStyle

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As enumButtonStlyes)

    m_ButtonStyle = New_ButtonStyle
    SetThemeColors          'Set colors
    CreateRegion            'Create Region Again
    RedrawButton            'Obviously, force redraw!!!
    PropertyChanged "ButtonStyle"

End Property

Public Property Get Caption() As String

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    SetAccessKey
    RedrawButton
    PropertyChanged "Caption"

End Property

Public Property Get CaptionAlign() As enumCaptionAlign

    CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As enumCaptionAlign)

    m_CaptionAlign = New_CaptionAlign
    RedrawButton
    PropertyChanged "CaptionAlign"

End Property

Public Property Get CheckBoxMode() As Boolean

    CheckBoxMode = m_bCheckBoxMode

End Property

Public Property Let CheckBoxMode(ByVal New_CheckBoxMode As Boolean)

    m_bCheckBoxMode = New_CheckBoxMode
    'If Not m_bCheckBoxMode Then m_Buttonstate = eStateNormal
    If Not m_bCheckBoxMode Then
        m_Buttonstate = eStateNormal
    End If
    RedrawButton
    PropertyChanged "Value"
    PropertyChanged "CheckBoxMode"

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_bEnabled
    'UserControl.Enabled = m_enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_bEnabled = New_Enabled
    UserControl.Enabled = m_bEnabled
    RedrawButton
    PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont

    Set Font = mFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)

    Set mFont = New_Font
    Refresh
    RedrawButton
    PropertyChanged "Font"
    Call mFont_FontChanged("")

End Property

Private Sub mFont_FontChanged(ByVal PropertyName As String)

    Set UserControl.Font = mFont
    Refresh
    RedrawButton
    PropertyChanged "Font"
    
End Sub

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_bColors.tForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

    m_bColors.tForeColor = New_ForeColor
    UserControl.ForeColor = m_bColors.tForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"

End Property

Public Property Get HandPointer() As Boolean
    
    HandPointer = m_bHandPointer
    
End Property

Public Property Let HandPointer(ByVal New_HandPointer As Boolean)
    
    m_bHandPointer = New_HandPointer
    PropertyChanged "HandPointer"
    
End Property

Public Property Get hWnd() As Long

    ' --Handle that uniquely identifies the control
    hWnd = UserControl.hWnd

End Property

Public Property Get MaskColor() As OLE_COLOR

    MaskColor = m_lMaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)

    m_lMaskColor = New_MaskColor
    RedrawButton
    PropertyChanged "MaskColor"

End Property

Public Property Get MouseIcon() As IPictureDisp

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)

    On Error Resume Next
        Set UserControl.MouseIcon = New_Icon
        If (New_Icon Is Nothing) Then
            UserControl.MousePointer = 0 ' vbDefault
        Else
            UserControl.MousePointer = 99 ' vbCustom
        End If
        PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_Cursor As MousePointerConstants)
    
    UserControl.MousePointer = New_Cursor
    PropertyChanged "MousePointer"

End Property

Public Property Get Picture() As StdPicture

    Set Picture = m_Picture

End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)

    Set m_Picture = New_Picture
    If Not New_Picture Is Nothing Then
        RedrawButton
        PropertyChanged "Picture"
    Else
        UserControl_Resize
    End If

End Property
Public Function SetPicture(ByVal New_Picture As Picture)

    Set m_Picture = New_Picture
    If Not New_Picture Is Nothing Then
        RedrawButton
        PropertyChanged "Picture"
    Else
        UserControl_Resize
    End If

End Function
Public Property Get PictureAlign() As enumPictureAlign

    PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As enumPictureAlign)

    m_PictureAlign = New_PictureAlign
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "PictureAlign"

End Property

Public Property Get PictureSize() As enumPictureSize
    
    PictureSize = m_PicSize
    
End Property

Public Property Let PictureSize(ByVal New_Size As enumPictureSize)

Dim tmpPic As New StdPicture
Set tmpPic = m_Picture

    m_PicSize = New_Size
    
    RedrawButton
    PropertyChanged "PictureSize"
    
End Property

Public Property Get PictureHover() As StdPicture

    Set PictureHover = m_PicOver

End Property

Public Property Set PictureHover(ByVal New_PictureHover As StdPicture)
    
    If m_Picture Is Nothing Then
        Set m_Picture = New_PictureHover        'Normal picture essential
        PropertyChanged "Picture"
    End If
    
    Set m_PicOver = New_PictureHover
    If Not New_PictureHover Is Nothing Then
        RedrawButton
        PropertyChanged "PictureHover"
    Else
        UserControl_Resize
    End If

End Property

Public Property Get ShowFocusRect() As Boolean

    ShowFocusRect = m_bShowFocus

End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)

    m_bShowFocus = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"

End Property

Public Property Get UseMaskColor() As Boolean

    UseMaskColor = m_bUseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)

    m_bUseMaskColor = New_UseMaskColor
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "UseMaskColor"

End Property

Public Property Get Value() As Boolean

    Value = m_bValue

End Property

Public Property Let Value(ByVal New_Value As Boolean)

    If m_bCheckBoxMode Then
        m_bValue = New_Value
        'If Not m_bValue Then m_Buttonstate = eStateNormal
        If Not m_bValue Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        PropertyChanged "Value"
    Else
        m_Buttonstate = eStateNormal
        RedrawButton
    End If

End Property

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)

'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler

    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zAddMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zAddMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With

End Sub

'Delete a message from the table of those that will invoke a callback.

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)

'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to be removed from the before, after or both callback tables

    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            zDelMsg uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And eMsgWhen.MSG_AFTER Then
            zDelMsg uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With

End Sub

'Return whether we're running in the IDE.

Private Function Subclass_InIDE() As Boolean

    Debug.Assert zSetTrue(Subclass_InIDE)

End Function

'Start subclassing the passed window handle

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long

'Parameters:
'lng_hWnd  - The handle of the window to be subclassed
'Returns;
'The sc_aSubData() index

Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
Const PATCH_0A              As Long = 186                                             'Address of the owner object
Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
Static pCWP                 As Long                                                   'Address of the CallWindowsProc
Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
Dim i                       As Long                                                   'Loop index
Dim j                       As Long                                                   'Loop index
Dim nSubIdx                 As Long                                                   'Subclass data index
Dim sHex                    As String                                                 'Hex code string

'If it's the first time through here..

    If aBuf(1) = 0 Then

        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                                'Next pair of hex characters

        'Get API function addresses
        If Subclass_InIDE Then                                                              'If we're running in the VB IDE
            aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                               'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                                    'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
        RtlMoveMemory ByVal .nAddrSub, aBuf(1), CODE_LEN                               'Copy the machine code from the static byte array to the code array in sc_aSubData
        zPatchRel .nAddrSub, PATCH_01, pEbMode                                         'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        zPatchVal .nAddrSub, PATCH_02, .nAddrOrig                                      'Original WndProc address for CallWindowProc, call the original WndProc
        zPatchRel .nAddrSub, PATCH_03, pSWL                                            'Patch the relative address of the SetWindowLongA api function
        zPatchVal .nAddrSub, PATCH_06, .nAddrOrig                                      'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        zPatchRel .nAddrSub, PATCH_07, pCWP                                            'Patch the relative address of the CallWindowProc api function
        zPatchVal .nAddrSub, PATCH_0A, ObjPtr(Me)                                      'Patch the address of this object instance into the static machine code buffer
    End With

End Function

'Stop all subclassing

Private Sub Subclass_StopAll()

Dim i As Long

    i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
    Do While i >= 0                                                                       'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
                Subclass_Stop .hWnd                                                        'Subclass_Stop
            End If
        End With

        i = i - 1                                                                           'Next element
    Loop

End Sub

'Stop subclassing the passed window handle

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

'Parameters:
'lng_hWnd  - The handle of the window to stop being subclassed

    With sc_aSubData(zIdx(lng_hWnd))
        SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig                                  'Restore the original WndProc
        zPatchVal .nAddrSub, PATCH_05, 0                                               'Patch the Table B entry count to ensure no further 'before' callbacks
        zPatchVal .nAddrSub, PATCH_09, 0                                               'Patch the Table A entry count to ensure no further 'after' callbacks
        GlobalFree .nAddrSub                                                           'Release the machine code memory
        .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                                       'Clear the before table
        .nMsgCntA = 0                                                                       'Clear the after table
        Erase .aMsgTblB                                                                     'Erase the before table
        Erase .aMsgTblA                                                                     'Erase the after table
    End With

End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)

Dim nEntry  As Long                                                                   'Message table entry index
Dim nOff1   As Long                                                                   'Machine code buffer offset 1
Dim nOff2   As Long                                                                   'Machine code buffer offset 2

    If uMsg = ALL_MESSAGES Then                                                           'If all messages
        nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
    Else                                                                                  'Else a specific message number
        Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
                Exit Sub                                                                        'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
                Exit Sub                                                                        'Bail
            End If
        Loop                                                                                'Next entry

        nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
        nOff1 = PATCH_04                                                                    'Offset to the Before table
        nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
    Else                                                                                  'Else after
        nOff1 = PATCH_08                                                                    'Offset to the After table
        nOff2 = PATCH_09                                                                    'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        zPatchVal nAddr, nOff1, VarPtr(aMsgTbl(1))                                     'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    zPatchVal nAddr, nOff2, nMsgCnt                                                  'Patch the appropriate table entry count

End Sub

'Return the memory address of the passed function in the passed dll

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first

End Function

'Worker sub for Subclass_DelMsg

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)

Dim nEntry As Long

    If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
        nMsgCnt = 0                                                                         'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
            nEntry = PATCH_05                                                                 'Patch the before table message count location
        Else                                                                                'Else after
            nEntry = PATCH_09                                                                 'Patch the after table message count location
        End If
        zPatchVal nAddr, nEntry, 0                                                     'Patch the table message count to zero
    Else                                                                                  'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                           'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
                Exit Do                                                                         'Bail
            End If
        Loop                                                                                'Next entry
    End If

End Sub

'Get the sc_aSubData() array index of the passed hWnd

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                                'If we're searching not adding
                    Exit Function                                                                 'Found
                End If
            ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
                If bAdd Then                                                                    'If we're adding
                    Exit Function                                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                                     'Decrement the index
    Loop

    If Not bAdd Then
        Debug.Assert False                                                                  'hWnd not found, programmer error
    End If

End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

    RtlMoveMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4

End Sub

'Patch the machine code buffer at the indicated offset with the passed value

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

    RtlMoveMemory ByVal nAddr + nOffset, nValue, 4

End Sub

'Worker function for Subclass_InIDE

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean

    zSetTrue = True
    bValue = True

End Function

'End of Subclassing routines

'---------------x---------------x--------------x--------------x-----------x---
' Oops! Control resulted Longer than expected!
' Lots of hours and lots of tedious work!   This is my first submission on PSC
' So if you want to vote for this, just do it ;)
' Comments are greatly appreciated...
'---------------x---------------x--------------x--------------x-----------x---
