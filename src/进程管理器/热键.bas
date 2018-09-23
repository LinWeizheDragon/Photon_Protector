Attribute VB_Name = "热键"
Option Explicit
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
'Public Const GWL_WNDPROC = (-4)

Public preWinProc As Long
Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    ll As Long
End Type

Private Type t2Int
    lWord As Integer
    hWord As Integer
End Type

Public Function Wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then
        If wParam = idHotKey Then
            Dim lp As taLong, i2 As t2Int
            lp.ll = lParam
            LSet i2 = lp
            If (i2.lWord = Modifiers) And i2.hWord = uVirtKey Then
                'Shell "Notepad", vbNormalFocus
                SetWindowPos hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
                SetWindowPos hWnd, -2, 0, 0, 0, 0, &H2 Or &H1
                frmMain.Visible = True
                frmMain.WindowState = 0

            End If
        End If
    End If
    '如果不是热键信息则调用原来的程序
    Wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
    
End Function

