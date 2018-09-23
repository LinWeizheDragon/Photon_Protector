Attribute VB_Name = "tray"
Option Explicit
Type NOTIFYICONDATA

     cbSize As Long

     hWnd As Long

     uID As Long

     uFlags As Long

     uCallbackMessage As Long

     hIcon As Long

     szTip As String * 64

    End Type

    Public a As Long

    '以下为 Shell_NotifyIcon将用到的常量

    Public Const NIF_ICON = &H2

    Public Const NIF_MESSAGE = &H1

    Public Const NIF_TIP = &H4

    Public Const NIM_ADD = &H0

    Public Const NIM_DELETE = &H2

    Public Const NIM_MODIFY = &H1

    'Shell_NotifyIcon的函数声明

    Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

    '处理消息将用到的结构、常量、API声明

    Type POINTAPI

     x As Long

     y As Long

    End Type

    Type Msg

     hWnd As Long

     message As Long

     wParam As Long

     lParam As Long

     time As Long

     pt As POINTAPI

    End Type

    Public Const WM_USER = &H400

    Public Const WM_RBUTTONDOWN = &H204

    Public Const WM_LBUTTONDOWN = &H201

    Public Const GWL_WNDPROC = -4

    Public trayflag As Boolean

    Global lpPrevWndProc As Long

    Global gHW As Long

    Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Global Const HWND_TOPMOST = -1
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


  '以下过程为消息循环处理

    Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If hw = frmMain.hWnd And uMsg = WM_USER + 100 Then '检测到鼠标点动托盘图标

     Select Case lParam

     Case WM_RBUTTONDOWN '鼠标右键按下
        'frmMain.WindowState = 1
        frmMain.Visible = False
     'Form1.PopupMenu Form1.traymnu '弹出菜单

     Case WM_LBUTTONDOWN '鼠标左键按下

                SetWindowPos hw, -1, 0, 0, 0, 0, &H1 Or &H2
                SetWindowPos hw, -2, 0, 0, 0, 0, &H2 Or &H1
                frmMain.Visible = True
                frmMain.WindowState = 0
     'Form1.PopupMenu Form1.mnutray2 '弹出菜单

     Case Else

     End Select

    Else '调用缺省窗口指针

     WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

    End If

    End Function



