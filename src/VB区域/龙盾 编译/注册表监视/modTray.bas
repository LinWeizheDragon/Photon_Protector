

Attribute VB_Name = "modTray"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''
'²Ù×÷ÍÐÅÌÄ£¿é
'''''''''''''''''''''''''''''''''''''''''
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const WM_MOUSEMOVE = &H200


Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private trayStructure As NOTIFYICONDATA
'Private IconObject As Object

Private Function AddIcon(ByVal obj As Object, ByVal IconID As Long, ByVal Icon As Object, ByVal ToolTip As String) 'Ôö¼ÓÍÐÅÌ
    trayStructure.cbSize = Len(trayStructure)
    trayStructure.hwnd = obj.hwnd
    trayStructure.uID = IconID
    trayStructure.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    trayStructure.uCallbackMessage = WM_TRAYICON
    trayStructure.hIcon = Icon
    trayStructure.szTip = ToolTip & Chr$(0)
    '½¨Á¢ÍÐÅÌ
    Call Shell_NotifyIcon(NIM_ADD, trayStructure)
End Function

Public Function DeleteSysTray() 'É¾³ýÍÐÅÌ
'    If IconObject Is Nothing Then Exit Function
    trayStructure.uID = frmRegMonitor.Icon.Handle
    Call Shell_NotifyIcon(NIM_DELETE, trayStructure)
End Function

Public Function SendToTray()
    AddIcon frmRegMonitor, frmRegMonitor.Icon.Handle, frmRegMonitor.Icon, "×¢²á±í¼à¿Ø" & vbNullChar
End Function
