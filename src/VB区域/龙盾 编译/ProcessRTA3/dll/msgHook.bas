Attribute VB_Name = "msgHook"
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'Download by http://www.codefans.net
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HC_ACTION = 0

Public Const WH_GETMESSAGE = 3

Public DLLhandle As Long

Function EnableHook() As Long
        GetData ShareMem
        ShareMem.hHookID = SetWindowsHookEx(WH_GETMESSAGE, AddressOf HookProc, DLLhandle, 0)
        SetData ShareMem
End Function

Function FreeHook()

        GetData ShareMem
        Call UnhookWindowsHookEx(hHook)
        ShareMem.hHookID = 0
        SetData ShareMem

End Function

Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        'GetData ShareMem
        HookProc = CallNextHookEx(ShareMem.hHookID, nCode, wParam, lParam)
End Function

