Attribute VB_Name = "mdOpenFile"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShowStyle
    vbHide
    vbMaximizedFocus
    vbMinimizedFocus
    vbMinimizedNoFocus
    vbNormalFocus
    vbNormalNoFocus
End Enum

Public Function OpenFile(ByVal OpenName As String, Optional ByVal InitDir As String = vbNullString, Optional ByVal msgStyle As ShowStyle = vbNormalFocus)
    ShellExecute 0&, vbNullString, OpenName, vbNullString, InitDir, msgStyle
End Function

