Attribute VB_Name = "mSetButtonFlat"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)

Public Function SetButtonFlat(ByVal hWnd As Long) As Boolean
'设置按钮样式的，不知道为什么很多人都喜欢这样的
Dim style As Long
style = GetWindowLong(hWnd, GWL_STYLE)
style = style Or BS_FLAT
SetButtonFlat = SetWindowLong(hWnd, GWL_STYLE, style)
End Function
