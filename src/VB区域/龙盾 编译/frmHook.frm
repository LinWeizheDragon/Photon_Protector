VERSION 5.00
Begin VB.Form frmHook 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "frmHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StartHook Lib "Hook.dll" () As Boolean
Private Declare Function UnLoadHook Lib "Hook.dll" () As Boolean


Private Sub Form_Load()
If StartHook Then
AddInfo "Hook OK!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If UnLoadHook Then
AddInfo "Hook Unload."
End If
Unload Me
End Sub
