Attribute VB_Name = "mdValue"
Public IniPath As String

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long



Public Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type

