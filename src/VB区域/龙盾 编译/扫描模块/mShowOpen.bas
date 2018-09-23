Attribute VB_Name = "mShowOpen"
Option Explicit
  '对话框
  Public Declare Function GetOpenFileName _
                  Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
                  (pOpenfilename As OPENFILENAME) As Long
  Public Declare Function GetSaveFileName _
                  Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
                  (pOpenfilename As OPENFILENAME) As Long
  Public Declare Function SHBrowseForFolder _
                  Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                  (lpBrowseInfo As BROWSEINFO) As Long
  Public Declare Function SHGetPathFromIDList _
                  Lib "shell32.dll" _
                  (ByVal pidl As Long, _
                  pszPath As String) As Long
  Public Declare Function CHOOSECOLOR _
                  Lib "comdlg32.dll" Alias "ChooseColorA" _
                  (pChoosecolor As CHOOSECOLOR) As Long
    
  Public Type OPENFILENAME
          lStructSize   As Long
          hwndOwner   As Long
          hInstance   As Long
          lpstrFilter   As String
          lpstrCustomFilter   As String
          nMaxCustFilter   As Long
          nFilterIndex   As Long
          lpstrFile   As String
          nMaxFile   As Long
          lpstrFileTitle   As String
          nMaxFileTitle   As Long
          lpstrInitialDir   As String
          lpstrTitle   As String
          flags   As Long
          nFileOffset   As Integer
          nFileExtension   As Integer
          lpstrDefExt   As String
          lCustData   As Long
          lpfnHook   As Long
          lpTemplateName   As String
  End Type
    
  Public Type BROWSEINFO
          hOwner   As Long
          pidlRoot   As Long
          pszDisplayName   As String
          lpszTitle   As String
          ulFlage   As Long
          lpfn   As Long
          lparam   As Long
          iImage   As Long
  End Type
    
  Private Type CHOOSECOLOR
          lStructSize   As Long
          hwndOwner   As Long
          hInstance   As Long
          rgbResult   As Long
          lpCustColors   As String
          flags   As Long
          lCustData   As Long
          lpfnHook   As Long
          lpTemplateName   As String
  End Type
    
  Public Const OFN_HIDEREADONLY = &H4           '隐藏只读打开
  Public Const OFN_READONLY = &H1           '只读打开为选中
  Public Const OFN_OVERWRITEPROMPT = &H2           '覆盖时提示
  Public Const OFN_ALLOWMULTISELECT = &H200           '多个选中
  Public Const OFN_EXPLORER = &H80000           '资源管理器
    
  '显示打开
Public Function ShowOpen(MehWnd As Long, FileOpen As String, Optional Title As String = "打开：", Optional Filter As String = vbNullChar + vbNullChar, Optional FilterIndex As Long = 0, Optional StartDir As String = vbNullChar, Optional flags As Long = OFN_HIDEREADONLY) As Long
  Dim OpenFN     As OPENFILENAME
  Dim Rc     As Long
    
  With OpenFN
    .hwndOwner = MehWnd
    .hInstance = App.hInstance
    .lpstrTitle = Title
    .lpstrFilter = Filter
    .nFilterIndex = FilterIndex
    .lpstrInitialDir = StartDir
    .lpstrFile = String$(256, 0)
    .nMaxFile = 255
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = 255
    .flags = flags
    .lStructSize = Len(OpenFN)
  End With
  Rc = GetOpenFileName(OpenFN)
  If Rc Then
    FileOpen = Left$(OpenFN.lpstrFile, OpenFN.nMaxFile)
    ShowOpen = True
  Else
    ShowOpen = False
  End If
End Function
    
'显示保存
Public Function ShowSave(MehWnd As Long, FileSave As String, Optional Title As String = "保存：", Optional Filter As String = vbNullChar + vbNullChar, Optional FilterIndex As Long = 0, Optional StartDir As String = vbNullChar, Optional flags As Long = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT) As Long
  Dim SaveFN As OPENFILENAME
  Dim Rc As Long
    
  With SaveFN
    .hwndOwner = MehWnd
    .hInstance = App.hInstance
    .lpstrTitle = Title
    .lpstrFilter = Filter
    .nFilterIndex = FilterIndex
    .lpstrInitialDir = StartDir
    .lpstrFile = FileSave + String$(255, Chr$(0))
    .nMaxFile = Len(.lpstrFile)
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = 255
    .flags = flags
    .lStructSize = Len(SaveFN)
  End With
  Rc = GetSaveFileName(SaveFN)
  If Rc Then
     FileSave = Left$(SaveFN.lpstrFile, SaveFN.nMaxFile)
     ShowSave = True
  Else
    ShowSave = False
  End If
End Function
    
'显示目录
Public Function ShowDir(MehWnd As Long, DirPath As String, Optional Title As String = "请选择文件夹：", Optional flage As Long = &H1, Optional DirID As Long) As Long
  Dim BI As BROWSEINFO
  Dim TempID As Long
  Dim TempStr As String
    
  TempStr = String$(255, Chr$(0))
  With BI
    .hOwner = MehWnd
    .pidlRoot = 0
    .lpszTitle = Title + Chr$(0)
    .ulFlage = flage
  End With
    
  TempID = SHBrowseForFolder(BI)
  DirID = TempID
    
  If SHGetPathFromIDList(ByVal TempID, ByVal TempStr) Then
    DirPath = Left$(TempStr, InStr(TempStr, Chr$(0)) - 1)
    ShowDir = -1
  Else
    ShowDir = 0
  End If
End Function
  

