Attribute VB_Name = "ModBFF"
'This is a hacked version of Bobo's Browse for folder.  Don't use this.
'Go to PSC and get the original!

'***************BOBO  ENTERPRISES  2001**********************
'Please report any bugs through PSC or to gtkerr@bigpond.com
'(Subject: Browse for Folders BUG)
'
'Still to be implemented features:
'       Context help
'       Popup menu from Treeview
'       New folder update without restarting BFF
'Credit to "Mr. BoBo"
Option Explicit
'**************Win 2K compliant FileExists********************
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'**************CHECK OS*****************************
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'*********************General Declares**************************
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
'constants required
Private Const GWL_WNDPROC = (-4)                'Used in setting hooks
Private Const GW_NEXT = 2                       'used to enumerate child windows
Private Const GW_CHILD = 5                      'used to enumerate child windows

Private Const WM_GETMINMAXINFO As Long = &H24&  'scrollbar settings
Private Const WM_LBUTTONUP = &H202              'used in hooks
Private Const WM_LBUTTONDOWN = &H201            'used in hooks
Private Const WM_CHAR = &H102                   'used in hooks
Private Const WM_SIZE = &H5                     'used in hooks
Private Const WM_GETFONT = &H31                 'used to get the current font
Private Const WM_SETFONT = &H30                 'used to set the font in any new windows
Private Const WM_EXITSIZEMOVE = &H232           'used in hooks
Private Const WM_GETTEXT = &HD                  'used to read textboxes
Private Const WM_GETTEXTLENGTH = &HE            'used to read textboxes
Private Const WM_HELP = &H53                    'used in hooks
Private Const WM_SETTEXT = &HC                  'used to update textboxes

Private Const WS_CHILD = &H40000000             'style setting
Private Const WS_EX_CLIENTEDGE = &H200&         'style setting
Private Const WS_EX_RIGHTSCROLLBAR = &H0&       'style setting
Private Const WS_DISABLED = &H8000000           'style setting
Private Const WS_EX_STATICEDGE = &H20000        'style setting

Private Const BM_GETCHECK = &HF0                'checking the state of the checkbox
Private Const BM_SETCHECK = &HF1                'checking the state of the checkbox
Private Const BM_CLICK = &HF5                   'simulate a button click

Private Const BS_CHECKBOX = &H2&                'style setting

Private Const EM_SETSEL = &HB1                  'used to update textboxes

Private Const ES_AUTOHSCROLL = &H80&            'style setting
Private Const ES_WANTRETURN = &H1000&           'style setting
Private Const ES_MULTILINE = &H4&               'style setting

Private Const SBS_SIZEGRIP = &H10&              'style setting
Private Const SBS_SIZEBOX = &H8&                'style setting

Private Const RDW_INVALIDATE = &H1              'redraw command



'Used to set the minimum scroll size
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
'used to create new buttons,labels,checkboxes etc.
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    Style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type
'used to locate window positions
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim R As RECT
'********************Browse for Folders*****************************
Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BFFM_ENABLEOK = &H465
Private Const BFFM_SETSELECTION = &H466
Private Const BFFM_SETSTATUSTEXT = &H464
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILED = 3
Public Const BIF_USENEWUI = &H40               '(SHELL32.DLL Version 5.0). Use the new user interface, including an edit box.

'****************Browse Load Variables********************
Public Type BoboBrowse
    Titlebar As String           'Browse for Folder window caption
    Prompt As String             'Descriptive text
    InitDir As String            'Start browsing from this folder
    CHCaption As String          'Checkbox caption
    OKCaption As String          'Browse for Folder OK button caption
    CancelCaption As String      'Browse for Folder Cancel button caption
    NewFCaption As String        'New folder button caption
    RootDir As Long              'Special folder to browse from
    AllowResize As Boolean       'Use the resize ability
    CenterDlg As Boolean         'Center the Browse for Folder window
    DoubleSizeDlg As Boolean     'Make the Browse for Folder window large (Not Double)
    FSDlg As Boolean             'Make the Browse for Folder window full screen
    ShowButton As Boolean        'Show the New folder button
    ShowCheck As Boolean         'Show the checkbox
    EditBoxOld As Boolean        'Use the default Browse for Folder Edit window
    EditBoxNew As Boolean        'Use Win2K style Browse for Folder Edit window
    StatusText As Boolean        'Show Browse for Folder Status text
    ShowFiles As Boolean         'Include files
    CHvalue As Integer           'Value returned by the checkbox
    OwnerForm As Long            'Handle to the calling form - if invalid Desktop window is used
End Type
Public BB As BoboBrowse

'*****************Browsing Variables**************
Dim DialogWindow As Long            'Browse for Folder window
Dim SysTreeWindow As Long           'Browse for Folder Treeview window
Dim OKbuttonWindow As Long          'Browse for Folder OK button window
Dim CancelbuttonWindow As Long      'Browse for Folder Cancel button window
Dim ScrollWindow As Long            'The scroll control to resize
Dim dummyWindow As Long             'Workaround Sizegrip for Win 95/98
Dim ButtonWindow As Long            'Either New folder button or checkbox
Dim StattxtWindow As Long           'Browse for Folder Status text window
Dim EditWindowOld As Long           'Browse for Folder Edit window
Dim EditWindow As Long              'New style edit window
Dim LabelWindow As Long             'Label for new style edit window
Dim EditTop As Long                 'Top of Browse for Folder Edit window
Dim EditHeight As Long              'Height of Browse for Folder Edit window
Dim StattxtTop As Long              'Top of Browse for Folder Status text window
Dim StattxtHeight As Long           'Height of Browse for Folder Status text window
Dim TreeTop As Long                 'Top of Browse for Treeview window
Dim CurrentDir As String            'Currently selected folder
Dim Newboy As Boolean               'User created a new folder
Dim RoomForSizer As Long            'Allow space for the scroll window
Private glPrevWndProc As Long       'Window hook for New Folder button
Private glPrevWndProcDlg As Long    'Window hook for Browse for Folder window
Private glPrevWndProcEdit As Long   'Window hook for new style edit window
Private glPrevWndProcFS As Long     'Window hook for Size grip (needed in Win2K)

Public Function BrowseFF() As String
'Call this function from your form

    'Example Calls :
    
    'Private Sub Command1_Click()
    '    BB.AllowResize = True
    '    BB.DoubleSizeDlg = True
    '    BB.OKCaption = "Open"
    '    BB.ShowFiles = True
    '    Label1 = BrowseFF
    'End Sub
    
    'or just:
    'Private Sub Command1_Click()
    '    Label1 = BrowseFF
    'End Sub
    
    Dim hFont As Long
    Dim IDList As Long
    Dim mTemp As String
    Dim mFlags As Long
    Dim tBrowseInfo As BrowseInfo
    BB.CHvalue = 0
startagain: 'If a new folder was created we need to come back here
    If IsWindow(BB.OwnerForm) = 0 Then BB.OwnerForm = GetDesktopWindow
    If Len(BB.Prompt) = 0 Then BB.Prompt = "Select a folder"
    mFlags = BIF_VALIDATE
    If BB.EditBoxOld Then mFlags = mFlags + BIF_EDITBOX
    If BB.StatusText Then mFlags = mFlags + BIF_STATUSTEXT
    If BB.ShowFiles Then mFlags = mFlags + BIF_BROWSEINCLUDEFILES
    With tBrowseInfo
      .hWndOwner = BB.OwnerForm
      .lpszTitle = lstrcat(BB.Prompt, "")
      .pIDLRoot = BB.RootDir
      .ulFlags = mFlags
      .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    IDList = SHBrowseForFolder(tBrowseInfo)
    If (IDList) Then
      mTemp = Space(MAX_PATH)
      SHGetPathFromIDList IDList, mTemp
      mTemp = Left(mTemp, InStr(mTemp, vbNullChar) - 1)
      BrowseFF = mTemp
        If Newboy = True Then GoTo startagain
        CleanUp
    Else
      BrowseFF = ""
        If Newboy = True Then GoTo startagain
        CleanUp
    End If

End Function
'Used to allow BrowseCallbackProc hook
Private Function GetAddressofFunction(Add As Long) As Long
  GetAddressofFunction = Add
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
'Messages to the Browse for Folder window are recieved here
Dim lpIDList As Long
Dim ret As Long, temp As String, TVr As RECT, CS As CREATESTRUCT
Dim sBuffer As String, hFont As Long
Dim hWnda As Long, ClWind As String * 14, ClCaption As String * 100
On Error Resume Next
DialogWindow = hWnd
If Len(BB.Titlebar) = 0 Then BB.Titlebar = "Browse for Folder"
SetWindowText DialogWindow, BB.Titlebar
If BB.AllowResize Then RoomForSizer = 50
Select Case uMsg
  Case BFFM_INITIALIZED 'Lets set things up the way we want
    If BB.InitDir = "" Then BB.InitDir = "c:\"
    Call SendMessage(hWnd, BFFM_SETSELECTION, 1, BB.InitDir) 'Start here please
    CurrentDir = BB.InitDir
    If Newboy = False Then 'If we are not just updating with a new folder
         'locate the window and then set its size
        Call GetWindowRect(DialogWindow, R)
        If BB.DoubleSizeDlg Then
            Call MoveWindow(DialogWindow, R.Left, R.Top, 480, 480, True)
        ElseIf BB.FSDlg Then
            Call MoveWindow(DialogWindow, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, True)
        Else
            Call MoveWindow(DialogWindow, R.Left, R.Top, 320, 320, True)
        End If
        Call GetWindowRect(DialogWindow, R)
        'Put the window where we want it
        If BB.CenterDlg Then
            Call MoveWindow(DialogWindow, (Screen.Width / Screen.TwipsPerPixelX) / 2 - (R.Right - R.Left) / 2, (Screen.Height / Screen.TwipsPerPixelY) / 2 - (R.Bottom - R.Top) / 2, R.Right - R.Left, R.Bottom - R.Top, True)
        End If
        Call GetWindowRect(DialogWindow, R) 'Remember the new position
    Else
        'If we are updating with a new folder use the old size and position
        Call MoveWindow(DialogWindow, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, True)
    End If
    Newboy = False 'reset this flag
    'Get a handle on the elements within the Browse for Folder window
    hWnda = GetWindow(hWnd, GW_CHILD) 'Get the child windows
        Do While hWnda <> 0 'Go through all the children and get its ClassName
            GetClassName hWnda, ClWind, 14
            If Left(ClWind, 6) = "Button" Then 'Found a button
                GetWindowText hWnda, ClCaption, 100
                If UCase(Left(ClCaption, 2)) = "OK" Then 'Its the OK button
                    OKbuttonWindow = hWnda 'Remember its handle
                    If Len(BB.OKCaption) = 0 Then BB.OKCaption = "OK" 'Default
                    SetWindowText OKbuttonWindow, BB.OKCaption 'Set its caption
                End If
                If UCase(Left(ClCaption, 6)) = "CANCEL" Then 'Its the Cancel button
                    CancelbuttonWindow = hWnda 'Remember its handle
                    If Len(BB.CancelCaption) = 0 Then BB.CancelCaption = "Cancel" 'Default
                    SetWindowText CancelbuttonWindow, BB.CancelCaption 'Set its caption
                End If
            End If
            If Left(ClWind, 13) = "SysTreeView32" Then 'Its the Treeview
                SysTreeWindow = hWnda 'Remember its handle
                Call GetWindowRect(SysTreeWindow, TVr)
                'Remember its Top position - used to locate other controls on the resize event
                TreeTop = TVr.Top - R.Top
            End If
            If BB.EditBoxOld Then
                If Left(ClWind, 4) = "Edit" Then 'Its the default Edit window because we haven't made one yet
                    EditWindowOld = hWnda 'Remember its handle
                    Call GetWindowRect(EditWindowOld, TVr)
                    'Remember its Top and height - used to locate other controls on the resize event
                    EditTop = TVr.Top - R.Top
                    EditHeight = TVr.Bottom - TVr.Top
                End If
            End If
            If Left(ClWind, 6) = "Static" Then 'label
                If UCase(Left(ClCaption, Len(BB.Prompt))) <> BB.Prompt Then
                    'If its not our descriptive text it must it must be the status text
                    StattxtWindow = hWnda 'Remember its handle
                    Call GetWindowRect(StattxtWindow, TVr)
                    'Remember its Top and height - used to locate other controls on the resize event
                    StattxtTop = TVr.Top - R.Top
                    StattxtHeight = TVr.Bottom - TVr.Top
                End If
            End If
            hWnda = GetWindow(hWnda, GW_NEXT)
        Loop
        If BB.RootDir <> 3 And BB.RootDir <> 4 And BB.RootDir <> 10 And BB.RootDir <> 18 And BB.RootDir <> 19 Then
        'if the RootDir is Control Panel,Printers,Recycle,Network or Nethood then no Buttons/Checkbox or Editbox
            If BB.EditBoxNew Then
                'Create a textbox and a label
                EditWindow = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "", WS_CHILD Or ES_MULTILINE Or ES_WANTRETURN Or ES_AUTOHSCROLL, 0, 0, 0, 23, DialogWindow, 0, App.hInstance, CS)
                LabelWindow = CreateWindowEx(0, "STATIC", "Folder :", WS_CHILD, 0, 0, 0, 23, DialogWindow, 0, App.hInstance, CS)
                'make the font the same as the OK button
                hFont = SendMessage(OKbuttonWindow, WM_GETFONT, 0&, ByVal 0&)
                SendMessage EditWindow, WM_SETFONT, hFont, ByVal 1&
                SendMessage LabelWindow, WM_SETFONT, hFont, ByVal 1&
                ShowWindow LabelWindow, 1
                ShowWindow EditWindow, 1
                'Place a window hook on the editbox so we can take action on any input
                glPrevWndProcEdit = fSubClassEdit()
            End If
            If BB.ShowButton Or BB.ShowCheck Then
                If BB.ShowButton Then
                    'make a standard button and set its caption
                    If Len(BB.NewFCaption) = 0 Then BB.NewFCaption = "New Folder" 'Default
                    ButtonWindow = CreateWindowEx(0, "BUTTON", BB.NewFCaption, WS_CHILD, 0, 0, 75, 23, DialogWindow, 0, App.hInstance, CS)
                Else
                    'make a standard checkbox
                    If Len(BB.CHCaption) = 0 Then BB.CHCaption = "Include subfolders" 'Default
                    ButtonWindow = CreateWindowEx(0, "BUTTON", BB.CHCaption, WS_CHILD Or BS_CHECKBOX, 20, 0, 110, 23, DialogWindow, 0, App.hInstance, CS)
                End If
                'Set its font to match the OK button
                hFont = SendMessage(OKbuttonWindow, WM_GETFONT, 0&, ByVal 0&)
                SendMessage ButtonWindow, WM_SETFONT, hFont, ByVal 1&
                ShowWindow ButtonWindow, 1
                'Place a window hook on the button or checkbox so we can take action on any input
                glPrevWndProc = fSubClass()
            End If
        Else
            'if the RootDir is Control Panel,Printers,Recycle,Network or Nethood then
            'make sure these are false or the resizing will go astray
            BB.ShowButton = False
            BB.ShowCheck = False
            BB.EditBoxNew = False
        End If
        If BB.AllowResize Then 'add some scrollbars with a size grip in the corner
            If Is2K Then
                ScrollWindow = CreateWindowEx(WS_EX_RIGHTSCROLLBAR, "SCROLLBAR", "", WS_CHILD Or SBS_SIZEGRIP Or SBS_SIZEBOX, R.Right - R.Left - 24, R.Bottom - R.Top - 44, 16, 16, DialogWindow, 0, App.hInstance, CS)
                ShowWindow ScrollWindow, 1 'show the scrollbox
                glPrevWndProcFS = fSubClassFS() 'we need to hook for Win2K- not sure why
            Else
                'I cant get the sizegrip to work under Win95/98
                'so I create a dummy scrollbar with a sizegrip
                'and disable it, setting it as the real scrollbars child
                ScrollWindow = CreateWindowEx(0, "SCROLLBAR", "", WS_CHILD Or SBS_SIZEBOX, R.Right - R.Left - 24, R.Bottom - R.Top - 44, 16, 16, DialogWindow, 0, App.hInstance, CS)
                dummyWindow = CreateWindowEx(0, "SCROLLBAR", "", WS_CHILD Or SBS_SIZEGRIP Or WS_DISABLED, -4, -4, 16, 16, ScrollWindow, 0, App.hInstance, CS)
                ShowWindow ScrollWindow, 1 'show the scrollbox
                ShowWindow dummyWindow, 1 'show the sizegrip
            End If
            'Place a window hook on the main window so when it resizes we can move all the controls appropriately
            glPrevWndProcDlg = fSubClassDlg()
        End If
        'Enter the start folder text in the edit box
        Call SendMessage(EditWindow, WM_SETTEXT, 0, FileOnly(CurrentDir))
        Call SendMessage(EditWindow, EM_SETSEL, Len(FileOnly(CurrentDir)), 0)
        'Done setting up, so call the resize event
        SizeAndPosition
  Case BFFM_SELCHANGED
        sBuffer = Space(MAX_PATH)
        ret = SHGetPathFromIDList(lp, sBuffer)
        'update the edit box and the status label with the new directory
        If ret = 1 Then
            Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            If Len(StripTerminator(sBuffer)) > 3 Then
                Call SendMessage(EditWindow, WM_SETTEXT, 0, FileOnly(sBuffer))
                Call SendMessage(EditWindow, EM_SETSEL, Len(FileOnly(sBuffer)), 0)
            Else
                Call SendMessage(EditWindow, WM_SETTEXT, 0, sBuffer)
                Call SendMessage(EditWindow, EM_SETSEL, Len(sBuffer), 0)
            End If
            CurrentDir = sBuffer
        End If
End Select
BrowseCallbackProc = 0
End Function
'****************Hook and Unhook our windows*****************************8
Private Function fSubClassDlg() As Long
fSubClassDlg = SetWindowLong(DialogWindow, GWL_WNDPROC, AddressOf pMyWindowProcDlg)
End Function
Private Sub pUnSubClassDlg()
Call SetWindowLong(DialogWindow, GWL_WNDPROC, glPrevWndProcDlg)
End Sub
Private Function fSubClassEdit() As Long
fSubClassEdit = SetWindowLong(EditWindow, GWL_WNDPROC, AddressOf pMyWindowProcEdit)
End Function
Public Sub pUnSubClassEdit()
Call SetWindowLong(EditWindow, GWL_WNDPROC, glPrevWndProcEdit)
End Sub
Private Function fSubClass() As Long
fSubClass = SetWindowLong(ButtonWindow, GWL_WNDPROC, AddressOf pMyWindowProc)
End Function
Private Sub pUnSubClass()
Call SetWindowLong(ButtonWindow, GWL_WNDPROC, glPrevWndProc)
End Sub
Private Function fSubClassFS() As Long
fSubClassFS = SetWindowLong(ScrollWindow, GWL_WNDPROC, AddressOf pMyWindowProcFS)
End Function
Public Sub pUnSubClassFS()
Call SetWindowLong(ScrollWindow, GWL_WNDPROC, glPrevWndProcFS)
End Sub
Private Function pMyWindowProcFS(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
pMyWindowProcFS = CallWindowProc(glPrevWndProcFS, hw, uMsg, wParam, lParam)
End Function
Private Function pMyWindowProcDlg(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'This is a hook on the main window
Dim iKeyCode As Long
Select Case uMsg
Case WM_GETMINMAXINFO 'stop the scroller if it gets too small
      Dim udtMINMAXINFO As MINMAXINFO
      Dim nWidthPixels&, nHeightPixels&
      nWidthPixels = Screen.Width \ Screen.TwipsPerPixelX
      nHeightPixels = Screen.Height \ Screen.TwipsPerPixelY
      CopyMemory udtMINMAXINFO, ByVal lParam, Len(udtMINMAXINFO)
      With udtMINMAXINFO
        .ptMinTrackSize.X = 320 'change to desired minimum size
        .ptMinTrackSize.Y = 320
      End With
      CopyMemory ByVal lParam, udtMINMAXINFO, Len(udtMINMAXINFO)
Case WM_SIZE
    Call GetWindowRect(DialogWindow, R) 'how big is it ?
    SizeAndPosition 'Move the controls to fit
Case WM_EXITSIZEMOVE
    Call GetWindowRect(DialogWindow, R) 'how big is it ?
    SizeAndPosition 'Move the controls to fit
End Select
pMyWindowProcDlg = CallWindowProc(glPrevWndProcDlg, hw, uMsg, wParam, lParam)
End Function
Private Function pMyWindowProcEdit(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'This is a hook on the edit box
Dim iKeyCode As Long, temp As String, tmpDir As String
Select Case uMsg
Case WM_CHAR
    iKeyCode = (wParam And &HFF)
    If iKeyCode = 13 Then 'user is finished - they pressed enter
        tmpDir = StripTerminator(CurrentDir)
        If Right(tmpDir, 1) = "\" Then tmpDir = Left(tmpDir, Len(tmpDir) - 1)
        temp = gettext(EditWindow) 'read the edit box
        If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
        If FileExists(PathOnly(tmpDir) + "\" + temp) Then 'does this work ?
            CurrentDir = PathOnly(tmpDir) + "\" + temp
            Call SendMessage(DialogWindow, BFFM_SETSELECTION, 1, CurrentDir)
        ElseIf FileExists(temp) Then 'well, try this instead
            CurrentDir = temp
            Call SendMessage(DialogWindow, BFFM_SETSELECTION, 1, CurrentDir)
        Else 'your a wacky user - what am I supposed to do ?
            MsgBox temp + vbCrLf + "The specified path does not exist." + vbCrLf + vbCrLf + "Check the path, and try again.", vbCritical, "Browse for Folder"
        End If
    Else
        pMyWindowProcEdit = CallWindowProc(glPrevWndProcEdit, hw, uMsg, wParam, lParam)
    End If
    Exit Function
End Select
pMyWindowProcEdit = CallWindowProc(glPrevWndProcEdit, hw, uMsg, wParam, lParam)
End Function
Private Function pMyWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'This is a hook on either the new folder button or checkbox
Select Case uMsg

Case WM_HELP 'comment this out if you think its too cheesy
    MsgBox "Creates a new folder in the currently selected directory", vbInformation, "Bobo Enterprises"
    Exit Function
'Case WM_LBUTTONUP
'    If BB.ShowButton Then GoNewboy ' new folder please
Case WM_LBUTTONDOWN
    If BB.ShowCheck Then 'check or uncheck the checkbox
        BB.CHvalue = SendMessage(ButtonWindow, BM_GETCHECK, 0&, ByVal 0&)
        If BB.CHvalue = 0 Then
            SendMessage ButtonWindow, BM_SETCHECK, 1, ByVal 1&
            BB.CHvalue = 1 'remember what it is
        Else
            SendMessage ButtonWindow, BM_SETCHECK, 0, ByVal 1&
            BB.CHvalue = 0 'remember what it is
        End If
    Putfocus DialogWindow
    End If
    Exit Function
End Select
pMyWindowProc = CallWindowProc(glPrevWndProc, hw, uMsg, wParam, lParam)
End Function

'Public Sub GoNewboy()
'Dim temp As String, tmpDir As String, blnBad As Boolean, strNewDir As String
'
'    tmpDir = StripTerminator(CurrentDir)
'    Do
'        temp = InputBox("Enter a name for your new folder" + vbCrLf + tmpDir + "\...", BB.Titlebar)
'        temp = StripIllegals(temp, "the folder name")
'        If Len(Trim(temp)) = 0 Then
'            Exit Sub
'        End If
'        strNewDir = getFullPath(tmpDir, temp)
'        If Dir(strNewDir, vbNormal + vbHidden + vbDirectory) <> "" Then
'            MsgBox "A Folder of that name already exists." + vbCrLf + "Enter a different name and try again.", vbCritical, BB.Titlebar
'            blnBad = True
'        End If
'    Loop Until Not blnBad
'    MkDir strNewDir
'    BB.InitDir = strNewDir  'set the new start folder
'    Newboy = True
'    'Clean up and close the window so we can re-open at the new folder
'    Call pUnSubClass
'    Call pUnSubClassDlg
'    Call pUnSubClassEdit
'    Call pUnSubClassFS
'    Call SendMessage(CancelbuttonWindow, BM_CLICK, 0, 0)
'    DestroyWindow LabelWindow
'    DestroyWindow EditWindow
'    DestroyWindow ButtonWindow
'    DestroyWindow ScrollWindow
'    DestroyWindow dummyWindow
'End Sub
Private Sub SizeAndPosition()
'Like a form_resize event except with API
Dim sysH As Long
sysH = 81
If BB.EditBoxNew Then sysH = 120
Call MoveWindow(EditWindow, 68, R.Bottom - R.Top - 107, R.Right - R.Left - 90, 23, True)
Call MoveWindow(LabelWindow, 19, R.Bottom - R.Top - 101, 45, 13, True)
Call MoveWindow(SysTreeWindow, 21, TreeTop, R.Right - R.Left - 44, R.Bottom - R.Top - TreeTop - sysH, True)
If Is2K Then
    Call MoveWindow(ScrollWindow, R.Right - R.Left - 24, R.Bottom - R.Top - 44, 16, 16, True)
Else
    Call MoveWindow(ScrollWindow, R.Right - R.Left - 18, R.Bottom - R.Top - 38, 16, 16, True)
End If
If BB.ShowButton Then
    Call MoveWindow(ButtonWindow, R.Right - R.Left - 96, R.Bottom - R.Top - 71, 75, 23, True)
    Call MoveWindow(CancelbuttonWindow, R.Right - R.Left - 177, R.Bottom - R.Top - 71, 75, 23, True)
    Call MoveWindow(OKbuttonWindow, R.Right - R.Left - 258, R.Bottom - R.Top - 71, 75, 23, True)
ElseIf BB.ShowCheck Then
    Call MoveWindow(CancelbuttonWindow, R.Right - R.Left - 96, R.Bottom - R.Top - 71, 75, 23, True)
    Call MoveWindow(OKbuttonWindow, R.Right - R.Left - 177, R.Bottom - R.Top - 71, 75, 23, True)
    Call MoveWindow(ButtonWindow, 20, R.Bottom - R.Top - 71, 110, 23, True)
Else
    Call MoveWindow(CancelbuttonWindow, R.Right - R.Left - 96, R.Bottom - R.Top - 71, 75, 23, True)
    Call MoveWindow(OKbuttonWindow, R.Right - R.Left - 177, R.Bottom - R.Top - 71, 75, 23, True)
End If
If BB.EditBoxOld Then Call MoveWindow(EditWindowOld, 21, EditTop, R.Right - R.Left - 44, EditHeight, True)
If BB.StatusText Then Call MoveWindow(StattxtWindow, 21, StattxtTop, R.Right - R.Left - 44, StattxtHeight, True)
RedrawWindow DialogWindow, ByVal 0&, ByVal 0&, RDW_INVALIDATE
End Sub
Private Sub CleanUp() 'Tidy things up when done
    Call pUnSubClass
    Call pUnSubClassDlg
    Call pUnSubClassFS
    Call pUnSubClassEdit
    DestroyWindow LabelWindow
    DestroyWindow EditWindow
    DestroyWindow ButtonWindow
    DestroyWindow ScrollWindow
    DestroyWindow dummyWindow
End Sub
'***************** WORKER FUNCTIONS **************************
'get text from a window
Private Function gettext(lngwindow As Long) As String
    Dim strBuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strBuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strBuffer$)
    Let gettext$ = strBuffer$
End Function
'parse out just the filename
Private Function FileOnly(ByVal FilePath As String) As String
    If Len(FilePath) = 3 Then
        FileOnly = FilePath
        Exit Function
    End If
    FileOnly = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
End Function
'parse out just the path
Private Function PathOnly(ByVal FilePath As String) As String
Dim temp As String
    temp = Mid$(FilePath, 1, InStrRev(FilePath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
'Fileexists that works with win2K
Private Function FileExists(sSource As String) As Boolean
If Right(sSource, 2) = ":\" Then
    Dim allDrives As String
    allDrives = Space$(64)
    Call GetLogicalDriveStrings(Len(allDrives), allDrives)
    FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
    Exit Function
Else
    If Not sSource = "" Then
        Dim WFD As WIN32_FIND_DATA
        Dim hFile As Long
        hFile = FindFirstFile(sSource, WFD)
        FileExists = hFile <> INVALID_HANDLE_VALUE
        Call FindClose(hFile)
    Else
        FileExists = False
    End If
End If
End Function
'Remove any null characters at the end of a string
Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
'Which operating system ?
Private Function Is2K() As Boolean
On Error Resume Next
Dim tempOSVerInfo As OSVERSIONINFO
Dim DL As Long
tempOSVerInfo.dwOSVersionInfoSize = 148
DL = GetVersionEx(tempOSVerInfo)
If tempOSVerInfo.dwMajorVersion > 4 Then
    Is2K = True
Else
    Is2K = False
End If
End Function

Public Function Browse(strPrompt As String, strTitle As String, strStart As String, _
    hWnd As Long, Optional blnNew As Boolean = False) As String
    'Dim bb As BoboBrowse
    With BB
        'All these settings are optional
        'Leave all of them out and you are
        'left with the default Browse for Folders
        .Titlebar = strTitle
        .Prompt = strPrompt
        If Right$(strStart, 1) = "\" And Len(strStart) > 3 Then
            .InitDir = Left$(strStart, Len(strStart) - 1)
        Else
            .InitDir = strStart
        End If
'        .CHCaption = "Add to Favorites"
'        .OKCaption = "OK"
'        .CancelCaption = "Cancel"
'        .NewFCaption = "New Folder"
        '.RootDir = 0
        .AllowResize = True
        .CenterDlg = True
        .DoubleSizeDlg = False
        .FSDlg = False
        .ShowButton = blnNew
        .ShowCheck = False
        .EditBoxOld = False
        .EditBoxNew = True
        .StatusText = False
        .ShowFiles = False
        .OwnerForm = hWnd
        .CHvalue = 0
        'call the function
        Browse = BrowseFF
        'If you included a checkbox this is where you
        'recieve the users' response
        'blnAddToFav = .CHvalue
    End With
End Function


Public Function StripIllegals(StrIn As String, Optional strSrcDesc As String) As String
Dim intLoop As Integer, strOut As String, strRem As String, _
    strChar As String * 1, strDesc As String

    strOut = ""
    strRem = ""
    For intLoop = 1 To Len(StrIn)
        strChar = Mid$(StrIn, intLoop, 1)
        
        Select Case Asc(strChar)
            Case 34, 42, 47, 58, 60, 62, 63, 92, 124
                strRem = strRem & strChar & "  "
                'If Not blnStrip Then
                strOut = strOut & ReplaceChar(strChar)
            Case 0 To 31, 128, 129, 141 To 144, 157, 158
                strRem = strRem & "ASCII: " & CStr(Asc(strChar)) & "  "
                'If Not blnStrip Then
                strOut = strOut & "_"
            Case Else
                strOut = strOut & strChar
        End Select
    Next intLoop
    If strRem <> "" And strSrcDesc <> "" Then
        'If blnStrip Then
            'strDesc = "removed from "
        'Else
            strDesc = "replaced in "
        'End If
        MsgBox "The following characters were " & strDesc & _
            vbNewLine & strSrcDesc & ":" & _
            vbNewLine & strRem, vbExclamation + vbOKOnly, "Illegal Characters"
    End If
    StripIllegals = strOut
End Function


Public Function ReplaceChar(strIllChar As String) As String
        Select Case strIllChar
            Case "/", "\", "*", "|"
                ReplaceChar = "_"
            Case "?"
                ReplaceChar = "."
            Case ":"
                ReplaceChar = "-"
            Case "<"
                ReplaceChar = "("
            Case ">"
                ReplaceChar = ")"
            Case Chr$(34)
                ReplaceChar = "''"
            Case Else 'Not Illegal!
                ReplaceChar = strIllChar
        End Select
End Function


