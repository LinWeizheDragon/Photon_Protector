VERSION 5.00
Begin VB.UserControl bkDLControl 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ScaleHeight     =   450
   ScaleWidth      =   4650
   ToolboxBitmap   =   "bkDLControl.ctx":0000
End
Attribute VB_Name = "bkDLControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Property Variables:
Private m_sFileURL As String, m_sSaveFilePath As String, blnDownloading As Boolean, sngPct As Single, _
    m_blnFailRedirect As Boolean, m_sSaveFileName As String, m_blnShowProgress As Boolean, _
    blnSuccess As Boolean, m_lFileSize As Long, m_sConn As String, m_lBytesRead As Long, _
    m_sCache As String, m_sRedirect As String, m_sMIMEType As String, m_blnRenameRedirect As Boolean
    
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
'Custom event declarations
Event DLProgress(Percent As Single, BytesRead As Long, TotalBytes As Long)
Event DLCanceled()
Event DLError(E As bkDLError, Error As String)
Event DLComplete(Bytes As Long)
Event DLConnected(ConnAddr As String)
Event DLRedirect(ConnAddr As String)
Event DLCacheFile(FileName As String)
Event DLMIMEType(MIMEType As String)
Event DLFileSize(Bytes As Long)
Event DLBeginDownload()

Public Enum bkDLError
    bkDLEUnavailable = 1
    bkDLERedirect = 2
    bkDLEZeroLength = 3
    bkDLESaveError = 4
    bkDLEUnknown = 99
End Enum
'Private bkDLEUnavailable, bkDLERedirect, bkDLEZeroLength, bkDLESaveError, bkDLEUnknown

'Typical stuff
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'A lot of these properties are runtime-only/read-only - they only have value when
'a DL is happening (or just after)
'Bytes read in so far
Public Property Get BytesRead() As Long
Attribute BytesRead.VB_MemberFlags = "400"
    BytesRead = m_lBytesRead
End Property

'Location of Cache file
Public Property Get CacheFile() As String
Attribute CacheFile.VB_MemberFlags = "400"
    CacheFile = m_sCache
End Property
    
'Address of connection (IP String)
Public Property Get ConnectionAddress() As String
Attribute ConnectionAddress.VB_MemberFlags = "400"
    ConnectionAddress = m_sConn
End Property

'MIME type of download
Public Property Get MIMEType() As String
Attribute MIMEType.VB_MemberFlags = "400"
    MIMEType = m_sMIMEType
End Property

'If redirected, this in the address of the new target
Public Property Get RedirectFile() As String
Attribute RedirectFile.VB_MemberFlags = "400"
    RedirectFile = m_sRedirect
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

'Download complete, attempt to save the file to disk
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error GoTo CompleteError
    Dim bFile() As Byte, FN As Long
    'Internal DL flag
    'Check see if file actuall recieved
    With AsyncProp
        If .BytesRead <> 0 Then
            'write file (in byte array .Value) to disk
            FN = FreeFile
            bFile = .Value
            If m_blnRenameRedirect And m_sRedirect <> vbNullString Then
                SetRedirectName
            End If
            Open m_sSaveFileName For Binary Access Write As #FN
            Put #FN, , bFile
            Close #FN
            blnSuccess = True
            RaiseEvent DLComplete(.BytesRead)
            Kill m_sCache
            blnDownloading = False
        Else
            'Occurs with bad URLs, No internet connection, etc.
            SendError bkDLEZeroLength, "Zero bytes retrieved"
        End If
    End With
    Exit Sub
CompleteError:
    'Typically permissions problem or invalid path
    Debug.Print Err.Number
    SendError bkDLESaveError, Err.Description & " [" & m_sSaveFileName & "]"
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    'Here's the guts of the whole thing!
    'All the interesting message come through this event
    With AsyncProp
        Select Case .StatusCode 'Determines message being recieved
            Case vbAsyncStatusCodeConnecting
                m_sConn = .Status 'Save for Get Property
                RaiseEvent DLConnected(.Status) 'Send back IP address of connection
            Case vbAsyncStatusCodeRedirecting
                m_sRedirect = .Status 'Save for Get Property
                If m_blnFailRedirect Then
                    UserControl.CancelAsyncRead m_sSaveFileName
                    SendError bkDLERedirect, "Redirected to " & .Status  'sends back a path
                    'thought about changing the save file name after Redirect,
                    'but then it's usually a 404error.html file, and who really wants
                    'that saved anyway?
                Else
                    'Keep going, but send message to program than
                    'DL has been redirected
                    RaiseEvent DLRedirect(.Status)
                End If
            Case vbAsyncStatusCodeDownloadingData, vbAsyncStatusCodeEndDownloadData
                'update progress (actual drawing is done in Paint(),
                'so save time if not visible
                If .BytesMax > 0 Then
                    sngPct = CSng(.BytesRead / .BytesMax)
                Else
                    sngPct = 0!
                End If
                m_lBytesRead = .BytesRead 'Save for Get Property
                'ChangeToolTip 'discarded
                RaiseEvent DLProgress(sngPct, .BytesRead, .BytesMax)
            Case vbAsyncStatusCodeMIMETypeAvailable
                'Another tidbit of info
                m_sMIMEType = .Status 'Save for Get Property
                RaiseEvent DLMIMEType(.Status)
            Case vbAsyncStatusCodeCacheFileNameAvailable
                'location of the local Cache file
                m_sCache = .Status 'Save for Get Property
                RaiseEvent DLCacheFile(.Status)
            Case vbAsyncStatusCodeBeginDownloadData
                'Connected, data transfer commenced.
                'Now we know the file size and can report it
                'This could also have gone under
                'vbAsyncStatusCodeCacheFileNameAvailable
                'Which occurs first, but this looks a little neater
                m_lFileSize = .BytesMax 'Save for Get Property
                RaiseEvent DLFileSize(.BytesMax)
                RaiseEvent DLBeginDownload
            Case vbAsyncStatusCodeError
                'Never found a situation that triggered this
                'help says error msg is in Value not Status, but then
                'there was one other typo on that page already...
                Debug.Print "ERROR: ", .Status, Now 'just in case
                SendError bkDLEUnknown, CStr(.Value)
        End Select
    End With
    UserControl.Refresh
End Sub

'Typical event wrappers
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'More Run-time/Read-only
'Total size of file in bytes
Public Property Get FileSize()
Attribute FileSize.VB_MemberFlags = "400"
    FileSize = m_lFileSize
End Property

'The URL to be downloaded from
Public Property Get FileURL() As String
Attribute FileURL.VB_Description = "URL of file to be Downloaded"
Attribute FileURL.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute FileURL.VB_MemberFlags = "400"
    FileURL = m_sFileURL
End Property

Public Property Let FileURL(ByVal New_FileURL As String)
    m_sFileURL = New_FileURL
    'determine the full filename
    SetFileName
    PropertyChanged "FileURL"
End Property

'Full filename: read-only at runtime
'Made from the File specified at the end of FileURL and
'the SaveFilePath
Public Property Get SaveFileName() As String
Attribute SaveFileName.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute SaveFileName.VB_MemberFlags = "400"
    SaveFileName = m_sSaveFileName
End Property

'Folder location to send all downloaded files to
Public Property Get SaveFilePath() As String
Attribute SaveFilePath.VB_Description = "Path to Save downloaded file to"
Attribute SaveFilePath.VB_ProcData.VB_Invoke_Property = ";Misc"
    SaveFilePath = m_sSaveFilePath
End Property

Public Property Let SaveFilePath(ByVal New_SaveFilePath As String)
    m_sSaveFilePath = New_SaveFilePath
    'determine the full filename
    SetFileName
    PropertyChanged "SaveFilePath"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_sFileURL = vbNullString
    m_sSaveFilePath = vbNullString
    'Defaults to FailOnRedirect=True on the assumtion
    'that it is most often a redirect to a 404error.html file!
    m_blnFailRedirect = True
    'If not fail, next best bet is to rename (i.e., save
    '"path\404error.html" rather than the intended filename)
    'Note that this will be shown as False in prop browser if
    'FailOnRedirect is still True!
    m_blnRenameRedirect = True
    'Use the control itself as a progress bar
    InitDL ' blank out the 'get property' fields
    m_blnShowProgress = True
End Sub

'Progress bar drawing here
'Here it gets done only when necessary (control visible on screen)
Private Sub UserControl_Paint()
    If m_blnShowProgress And sngPct > 0! Then
        UserControl.Line (0, 0)-(UserControl.Width * sngPct, UserControl.Height), UserControl.ForeColor, BF
    End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_sFileURL = PropBag.ReadProperty("FileURL", vbNullString)
    m_sSaveFilePath = PropBag.ReadProperty("SaveFilePath", vbNullString)
    m_blnFailRedirect = PropBag.ReadProperty("FailOnRedirect", True)
    m_blnRenameRedirect = PropBag.ReadProperty("RenameOnRedirect", True)
    m_blnShowProgress = PropBag.ReadProperty("ShowProgress", True)
End Sub

'cancel any ongoing downloads
Private Sub UserControl_Terminate()
    If blnDownloading Then
        On Error Resume Next 'Might throw error if just begun and
                                'first conn. has not yet occured
        UserControl.CancelAsyncRead m_sSaveFileName
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", UserControl.BackColor, &H8000000F
    PropBag.WriteProperty "ForeColor", UserControl.ForeColor, &H80000012
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "Font", Font, Ambient.Font
    PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle, 1
    PropBag.WriteProperty "FileURL", m_sFileURL, vbNullString
    PropBag.WriteProperty "SaveFilePath", m_sSaveFilePath, vbNullString
    PropBag.WriteProperty "FailOnRedirect", m_blnFailRedirect, True
    PropBag.WriteProperty "RenameOnRedirect", m_blnRenameRedirect, True
    PropBag.WriteProperty "ShowProgress", m_blnShowProgress, True
End Sub

'Trigger the
Public Function BeginDownload(Optional Wait As Boolean = False) As Boolean
    If blnDownloading Then Exit Function
    'check that we have a "to" and "from"...
    If m_sFileURL = vbNullString Or m_sSaveFilePath = vbNullString Then Exit Function
    On Error GoTo BeginDownloadError
    'here's the heart of it:
    UserControl.AsyncRead m_sFileURL, vbAsyncTypeByteArray, m_sSaveFileName, vbAsyncReadForceUpdate
    blnDownloading = True 'Internal check
    'blank of the dl property gets
    InitDL
    If Wait Then
        'wait until the downlad is complete, then return success or failure
        'Disadvantage: main code can't Cancel the download with this option
        'because main code is suspended!
        DoWait
        BeginDownload = blnSuccess
    Else
        BeginDownload = True 'Signal successful start and return to main code
    End If
    Exit Function
BeginDownloadError:
    SendError bkDLEUnavailable, Err.Description
    MsgBox Err & "Error: " & vbCrLf & Err.Description, vbCritical, "bkDLControl Internal Error: " & CStr(Err.Number)
End Function

Private Sub InitDL()
    m_lFileSize = 0&
    m_lBytesRead = 0&
    m_sConn = vbNullString
    m_sCache = vbNullString
    m_sRedirect = vbNullString
    m_sMIMEType = vbNullString
    blnSuccess = False
End Sub

'Basic loop 'til control variable changes
'the Async download will continue to fire events during this loop
'At some point, the download will complete itself or fail, and
'blndownloading will be false and the loop will exit
Private Sub DoWait()
    Do
        DoEvents
    Loop Until Not blnDownloading
End Sub

Public Sub CancelDownload()
    If Not blnDownloading Then Exit Sub
    'Throws an error if DL really hasn't started yet (no progress)
    '(safe to ignore error)
    On Error Resume Next
    UserControl.CancelAsyncRead m_sSaveFileName
    On Error GoTo 0
    sngPct = 0!
    Refresh
    blnDownloading = False
    RaiseEvent DLCanceled
End Sub

'If download is re-directed we might not get the file
'we wanted.  Defaults to fail (DLError) if redirected
Public Property Get FailOnRedirect() As Boolean
    FailOnRedirect = m_blnFailRedirect
End Property

Public Property Let FailOnRedirect(NewFail As Boolean)
    m_blnFailRedirect = NewFail
    PropertyChanged "FailOnRedirect"
End Property

'If download is re-directed we might not get the file
'we wanted., but we can save whatever we do get under
'it's original name
Public Property Get RenameOnRedirect() As Boolean
Attribute RenameOnRedirect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RenameOnRedirect = m_blnRenameRedirect And Not m_blnFailRedirect
End Property

Public Property Let RenameOnRedirect(NewRename As Boolean)
    m_blnRenameRedirect = NewRename
    PropertyChanged "RenameOnRedirect"
End Property

'Use control as progress bar
'Control also still has .hWnd and .hDC, so you can
'draw your own if you don't like my primitive prog bar.
Public Property Get ShowProgress() As Boolean
Attribute ShowProgress.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowProgress = m_blnShowProgress
End Property

Public Property Let ShowProgress(NewShowProgress As Boolean)
    m_blnShowProgress = NewShowProgress
    PropertyChanged "ShowProgress"
End Property

'This works but it is so ugly I took it out!
'Public Sub ChangeToolTip()
'Dim sName As String, iOpenParen As Integer, iCloseParen As Integer
'    sName = UserControl.Ambient.DisplayName
'    'Get reference to self and change own tooltip
'    'Would have sworn there was an easier way but can't find it.  Brain fart.
'    On Error GoTo InArray
'    UserControl.Parent.Controls.Item(sName).ToolTipText = Format(sngPct, "0%")
'    Exit Sub
'InArray:
'    If Err = 730 Then
'        iOpenParen = InStr(1, sName, "(")
'        iCloseParen = InStr(iOpenParen, sName, ")")
'        UserControl.Parent.Controls.Item(Left(sName, iOpenParen - 1), _
'            CInt(Mid(sName, iOpenParen + 1, iCloseParen - iOpenParen - 1))).ToolTipText = Format(sngPct, "0%")
'    End If
'End Sub

'Little function to combine path & filename
Private Function getFullPath(strPath As String, strFile As String, Optional strDelim As String = "\") As String
    If Right$(strPath, 1) = strDelim Then
        getFullPath = strPath & strFile
    Else
        getFullPath = strPath & strDelim & strFile
    End If
End Function

'little function to retrieve filename (text after last "\" in path)
Private Function getFileFromPath(strPath As String, Optional strDelim As String = "\") As String
Dim iPos As Integer
    iPos = InStrRev(strPath, strDelim)
    If iPos = 0 Then
        getFileFromPath = strPath
    Else
        getFileFromPath = Mid$(strPath, iPos + 1)
    End If
End Function

'The filename is made of the path + the file name from the URL
'i.e., if path is "C:\Downloads" and url is '
'"http://www.blueknot.com/Downloads/ProjecTile.zip"
'the SaveFileName will be "C:\Downloads\ProjecTile.zip"
Private Sub SetFileName()
    If m_sFileURL = vbNullString Or m_sSaveFilePath = vbNullString Then
        m_sSaveFileName = vbNullString
    Else
        m_sSaveFileName = getFullPath(m_sSaveFilePath, getFileFromPath(Replace$(m_sFileURL, "/", "\")))
    End If
End Sub

'Like above, but use the redirect path to get the 'new' file name
Private Sub SetRedirectName()
    m_sSaveFileName = getFullPath(m_sSaveFilePath, getFileFromPath(Replace$(m_sRedirect, "/", "\")))
End Sub

'When stopping on an error, a message is sent to the DLError event,
'Plus the DLComplete event fires w/ 0& bytes downloaded
Private Sub SendError(E As bkDLError, strMessage As String)
    sngPct = 0!
    Refresh
    blnDownloading = False
    RaiseEvent DLError(E, strMessage)
    RaiseEvent DLComplete(0&)
End Sub
