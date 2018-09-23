VERSION 5.00
Begin VB.Form frmDLTest 
   Caption         =   "bkDownload Control Test Form"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmDLTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11580
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin prjDLTest.bkDLControl DL 
      Height          =   225
      Index           =   0
      Left            =   180
      Top             =   1410
      Width           =   5565
      _extentx        =   9816
      _extenty        =   397
      forecolor       =   -2147483635
      font            =   "frmDLTest.frx":014A
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Log"
      Height          =   285
      Left            =   9240
      TabIndex        =   12
      Top             =   6690
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   7530
      TabIndex        =   10
      Top             =   2610
      Width           =   1575
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Index           =   3
      Left            =   5940
      TabIndex        =   9
      Top             =   2610
      Width           =   1575
   End
   Begin VB.ComboBox cboURL 
      Height          =   315
      ItemData        =   "frmDLTest.frx":0176
      Left            =   3420
      List            =   "frmDLTest.frx":0195
      TabIndex        =   0
      Top             =   120
      Width           =   8025
   End
   Begin VB.CommandButton cmdBrowse 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11160
      TabIndex        =   2
      Top             =   510
      Width           =   285
   End
   Begin VB.ListBox lstOut 
      Appearance      =   0  'Flat
      Height          =   3450
      ItemData        =   "frmDLTest.frx":0431
      Left            =   0
      List            =   "frmDLTest.frx":0433
      TabIndex        =   11
      Top             =   3060
      Width           =   11535
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   3390
      TabIndex        =   1
      Top             =   510
      Width           =   7755
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Index           =   2
      Left            =   5910
      TabIndex        =   5
      Top             =   1650
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   7500
      TabIndex        =   6
      Top             =   1650
      Width           =   1575
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   2610
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   2610
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1770
      TabIndex        =   4
      Top             =   1650
      Width           =   1575
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1650
      Width           =   1575
   End
   Begin prjDLTest.bkDLControl DL 
      Height          =   225
      Index           =   1
      Left            =   180
      Top             =   2370
      Width           =   5565
      _extentx        =   9816
      _extenty        =   397
      font            =   "frmDLTest.frx":0435
   End
   Begin prjDLTest.bkDLControl DL 
      Height          =   225
      Index           =   2
      Left            =   5910
      Top             =   1410
      Width           =   5565
      _extentx        =   9816
      _extenty        =   397
      forecolor       =   255
      font            =   "frmDLTest.frx":0461
      failonredirect  =   0   'False
   End
   Begin prjDLTest.bkDLControl DL 
      Height          =   225
      Index           =   3
      Left            =   5970
      Top             =   2370
      Width           =   5565
      _extentx        =   9816
      _extenty        =   397
      font            =   "frmDLTest.frx":048D
      failonredirect  =   0   'False
      renameonredirect=   0   'False
   End
   Begin VB.Label lblInstruct 
      AutoSize        =   -1  'True
      Caption         =   "#2 has RenameOnRedirect set to true, #2 and 3 have FailOnRedirect set to false -- try them with last URL in combo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   27
      Top             =   870
      Width           =   10230
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   5790
      TabIndex        =   26
      Top             =   2370
      Width           =   120
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   5790
      TabIndex        =   25
      Top             =   1380
      Width           =   120
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   24
      Top             =   2340
      Width           =   120
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   23
      Top             =   1410
      Width           =   120
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Folder to save files to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   22
      Top             =   570
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInstruct 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Select or type URL to download from:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   21
      Top             =   180
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   3
      Left            =   11460
      TabIndex        =   20
      Top             =   2610
      Width           =   45
   End
   Begin VB.Label lblFile 
      Height          =   255
      Index           =   3
      Left            =   5970
      TabIndex        =   19
      Top             =   2100
      Width           =   5565
   End
   Begin VB.Label lblFile 
      Height          =   255
      Index           =   2
      Left            =   5940
      TabIndex        =   18
      Top             =   1140
      Width           =   5565
   End
   Begin VB.Label lblProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   2
      Left            =   11430
      TabIndex        =   17
      Top             =   1650
      Width           =   45
   End
   Begin VB.Label lblFile 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   2100
      Width           =   5565
   End
   Begin VB.Label lblProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   5730
      TabIndex        =   15
      Top             =   2610
      Width           =   45
   End
   Begin VB.Label lblProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   5700
      TabIndex        =   14
      Top             =   1650
      Width           =   45
   End
   Begin VB.Label lblFile 
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   1170
      Width           =   5565
   End
End
Attribute VB_Name = "frmDLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Note: the last two items in the combo are links to my site, but I didn't just put them in
'for 'free publicity' ;-)
'The point is that the second one doesn't exist - and rather than return 0 bytes,
'my host redirects the request to a 404error.html file.  Try to download 'FreeCheese.zip'
'and see what happens
'#0 and #1 Fail with an error because FailOnRedirect is True
'#2 and #3 have FailOnRedirect = False
'#2 has RenameOnRedirect set to True, so the name of the redirected file will be used to
'save what gets downloaded (404error.html)
'#3 has RenameOnRedirect set to False, so the original filename (FreeCheese.zip) will be
'used -- but drag the 'zip' file into notepad and look at it.
Private Sub cboURL_GotFocus()
    With cboURL
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdBegin_Click(Index As Integer)
    With DL(Index)
        .FileURL = cboURL.Text
        .SaveFilePath = txtPath.Text
        ClearLabels Index
        LogItem Index, "Requesting Download of " & cboURL.Text
        SetCancel Index, .BeginDownload 'Function returns True if successful
        'note that if we send True as the parameter of BeginDownload,
        'the program would have stopped until the download ended.
        'It would then return True if the d/l was successful, False if it failed
    End With
End Sub

Private Sub cmdBrowse_Click()
Dim strTemp As String
    'Thanks to Mr. Bobo for browse for folder routines
    'It's people like that who make me want to contribute to PSC!
    strTemp = Browse("Select folder to save files to", "Save Directory", txtPath.Text, Me.hWnd, False)
    If strTemp <> vbNullString Then
        txtPath.Text = strTemp
    End If
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    'User changes mind (probably right after seeing file size!)
    DL(Index).CancelDownload
    SetCancel Index, False
    ClearLabels Index
End Sub

Private Sub cmdClear_Click()
    'zap
    lstOut.Clear
End Sub

'The following are all events recieved from the DL control.
'I chose to make them seperate events rather than a single event
'with a status code to make the end code more readable and
'more easily give new programmers access to functions they
'might not realize were there.
Private Sub DL_DLBeginDownload(Index As Integer)
    LogItem Index, "Download started from " & DL(Index).FileURL
    With lblFile(Index)
        .ToolTipText = DL(Index).SaveFileName
        .Caption = FitPathToSize(.ToolTipText, .Width)
    End With
End Sub

Private Sub DL_DLCacheFile(Index As Integer, FileName As String)
    'returns local cache file location
    LogItem Index, "Cache File: " & FileName
End Sub

Private Sub DL_DLCanceled(Index As Integer)
    'canceled by user
    ClearLabels Index
    LogItem Index, "Download Canceled"
End Sub

Private Sub DL_DLComplete(Index As Integer, Bytes As Long)
    'download terminated - bytes is > 0 if successful (file size)
    If Bytes > 0& Then
        LogItem Index, "Complete. " & SizeString(Bytes) & " downloaded and saved as " & DL(Index).SaveFileName
    Else
        LogItem Index, "Download failed."
    End If
    SetCancel Index, False
End Sub

'Returns IP address of successful connection
Private Sub DL_DLConnected(Index As Integer, ConnAddr As String)
    LogItem Index, "Connected to " & ConnAddr
End Sub
'Error!  See UC code for different possible errors
'This event is always followed by DLComplete returning 0 bytes
Private Sub DL_DLError(Index As Integer, E As bkDLError, Error As String)
Dim strErrType As String
    Select Case E
        Case bkDLEUnavailable
            strErrType = "Download Unavailable"
        Case bkDLERedirect
            strErrType = "Redirected"
        Case bkDLEZeroLength
            strErrType = "Zero Bytes Returned"
        Case bkDLESaveError
            strErrType = "File Save Error"
        Case bkDLEUnknown
            strErrType = "Unknown"
    End Select
    ClearLabels Index
    LogItem Index, "Error - " & strErrType & ": " & Error
End Sub

Private Sub DL_DLFileSize(Index As Integer, Bytes As Long)
    'Size in bytes.  returned when connection to file is complete
    'and download actually begins
    LogItem Index, "File size is " & SizeString(Bytes) & " (" & CStr(Bytes) & " bytes)"
End Sub

Private Sub DL_DLMIMEType(Index As Integer, MIMEType As String)
    'handy info!
    LogItem Index, "MIME type is " & MIMEType
End Sub

Private Sub DL_DLProgress(Index As Integer, Percent As Single, BytesRead As Long, TotalBytes As Long)
    'Progress two ways: Percentage, or BytesRead vs. Total Bytes (yeah, I know, with that
    'you can figure it out yourself, but since I was already calculating it for the
    'control figured I'd save you the duplication of work and pass it on!
    'Hey, this is source code-- change it if you don't like it!
    lblProg(Index) = Format(Percent, "0%") & " of " & SizeString(TotalBytes)
End Sub

Private Sub DL_DLRedirect(Index As Integer, ConnAddr As String)
    'Returns path to file if redirected
    'This event wont fire at all if FailOnRedirect is True! (DLError instead)
    LogItem Index, "Redirected to " & ConnAddr
End Sub

Private Sub Form_Load()
    'initialize sample inputs
    txtPath.Text = App.Path
    cboURL.ListIndex = 0
End Sub

Private Sub txtPath_GotFocus()
    txtPath.SelStart = Len(txtPath.Text)
End Sub

'Common Functions
Private Sub ClearLabels(Index As Integer)
    lblFile(Index) = vbNullString
    lblProg(Index) = vbNullString
End Sub

Private Sub SetCancel(Index As Integer, blnCancel As Boolean)
    cmdCancel(Index).Enabled = blnCancel
    cmdBegin(Index).Enabled = Not blnCancel
End Sub

Private Sub LogItem(Index As Integer, strItem As String)
    With lstOut
        .AddItem CStr(Index) & "> " & strItem
        If .NewIndex > .TopIndex + 17 Then
            'Yes, I cheated and hard-coded the numbers rather than
            'figure out how many lines are in the listbox through code.
            'List boxes are not the point of this project! ;-)
            .TopIndex = .NewIndex - 16
        End If
    End With
End Sub

'Misc Functions you may find useful...
'Convert size in bytes to string representation in
Private Function SizeString(lBytes As Long) As String
    If lBytes < &H400& Then '1024 = 1K
        SizeString = CStr(lBytes) & "b"
    ElseIf lBytes < &H100000 Then '1024 ^ 2 = 1M
        SizeString = CStr(lBytes \ 1024) & "k"
    ElseIf lBytes < &H20000000 Then  '1024 ^ 2 * 512 = up to 0.5G
        SizeString = Replace$(Format$((lBytes \ 1024) / 1024, "0.0"), ".0", vbNullString) & "M"
    Else 'Not bothering to code for Terrabytes...
        'If you're doing that you should probably be using a more robust control!
        SizeString = Replace$(Format$((lBytes \ (1024 ^ 2)) / 1024, "#,##0.0"), ".0", vbNullString) & "G"
    End If
End Function

'Truncate path to fit size but leave filename
Private Function FitPathToSize(strPath As String, lTarget As Long) As String
Dim iPos As Integer, iLastSlash As Integer, strEnd As String, lSize As Long, strTemp As String
    'Yes, I know this only works when the Form font
    'matches the Label you're putting it in...
    'so just make sure it does!
    strTemp = strPath
    iLastSlash = InStrRev(strPath, "\")
    If iLastSlash = 0 Then
        FitPathToSize = strPath
        Exit Function
    End If
    lSize = Me.TextWidth(strTemp)
    iPos = InStrRev(strPath, "\", iLastSlash - 1)
    Do While iPos > 1 And lSize > lTarget
        strTemp = Left$(strPath, iPos) & "..." & Mid$(strPath, iLastSlash)
        lSize = Me.TextWidth(strTemp)
        iPos = InStrRev(strPath, "\", iPos - 1)
    Loop
    FitPathToSize = strTemp
End Function
