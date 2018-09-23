Attribute VB_Name = "searchFiles"
Public FilePathGroup(1000000) As String
Public filePathNum
Public Sub ShowFolderList(folderspec)
     Dim fs, f, f1, s, sf
     Dim hs, h, h1, hf
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set sf = f.SubFolders
     For Each f1 In sf
     'frmMain.lstFiles.AddItem folderspec & f1.Name
     Next
End Sub



'遍历某文件夹下的文件
Public Sub Showfilelist(folderspec, List As ListBox)
     Dim fs, f, f1, fc, s
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set fc = f.Files
     For Each f1 In fc
     If selectover = True Then Exit Sub
     If frmMain.StopScan = True Then
   Do Until frmMain.StopScan = False Or frmMain.TeScan = True
   SuperSleep 1
   Loop
   If frmMain.StopScan = False Then
   End If
   If frmMain.TeScan = True Then
   frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   
   FileList.ListItems.Remove 1
   frmMain.Time.Enabled = False
   frmRow.DoClean
   Exit Sub
   End If
   End If
     
FilePathGroup(filePathNum) = f1.Name
frmMain.Lbl_Status.Caption = "搜索文件夹......" & FilePathGroup(filePathNum)
filePathNum = filePathNum + 1
     Next
End Sub


'遍历某文件夹及子文件夹中的所有文件
Public Sub sousuofile(MyPath As String, List As ListBox)
On Error Resume Next
DoEvents '转让控制权，防止假死。
Dim Myname As String
Dim a As String
Dim b() As String
Dim dir_i() As String
Dim i, idir As Long


If Right(MyPath, 1) <> "\" Then MyPath = MyPath + "\"
Myname = Dir(MyPath, vbDirectory Or vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
Do While Myname <> ""
If Myname <> "." And Myname <> ".." Then
If (GetAttr(MyPath & Myname) And vbDirectory) = vbDirectory Then '如果找到的是目录
idir = idir + 1
ReDim Preserve dir_i(idir) As String
dir_i(idir - 1) = Myname
Else
If selectover = True Then Exit Sub
If frmMain.StopScan = True Then
   Do Until frmMain.StopScan = False Or frmMain.TeScan = True
   SuperSleep 1
   Loop
   If frmMain.StopScan = False Then
   End If
   If frmMain.TeScan = True Then
   frmMain.Lbl_Object.Caption = ""
   frmMain.Lbl_Status.Caption = "终止"
   frmMain.Lbl_Target.Caption = "暂无"
   selectover = True
   
   FileList.ListItems.Remove 1
   frmMain.Time.Enabled = False
   frmRow.DoClean
   Exit Sub
   End If
End If
FilePathGroup(filePathNum) = MyPath & Myname
'List.AddItem MyPath & Myname   '把找到的文件显示到列表框中
frmMain.Lbl_Status.Caption = "搜索文件夹......" & FilePathGroup(filePathNum)
filePathNum = filePathNum + 1
End If
End If
Myname = Dir '搜索下一项
Loop
For i = 0 To idir - 1
Call sousuofile(MyPath + dir_i(i), List)
Next i
ReDim dir_i(0) As String
End Sub

       '在这里可以处理目录中的文件
       'Fn.Name       '得到文件名
       'Fn.Size       '得到文件大小
       'Fn.Path       '得到文件路径
       'Fn.Type       '得到文件类型
       'Fn.DateLastModified       '得到文件最后的修改日期

