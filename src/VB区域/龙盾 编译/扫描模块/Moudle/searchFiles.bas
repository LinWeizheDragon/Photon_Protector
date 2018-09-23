Attribute VB_Name = "searchFiles"

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
Private Sub Showfilelist(folderspec)
     Dim fs, f, f1, fc, s
     Set fs = CreateObject("Scripting.FileSystemObject")
     Set f = fs.GetFolder(folderspec)
     Set fc = f.Files
     For Each f1 In fc
     'List1.AddItem f1.Name
     Next
End Sub


'遍历某文件夹及子文件夹中的所有文件
Public Sub sousuofile(MyPath As String)
On Error Resume Next
DoEvents '转让控制权，防止假死。

Dim Myname As String
Dim a As String
Dim B() As String
Dim dir_i() As String
Dim I, idir As Long
If Right(MyPath, 1) <> "\" Then MyPath = MyPath + "\"
Myname = Dir(MyPath, vbDirectory Or vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
Do While Myname <> ""
If Myname <> "." And Myname <> ".." Then
If (GetAttr(MyPath & Myname) And vbDirectory) = vbDirectory Then '如果找到的是目录
idir = idir + 1
ReDim Preserve dir_i(idir) As String
dir_i(idir - 1) = Myname
Else

'frmMain.lstFiles.AddItem MyPath & Myname   '把找到的文件显示到列表框中

End If
End If
Myname = Dir '搜索下一项
Loop
For I = 0 To idir - 1
Call sousuofile(MyPath + dir_i(I))
Next I
ReDim dir_i(0) As String
End Sub

       '在这里可以处理目录中的文件
       'Fn.Name       '得到文件名
       'Fn.Size       '得到文件大小
       'Fn.Path       '得到文件路径
       'Fn.Type       '得到文件类型
       'Fn.DateLastModified       '得到文件最后的修改日期

