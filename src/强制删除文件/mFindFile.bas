Attribute VB_Name = "mFindFile"
Public mydirectory() As String
Public intresult As Integer

Public Sub Findfile(getPath As String, lst As ListBox)
Dim mypath As String
Dim myname As String
Dim i As Integer
    mypath = getPath
    If mypath = "" Then Exit Sub
    intresult = 2
    ReDim mydirectory(intresult)
    mydirectory(1) = mypath
    i = 1
    Do Until mydirectory(i) = ""
        mypath = mydirectory(i)
        If Right(mypath, 1) <> "\" Then mypath = mypath & "\"
        myname = Dir(mypath, vbDirectory)
            Do While myname <> ""
                If myname <> "." And myname <> ".." Then
                    If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
                        mydirectory(intresult) = mypath & myname
                        intresult = intresult + 1
                        ReDim Preserve mydirectory(intresult)
                    Else
                        lst.AddItem mypath & myname
                    End If
                End If
                myname = Dir
            Loop
        i = i + 1
    Loop
End Sub
