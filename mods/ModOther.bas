Attribute VB_Name = "ModOther"
Public MenuHangle As Long
Public DialogHangle As Long
Public ResFileName As String
Public SaveOption As String ' this tells the save button what action to save based on the res type

Function GetWindowPosition(ByVal hangle As Long) As Variant()
Dim Holder(3) As Variant
Dim mRect As RECT
' I made this to find the position of a window
    
    GetWindowRect hangle, mRect
    Holder(0) = mRect.Left ' left
    Holder(1) = mRect.Top ' top
    Holder(2) = (mRect.Right - mRect.Left) ' window Width
    Holder(3) = (mRect.Bottom - mRect.Top) ' window Height
    GetWindowPosition = Holder
    
End Function

Public Function CleanStr(lzStr As String) As String
Dim StrA As String
   ' this just tidys the string by removeing the linefeed and carriage return
    StrA = lzStr
    StrA = Replace(StrA, vbLf, "") ' Remove linefeed
    StrA = Replace(StrA, vbCr, vbNewLine) ' Remove carriage return and replace with a newline
    CleanStr = StrA
    StrA = ""
End Function

Public Function GetFileExt(lzFile As String) As String
Dim ipos As Integer
    ' used to get the files ext eg GetFileExt "hello.txt" returns txt
    ipos = InStrRev(lzFile, ".", Len(lzFile), vbTextCompare)
    GetFileExt = Mid(lzFile, ipos + 1, Len(lzFile))
End Function

Public Function GetFileTitle(lzFile As String) As String
Dim ipos As Long
    ' used to get the given filename from a file and it's path
    ' eg GetFileTitle "c:\thispath\ben.txt" returns ben.txt
    ipos = InStrRev(lzFile, "\", Len(lzFile), vbTextCompare) ' Get the last backslash
    If ipos <> 0 Then GetFileTitle = Mid(lzFile, ipos + 1, Len(lzFile)) ' Return the filename
End Function

Public Function RemoveFileExt(lzFile As String) As String
    ' used to remove the file extension from a given file
    ' eg RemoveFileExt "ben.txt" returns ben
    ipos = InStrRev(lzFile, ".", Len(lzFile), vbTextCompare) ' Get the last . before the extension
    If ipos <> 0 Then RemoveFileExt = Mid(lzFile, 1, ipos - 1) ' Return the filename with no extension
End Function

Function RemoveLeftStr(dwStrA As String, nPlaces As Long) As String
    ' This function is used to removes chars form the left side of a string given it's nPlaces to remove
    ' eg RemoveLeftStr "#128",1 ' returns 128
    RemoveLeftStr = Right(dwStrA, Len(dwStrA) - nPlaces)
End Function

Public Function RemoveTempFile(sTempFileNameA As String)
On Error Resume Next
    If Len(sTempFileNameA) <> 0 And IsFileHere(sTempFileNameA) Then
        ' Line above checks if the file is found and length of file name is more than zero
        ' if the return is true the temp file is then deleted
        Kill sTempFileNameA
    End If
End Function

Public Function IsFileHere(lzFilename As String) As Boolean
    ' Checks if a given filename is found
    If Dir(lzFilename) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Function SaveToFile(lzFile As String, sData As String)
Dim nFile As Long
    nFile = FreeFile ' pointer to free file
    Open lzFile For Binary As #nFile
        Put #nFile, , sData ' save sdata contents to the file
    Close #nFile ' close file
End Function

Public Function GetTempPathA() As String
Dim TmpResult As Long
Dim TmpStrBuff As String
' function to get the temp folder path on your system
    TmpStrBuff = Space(216)
    TmpResult = GetTempPath(Len(TmpStrBuff), TmpStrBuff)
    If TmpResult <> 0 Then GetTempPathA = Left(TmpStrBuff, TmpResult)
    
    TmpResult = 0
    TmpStrBuff = ""
    
End Function

Public Function WriteByteArrayToFile(lzFile As String, bbByte() As Byte)
Dim nFile As Long
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        Put #nFile, , bbByte ' save bbByte contents to the file
    Close #nFile '
End Function
'DlgTitle As String, dlgFilter As String, ResType As Integer, Optional DefFilename As String = "", Optional SaveData As String) As Boolean
Public Function GetDescFromFileExt(lzFile As String) As Variant
' This small function builds an array with properties for the common dialog
' I added this because some resource types were JPG,GIF and Html
' and I wanted the dialog to be more meaningfull
Dim vList(2) As Variant

    Select Case LCase(GetFileExt(lzFile))
        Case "jpg", "jpeg"
            vList(0) = "JPEG" ' dialogs title
            vList(1) = "JPG (*.jpg)|*.jpg|JPEG(*.jpeg)|*.jpeg|JPE(*.jpe)|*.jpe|JFIF(*.JFIF)|*.JFIF|"
            vList(2) = RemoveFileExt(lzFile)
            GetDescFromFileExt = vList
            Erase vList
            Exit Function
        Case "gif"
            vList(0) = "GIF" ' dialogs Title
            vList(1) = "GIF Files(*.gif)|*.gif|" ' dialogs Filter
            vList(2) = RemoveFileExt(lzFile)
            GetDescFromFileExt = vList
            Erase vList
            Exit Function
        Case "htm", "html", "xml", "asp", "css"
            vList(0) = "Web Page" ' dialogs Title
            vList(1) = "HTM (*.htm)|*.htm|HTML(*.html)|*.html|XML(*.xml)|*.xml|CSS(*.css)|*.css|ASP(*.asp)|*.asp|" ' dialogs Filter
            vList(2) = RemoveFileExt(lzFile)
            GetDescFromFileExt = vList
            Erase vList
            Exit Function
        Case Else
            vList(0) = lzFile ' dialogs Title
            vList(1) = "All Files(*.*)|*.*|" ' dialogs Filter
            vList(2) = lzFile
            GetDescFromFileExt = vList
            Erase vList
            Exit Function
    End Select
    
End Function

