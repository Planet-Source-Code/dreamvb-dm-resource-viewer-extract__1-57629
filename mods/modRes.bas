Attribute VB_Name = "ModRes"
Public nCallBack As New ClsMain

Public Function GetStringTable(DmResName As String, DmResType As String) As Variant
Dim mHangle As Long
Dim ResNameVal As Long
Dim StrA As String, StrTableList As String, ipos As Long, iCount As Long

    mHangle = LoadLibrary(ResFileName) ' Get the files libary habgle
    If mHangle = 0 Then GetStringTable = "": Exit Function ' Return false if we can't find the program hangle

    ResNameVal = RemoveLeftStr(DmResName, 1) ' Strip of leading # to the left and convert to a long datatype

    StrA = Space(128) ' Fill our string with spaces

    For iCount = ((ResNameVal - 1) * 16) To (ResNameVal * 16) ' loop till we find the last string index in the table
        ipos = LoadString(mHangle, iCount, StrA, Len(StrA)) ' load the string based on icount
        If ipos <> 0 Then ' check if we found a string
            StrTableList = StrTableList & "    " & iCount & "                      " & Chr(34) & Left(StrA, ipos) & Chr(34) & vbCrLf
            ' Line above builds the new string list
        End If
   Next
   
    StrTableList = CleanStr(StrTableList) 'Clean the string see CleanStr for more info
    
    StrTableList = "STRINGTABLE DISCARDABLE" & " //" & GetFileTitle(ResFileName) & ResNameVal & vbCrLf & vbCrLf & "BEGIN" & vbCrLf & StrTableList & "END" & vbCrLf
    ' Code above adds the title string table and the appliaction name it came from
    GetStringTable = StrTableList
    
    ' Clean up vars
    FreeLibrary mHangle
    ResNameVal = 0
    ResNameVal = 0
    iCount = 0
    ipos = 0
    StrA = ""
    StrTableList = ""
End Function

Public Property Get HiWord(LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Property

Public Function GetResourceBlob(DmResName As String, DmResType As String) As Variant
'GetResourceBlob returns the data form a resource in a file I called it a blob but it ready just binary data
' this is usefull when dealing with Bitmaps, and Some Icon files the data is returnd as a stream of data

Dim mHangle As Long, hObject As Long
Dim FindRes As Long, Reslock As Long, ResSize As Long
Dim DataByte() As Byte
Dim dmDwResName As String, dmDwResType As String

    mHangle = LoadLibrary(ResFileName) ' Get the files libary habgle
    If mHangle = 0 Then GetResourceBlob = "": Exit Function ' Return false if we can't find the program hangle

    If IsNumeric(RemoveLeftStr(DmResName, 1)) Then
        dmDwResName = DmResName
    Else
        dmDwResName = RemoveLeftStr(DmResName, 1)
    End If

    If IsNumeric(RemoveLeftStr(DmResType, 1)) Then
        ' the above is used to check if the resname is of type number
        ' if it is then the default name is used
        FindRes = FindResource(mHangle, dmDwResName, DmResType) ' See if the resource is found in the file
    Else
        FindRes = FindResource(mHangle, dmDwResName, RemoveLeftStr(DmResType, 1)) ' See if the resource is found in the file
    End If
    
    If FindRes = 0 Then GetResourceBlob = "": Exit Function ' Was not found so return false and exit
    
    hObject = LoadResource(mHangle, FindRes) ' Load the resource
    If hObject = 0 Then GetResourceBlob = "": Exit Function ' Return false can't find the resource in the file
    
    Reslock = LockResource(hObject) ' Lock the current resource
    If Reslock = 0 Then GetResourceBlob = "": Exit Function ' exit if we can't lock the resource
    
    ResSize = SizeofResource(mHangle, FindRes) ' Get the current size of the resource returns size retured in bytes
    
    ReDim DataByte(0 To ResSize - 1) ' Resize our databyte array we use this to store the resources data
    
    CopyMemory DataByte(0), ByVal Reslock, ResSize  ' Copy the resources data to our DataByte Byte array
    
    GetResourceBlob = DataByte
    
    Erase DataByte ' clean up
    FreeLibrary mHangle ' free the Library
    FreeResource hObject ' free resource object
    ' clean up some vars
    FindRes = 0
    Reslock = 0
    ResSize = 0
    
End Function

Function DisplayMenu(ResName As String, FrmTempMnuHolder As Form) As Long
Dim mHangle As Long
    ' this function is used to load a menu form a resource file and display it
    DisplayMenu = 1 ' Ok value if eveything was fine
    mHangle = LoadLibrary(ResFileName) ' Get the files libary hangle
    If mHangle = 0 Then DisplayMenu = 0: Exit Function
    MenuHangle = LoadMenu(mHangle, ResName)
    If MenuHangle = 0 Then DisplayMenu = 0: Exit Function ' Send back error code
    FreeLibrary mHangle
End Function

Function DisplayDialog(ResName As String, ResType As String, ParentObj As PictureBox) As Long
Dim mHangle As Long
Dim WndInfo1 As Variant

    DisplayDialog = 1 ' Ok value if eveything was fine
    mHangle = LoadLibrary(ResFileName) ' Get the files libary hangle
    
    If mHangle = 0 Then DisplayDialog = 0: Exit Function ' check that the hangle was found in the resource
    DialogHangle = CreateDialogParam(mHangle, ResName, ParentWnd, 0, 0) ' Create and show the new dialog
    
    WndInfo1 = GetWindowPosition(DialogHangle) ' Get the position height and size of the dialog

    ShowWindow DialogHangle, vbNormalFocus ' show the new dialog window
    
    MoveWindow DialogHangle, 0, 0, WndInfo1(2), WndInfo1(3), 1
    If DialogHangle = 0 Then DisplayDialog = 0: Exit Function
    
    ' Free up time
    FreeLibrary mHangle
    FreeResource hObject
    Erase WndInfo1
End Function

Public Function GetResTypeFromStr(ResType As String) As String
' This is used to return the ID of the resource type string been assigned
    Select Case LCase(ResType)
        Case "cursor"
            GetResTypeFromStr = "#12"
        Case "bitmap"
            GetResTypeFromStr = "#2"
        Case "icon"
            GetResTypeFromStr = "#14"
        Case "menu"
            GetResTypeFromStr = "#4"
        Case "dialog"
            GetResTypeFromStr = "#5"
        Case "string table"
            GetResTypeFromStr = "#6"
        Case "html"
            GetResTypeFromStr = "#23"
        Case "manifest"
            GetResTypeFromStr = "#24"
        Case "hardware-cursor"
            GetResTypeFromStr = "#1"
        Case "hardware-icon"
            GetResTypeFromStr = "#3"
        Case "2110"
            GetResTypeFromStr = "#2110"
        Case "text"
            GetResTypeFromStr = "#TEXT"
        Case "registry"
            GetResTypeFromStr = "#REGISTRY"
        Case "reginst"
            GetResTypeFromStr = "#REGINST"
        Case "avi"
            GetResTypeFromStr = "#AVI"
    End Select
    
End Function

Public Function GetRealTypeName(dwResId As String) As String
' All this code does is return the string name from the dwResId been assigned

    Select Case dwResId
        Case 1
            GetRealTypeName = "Hardware-cursor" 'Hardware dependent cursor
        Case 3
            GetRealTypeName = "Hardware-icon" 'Hardware dependent icon
        Case 12
            GetRealTypeName = "Cursor"  ' Cursor
        Case 2
            GetRealTypeName = "Bitmap"  ' Bitmap image file
        Case 4
            GetRealTypeName = "Menu"    ' Menu resource
        Case 14
            GetRealTypeName = "Icon"    ' Icon Resource
        Case 5, 17
            GetRealTypeName = "Dialog"  ' Dialog resource
        Case 6
            GetRealTypeName = "String Table" ' String table resource
        Case 23
             GetRealTypeName = "Html"  ' Webpage data
        Case 2110
            GetRealTypeName = "2110"   ' May contain webpages, JPEGS,GIF, cursom data
        Case 24
            GetRealTypeName = "Manifest" 'XP theme style
        Case "TEXT"
            GetRealTypeName = "TEXT" ' Text Tips
        Case "REGISTRY"
            GetRealTypeName = "REGISTRY" ' Reg information
        Case "REGINST"
            GetRealTypeName = "REGINST"  ' Same as above
        Case "AVI"
            GetRealTypeName = "AVI"  ' AVI Movie File
    End Select
    
End Function

Public Function AddResType(DwResType As String) As Boolean
' This select case is used to tell the ObjCallMeBack what resource types can be added

    Select Case DwResType
        ' Res types that are allowed in this version
        ' Cursor, Bitmap, Icon, Custom resources
        Case 1
            AddResType = True 'Hardware dependent cursor
        Case 3
            AddResType = True 'Hardware dependent icon.
        Case 2
            AddResType = True
        Case 12
            AddResType = True
        Case 14
            AddResType = True
        Case 4
            AddResType = True
        Case 5
            AddResType = True
        Case 6
            AddResType = True
        Case 23
            AddResType = True
        Case 24
            AddResType = True
        Case 2110
            AddResType = True
        Case "TEXT"
            AddResType = True
        Case "REGISTRY"
            AddResType = True
        Case "REGINST"
            AddResType = True
        Case "AVI"
            AddResType = True
        Case Else
            AddResType = False
    End Select
    
End Function

Public Function EnumLoadResources(lzFilename As String) As Integer
   Dim mHangle As Long, mObj As Long
   
   Static isWorking As Boolean
   
    isWorking = True
    mHangle = LoadLibrary(lzFilename) ' Get the long hangle of the module in the file
    
    If mHangle = 0 Then EnumLoadResources = 0: Exit Function
    ' If we get a zero we must exit
    
    If isWorking Then
        mObj = EnumResourceTypes(mHangle, AddressOf EnumResTypeProc, 0)
        FreeLibrary mHangle ' Free up
        EnumLoadResources = 1 ' Eveything seems to have gone well
        isWorking = False ' Not working anymore
   End If
   
End Function

Private Function EnumResTypeProc(ByVal mHangle As Long, ByVal ResType As Long, ByVal lParam As Long) As Long
   'Dim ResType As String
   Dim isWorking As Boolean

   nCallBack.ObjCallMeBack mHangle, vbNullString, ResType, isWorking
   isWorking = True ' Still working
   ' Get all the resource names for ResType
    'If isWorking Then
        EnumResourceNames mHangle, ByVal ResType, AddressOf EnumResNameProc, lParam
    'End If
    
    EnumResTypeProc = isWorking
End Function
   
Private Function EnumResNameProc(ByVal mHangle As Long, ByVal ResType As Long, ByVal ResName As Long, ByVal lParam As Long) As Long
Dim isWorking As Boolean

   isWorking = True ' We are working
   
   nCallBack.ObjCallMeBack mHangle, ResName, ResType, isWorking  ' Call our ObjCallMeBack Class
   EnumResNameProc = isWorking ' Send back still working state
   
End Function

