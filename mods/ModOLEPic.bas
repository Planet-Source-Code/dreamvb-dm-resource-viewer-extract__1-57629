Attribute VB_Name = "ModOLEPic"
Private Type BITMAPFILEHEADER ' 14 Bytes Bitmap File Header
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes Bitmap Information Header
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

' This module deals with all the picture resource stuff
' eg Picture converting and Saveing of bitmaps

Public Function DisplayPicture(ByVal ResName As String, ByVal ResType As String, dwPictureType As PictureTypeConstants) As Long
Dim PtrToIcoPicture As Long, PtrToBitmap, dwLnResType As Long
Dim PicData() As Byte

    ' This function deals with loading and displaying Bitmaps, Group Icons, Group Cursors and Hardware Icons and Cursors
    dwLnResType = CLng(Right(ResType, Len(ResType) - 1)) ' Remove the # from the left side and convert to type long
    
    Select Case dwPictureType
        Case vbPicTypeBitmap
            PtrToBitmap = GetPictureHangle(ResName, ResType, vbPicTypeBitmap) ' Get's the hangle of the picture resource
            If PtrToBitmap = 0 Then DisplayPicture = 0: Exit Function 'if hangle is 0 then we return error code 0
            Set ViewPicture = BitmapToPicture(PtrToBitmap, vbPicTypeBitmap) ' Set the picture object with the new picture as a bitmap type
            DisplayPicture = 1 ' Return good code
            Exit Function ' We stop here
        Case vbPicTypeIcon ' hardware Icon or cursor
            PicData = GetResourceBlob(ResName, ResType) ' Get the resource data form the file
            PtrToIcoPicture = CreateIconFromResourceEx(PicData(0), UBound(PicData) + 1, dwLnResType - 1, &H30000, 0, 0, &H1000)
            ' The code above creates an Icon from the data from PicData
            If PtrToPicture = 0 Then DisplayPicture = 0: 'No picture pointer was returned so we send back errror code
            Set ViewPicture = BitmapToIcon(PtrToIcoPicture, vbPicTypeIcon) ' as icon type
            DisplayPicture = 1 ' return ok code
            Exit Function ' We stop here
        Case 12 ' Cursor group
            PtrToBitmap = GetPictureHangle(ResName, ResType, vbPicTypeIcon) ' Get's the hangle of the picture resource as icon type
            If PtrToBitmap = 0 Then DisplayPicture = 0: Exit Function 'if hangle is 0 then we return error code 0
            Set ViewPicture = BitmapToIcon(PtrToBitmap, vbPicTypeIcon) ' Set the picture object with the new picture
            DisplayPicture = 1 ' return ok code
            Exit Function ' We stop here
        Case 14 ' Icon group
            PtrToBitmap = GetPictureHangle(ResName, ResType, vbPicTypeIcon) ' Get's the hangle of the picture resource as icon type
            If PtrToBitmap = 0 Then DisplayPicture = 0: Exit Function 'if hangle is 0 then we return error code 0
            Set ViewPicture = BitmapToIcon(PtrToBitmap, vbPicTypeIcon) ' Set the picture object with the new picture
            DisplayPicture = 1 ' return ok code
            Exit Function ' We stop here
    End Select
         
    ' Clean up vars
    Erase PicData
    dwLnResType = 0
    PtrToBitmap = 0
    PtrToIcoPicture = 0
         
End Function

Public Function SaveBitmapToFile(lzFileBmpFile As String, ResName As String, ResType As String) As Boolean
Dim BmpFHead As BITMAPFILEHEADER ' Bitmap file header
Dim BmpInfoHead As BITMAPINFOHEADER ' Bitmap Information header holdes all the info eg size, color, bitdepth etc
Dim mHangle As Long, lReturned As Long, hObject As Long
Dim FindRes As Long, Reslock As Long, ResSize As Long
Dim DataByte() As Byte
Dim ColTable As Long
Dim nFile As Long

' This function allows to save bitmap files from a resource found in a file.
' As you maybe or not awere when a graphic file is placed into a resource file
' the header information is striped away and the main data is left to be placed into the resource file.
' this is were this little pice of code comes in handy.
'
' Well I hope this code works as I only done this quick I mean I had something
' Like a 5 min read over the Bitmap fileformat let us know if ya find problums
' vbdream2k@yahoo.com

    mHangle = LoadLibrary(ResFileName) ' Get the files libary habgle
    If mHangle = 0 Then SaveBitmapToFile = False: Exit Function ' Return false if we can't find the program hangle

    FindRes = FindResource(mHangle, ResName, ResType) ' See if the resource is found in the file
    If FindRes = 0 Then SaveBitmapToFile = False: Exit Function ' Was not found so return false and exit
    
    hObject = LoadResource(mHangle, FindRes) ' Load the resource
    If hObject = 0 Then SaveBitmapToFile = False: Exit Function ' Return false can't find the resource in the file
    
    Reslock = LockResource(hObject) ' Lock the current resource
    If Reslock = 0 Then SaveBitmapToFile = False: Exit Function ' exit if we can't lock the resource
    
    ResSize = SizeofResource(mHangle, FindRes) ' Get the current size of the resource returns size retured in bytes
    
    ReDim DataByte(0 To ResSize - 1) ' Resize our databyte array we use this to store the resources data
    
    CopyMemory DataByte(0), ByVal Reslock, ResSize      ' Copy the resources data to our DataByte Byte array
    
    CopyMemory BmpInfoHead.biSize, DataByte(0), Len(BmpInfoHead)   ' Copy the bitmaps information header from DataByte
    
    'For people that do not know why I might be adding this other than just saveing
    'from the normal VB picture box is becuase this allow you to save bitmap to the correct colour depth
    
    If BmpInfoHead.biBitCount <= 8 Then ' check if the bitmaps bitcount is lee or equal to 8
        If BmpInfoHead.biClrUsed <> 0 Then ' check if the bitmap color uses is more than zero
            ColTable = BmpInfoHead.biClrUsed ' Srore the colour used in ColTable
        Else
            ColTable = (2 ^ BmpInfoHead.biBitCount) ' we need to set the colour table up Pow 2
        End If
    End If

    BmpFHead.bfType = 19778 ' Bitmap file type
    BmpFHead.bfSize = UBound(DataByte) + 14 ' set the size
    BmpFHead.bfOffBits = Len(BmpFHead) + Len(BmpInfoHead) + ColTable * 4
    
    nFile = FreeFile ' Pointer to free file
    Open lzFileBmpFile For Binary As #nFile
        Put #nFile, , BmpFHead ' Write the new bitmaps header with info collected above
        Put #nFile, , DataByte ' Write the resources bitmap data
    Close #1 ' Close the file
    
    SaveBitmapToFile = True 'Cool we made it return true value everything gone fine
    ' I think a little tidy up is needed
    
    Erase DataByte
    FreeLibrary mHangle
    FreeResource hObject
    ColTable = 0
    hObject = 0
    ResSize = 0
    FindRes = 0
    mHangle = 0
    
End Function

Public Function GetPictureHangle(ResName As String, ResType As String, PicType As PictureTypeConstants) As Long
Dim mHangle As Long, lReturned As Long, hObject As Long
Dim FindRes As Long, hLock As Long, ResSize As Long, dmDwResName As String


    ' This function is used to return the hangle of the pictures data
    
    GetPictureHangle = 1 ' Ok value if eveything went fine
    
    mHangle = LoadLibrary(ResFileName) ' Get the files libary hangle
    
    If mHangle = 0 Then ' Check if we found the hangle of the resource
        GetPictureHangle = 0 ' bad error code
        Exit Function ' stop
    End If

    ' ok seems fine
    
    If IsNumeric(RemoveLeftStr(ResName, 1)) Then
        ' the above is used to check if the resname is of type number
        ' if it is then the default name is used
        FindRes = FindResource(mHangle, ResName, ResType)  ' See if the resource is found in the file
        dmDwResName = ResName
    Else
        ' this is not a vauld number and must be text so we must remove the # sign from the resname
        ' I found this problum while looking at a program written in Borland Delphi
        FindRes = FindResource(mHangle, RemoveLeftStr(ResName, 1), ResType) ' See if the resource is found in the file
        dmDwResName = RemoveLeftStr(ResName, 1)
    End If
    
    If FindRes = 0 Then GetPictureHangle = 0: Exit Function ' Exit if we canot find the resource

    hObject = LoadResource(mHangle, FindRes)   ' load the resource of the file

    If hObject = 0 Then GetPictureHangle = 0: Exit Function ' Exit if we canot find the resource
    
    Select Case PicType
        Case vbPicTypeBitmap ' bitmap picture type
            lReturned = LoadBitmap(mHangle, dmDwResName) ' get the hangle of the bitmap
            If lReturned = 0 Then GetPictureHangle = 0: Exit Function
            GetPictureHangle = lReturned
        Case vbPicTypeIcon ' Icon picture type
            If ResType = "#14" Then ' Check if it is type icon
                lReturned = LoadImage(mHangle, dmDwResName, 1, 0, 0, 0) ' get the hangle of the icon
            Else
                lReturned = LoadImage(mHangle, dmDwResName, 2, 0, 0, 0) ' get the hangle of the cursor
            End If
            If lReturned = 0 Then GetPictureHangle = 0: Exit Function
            GetPictureHangle = lReturned
    End Select
    
    ' Free up resources and vars
    FreeLibrary mHangle
    FreeResource hObject
    mHangle = 0
    lReturned = 0
    FindRes = 0
    hObject = 0
    hLock = 0
    
End Function

Public Function BitmapToPicture(ByVal hBmp As Long, Optional PicType As PictureTypeConstants) As IPicture
' used to convert the Bitmap to a visual basic picture type
    If (hBmp = 0) Then Exit Function
    
    Dim vbPic As Picture, tPicConv As PictDesc, IGuid As Guid
    If PicType = 3 Then End
    
        ' Fill PictDesc structure with necessary parts:
        With tPicConv
            .cbSizeofStruct = Len(tPicConv)
            .PicType = PicType
            .hImage = hBmp
        End With
    
        ' Magic GUID for Bitmap picture
        With IGuid
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        
        ' Create a picture object:
        OleCreatePictureIndirect tPicConv, IGuid, True, vbPic
        Set BitmapToPicture = vbPic ' Return the new picture
    
    
End Function

Public Function BitmapToIcon(ByVal hBmp As Long, Optional PicType As PictureTypeConstants) As IPicture
' used to convert the Bitmap to a visual basic picture type
    If (hBmp = 0) Then Exit Function
    
    Dim vbPic As Picture, tPicConv As PictDesc, IGuid As Guid
    
        ' Fill PictDesc structure with necessary parts:
        With tPicConv
            .cbSizeofStruct = Len(tPicConv)
            .PicType = PicType
            .hImage = hBmp
        End With
    
        ' Magic GUID for cursors
        With IGuid
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        
        ' Create a picture object:
        OleCreatePictureIndirect tPicConv, IGuid, True, vbPic
        Set BitmapToIcon = vbPic ' Return the new picture
        
End Function

