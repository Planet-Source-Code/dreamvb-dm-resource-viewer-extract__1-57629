VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmmain 
   Caption         =   "DM Resource Extractor Personal"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8985
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   3345
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   2745
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1612
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1964
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2008
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":26AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":30A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":33F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3DEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicBase 
      BorderStyle     =   0  'None
      Height          =   4680
      Left            =   2640
      ScaleHeight     =   4680
      ScaleWidth      =   6180
      TabIndex        =   1
      Top             =   30
      Width           =   6180
      Begin VB.CommandButton cmdBut2 
         Caption         =   "&Close Dialog"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1815
         TabIndex        =   7
         Top             =   4185
         Visible         =   0   'False
         Width           =   1680
      End
      Begin RichTextLib.RichTextBox txtText 
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   705
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   65535
         TextRTF         =   $"frmmain.frx":413C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SHDocVwCtl.WebBrowser WebView 
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   1215
         Visible         =   0   'False
         Width           =   405
         ExtentX         =   714
         ExtentY         =   688
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "#1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   15
         TabIndex        =   2
         Top             =   4185
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.PictureBox picView 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   105
         ScaleHeight     =   390
         ScaleWidth      =   405
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
         Width           =   405
      End
      Begin MediaPlayerCtl.MediaPlayer MMPlayer 
         Height          =   990
         Left            =   90
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   1515
         AudioStream     =   -1
         AutoSize        =   -1  'True
         AutoStart       =   0   'False
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   0   'False
         EnableFullScreenControls=   0   'False
         EnableTracker   =   0   'False
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   -1  'True
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   4575
      Left            =   -15
      TabIndex        =   0
      Top             =   60
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   8070
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImgLst"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10037
            MinWidth        =   2364
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   1575
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   1575
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DM Resource Viewer/Extractor
'
' Hello this is a small project I made for viewing and extrating Bitmaps from appliactions
' Some of you may remmber me submitting it called DM Bitmap Extractor
' well anyway from that version it has changed in a big way.

'Fearures
' View and Extract bitmaps from exe,dll,ocx
' Removed For loop so you can now view evey bitmap in a program or dll
' Added Hardware Icon, Icon, Hardware Cursor and Cursor view and extract
' Note that Icons don't save they only save to bitmaps
' Added new bitmap saver to save bitmaps to there orginal color depth
' Added feature to view dialog resources.
' Added feature to export dialog to a Visual Basic Form ' Note Demo Version
' Added feature to view and extract AVI resources
' Added feature to view and extract Windows XP mainfest files
' Added feature to view and extract resource string tables
' Added feature to view and extract Tips Text
' Added feature to view and extract Reg files
' Added feature to view and Menu resources
' Added feature to view and extact HTML,JPG,CSS,GIF tip outlook has a nise dll to do this msoeres.dll

' Well people that is about it all the code is commeted maybe a little to much but there you go

' Like to say thanks to all the people that voted for my orginal project that started this of
' If you like to use any code in your own project please do so.
' If you can just add my name somewere that all I ask you can use dreamvb or vbdream2k

Private dwResName As String, DwResType As String ' Hold the resname and restype
Private DataBlob() As Byte
Private vDlgProp As Variant
Private sTempFileName As String

Private Sub GlobalCleanUp()
On Error Resume Next
    ' Clean up eveything
    If MMPlayer.PlayState = mpPlaying Then MMPlayer.Stop
    MMPlayer.FileName = ""
    dwResName = ""
    DwResType = ""
    sTempFileName = ""
    ResFileName = ""
    SaveOption = ""
    Erase DataBlob
    Erase vDlgProp
    Set nCallBack = Nothing
    Set ViewPicture = Nothing
    txtText.Text = ""
    WebView.Navigate "about:blank"
    Set picView = Nothing
    CheckForDialogWnd
    HideShowOBj False, False, False, False, False
    If MenuHangle <> 0 Then DestroyMenu MenuHangle
End Sub

Private Sub CheckForDialogWnd()
    If DialogHangle Then
        DestroyWindow DialogHangle
        DialogHangle = 0
    End If
End Sub

Private Sub HideShowOBj(PicViewer As Boolean, TextViewer As Boolean, WebViewer As Boolean, CloseDlgButton As Boolean, MedaPalyer)
    ' Code below hides or shows the control based on there boolean value been passed
    picView.Visible = PicViewer
    txtText.Visible = TextViewer
    WebView.Visible = WebViewer
    MMPlayer.Visible = MedaPalyer
    cmdBut2.Visible = CloseDlgButton
    cmdBut2.Enabled = True
End Sub

Private Sub LastStatus(lzCode As Long)
' some error messages
    If lzCode = 0 Then MsgBox "There was a error while reading the resource." _
    & vbCrLf & "this may be due to the file been compressed with an EXE Packer or similar tool.", vbCritical, frmmain.Caption

End Sub
Function DisplayDialogSave(DlgTitle As String, dlgFilter As String, ResType As Variant, Optional DefFilename As String = "", Optional SaveData As String) As Boolean
Dim iRet As Long
Dim sTmp As String, sOld As String


On Error GoTo CanErr:
    
    DisplayDialogSave = True
    
    With CDLG
        .CancelError = True
        .DialogTitle = "Save " & DlgTitle
        .Filter = dlgFilter
        .FileName = DefFilename
        .ShowSave
        If Len(.FileName) = 0 Then Exit Function
        
        Select Case ResType
            Case "REGINST", "REGISTRY" ' Registery File
                SaveToFile .FileName, SaveData ' Save the Text Tips and it's data
            Case "AVI"
                WriteByteArrayToFile .FileName, DataBlob ' Save the AVI and it's contents
            Case "TEXT" ' Text Tips Resource
                SaveToFile .FileName, SaveData ' Save the Text Tips and it's data
            Case 1 ' Hardware cursor
                SavePicture picView.Image, .FileName
            Case 2 ' resource bitmap
                'SavePicture picView.Image, .FileName
                DisplayDialogSave = SaveBitmapToFile(.FileName, dwResName, DwResType)
            Case 3
                SavePicture picView.Image, .FileName
            Case 5
                SaveToFile .FileName, SaveData  ' save Visual Basic generated form data
            Case 6 ' Resource String Table
                SaveToFile .FileName, SaveData
            Case 12 ' cursor
                SavePicture picView.Image, .FileName
            Case 14 ' icon
                SavePicture picView.Image, .FileName
            Case 24 'Resource manifest file
                SaveToFile .FileName, SaveData  ' save manifest data
            Case 2110, 23
                WriteByteArrayToFile .FileName, DataBlob ' Save the AVI and it's contents
        End Select
        .FileName = ""
    End With
    Exit Function
CanErr:
 If Err Then Err.Clear

End Function

Private Sub Setup()
    tv1.Nodes.Clear ' Clear all the items in treeview
    tv1.Indentation = 20 '  set the indentation to 20
    nCallBack.NodeCnt = 0 ' Reset Node Counter
    nCallBack.ResourcesFound = 0 ' Reset recorces found
    nCallBack.TFORM = Me ' Set up the form to we be using
End Sub

Function HasLoaded(lzFile As String) As Integer
    Setup
    HasLoaded = EnumLoadResources(lzFile)  'This will coolect all the resource types and names in file
End Function

Private Sub cmdBut2_Click()
On Error Resume Next
    If DwResType = "#5" Then ' Dialog type is selected
        ' This will allow us to close a dialog that maybe open
        If DialogHangle <> 0 Then CheckForDialogWnd: cmdBut2.Enabled = False
    End If
   
End Sub

Private Sub cmdsave_Click()
Dim def_SaveName As String, Result As Boolean

    Select Case SaveOption
        Case "#REGINST", "#REGISTRY"
            def_SaveName = StrConv(RemoveLeftStr(dwResName, 1), vbProperCase) ' create a temp filename
            Result = DisplayDialogSave(def_SaveName, "Text Files(*.txt)|*.txt|Reg Files(*.reg)|*.reg|Inf Files(*.inf)|*.inf|", RemoveLeftStr(DwResType, 1), def_SaveName, txtText.Text)

        Case "#2110", "#23" ' Custiom data html,JPEG,CSS etc
            def_SaveName = vDlgProp(2)
            DisplayDialogSave CStr(vDlgProp(0)), CStr(vDlgProp(1)), Int(RemoveLeftStr(DwResType, 1)), def_SaveName
        Case "#AVI" ' AVI Files
            def_SaveName = StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)  ' create a temp filename
            DisplayDialogSave "Save AVI", "AVI Files(*.avi)|*.avi|", RemoveLeftStr(DwResType, 1), def_SaveName
        Case "#TEXT"
            def_SaveName = StrConv(RemoveLeftStr(dwResName, 1), vbProperCase) ' create a temp filename
            Result = DisplayDialogSave(def_SaveName, "Text Files(*.txt)|*.txt|", RemoveLeftStr(DwResType, 1), def_SaveName, txtText.Text)
        Case "#1" ' Hardware cursor
            def_SaveName = RemoveLeftStr(dwResName, 1) ' create a temp filename
            Result = DisplayDialogSave("Bitmap", "Bitmap Files(*.bmp)|*.bmp|", RemoveLeftStr(DwResType, 1), def_SaveName)
            If Not Result Then MsgBox "There was an error while saveing the file.", vbCritical, frmmain.Caption
        Case "#2" ' Save as bitmap
           def_SaveName = RemoveLeftStr(dwResName, 1) ' create a temp filename
           Result = DisplayDialogSave("Bitmap", "Bitmap Files(*.bmp)|*.bmp|", RemoveLeftStr(DwResType, 1), def_SaveName)
           If Not Result Then MsgBox "There was an error while saveing the file.", vbCritical, frmmain.Caption
        Case "#3" ' Hardware Icon
           def_SaveName = RemoveLeftStr(dwResName, 1) ' create a temp filename
           Result = DisplayDialogSave("Bitmap", "Bitmap Files(*.bmp)|*.bmp|", RemoveLeftStr(DwResType, 1), def_SaveName)
           If Not Result Then MsgBox "There was an error while saveing the file.", vbCritical, frmmain.Caption
        Case "#5" ' Shows the Dialog to the user
            def_SaveName = StrConv(RemoveLeftStr(dwResName, 1), vbProperCase) & ".frm" ' create a temp filename
            DisplayDialogSave "Visual Basic Form", "Visual Basic Form Files(*.frm)|*.frm|", Int(RemoveLeftStr(DwResType, 1)), def_SaveName, txtText.Text
        Case "#6" ' save a string table
            def_SaveName = RemoveFileExt(GetFileTitle(ResFileName)) & RemoveLeftStr(dwResName, 1)  ' Build a temp filename
            DisplayDialogSave "Resource Template", "Resource Templates(*.rc)|*.rc|", Int(RemoveLeftStr(DwResType, 1)), def_SaveName, txtText.Text
        Case "#12" ' Cursor
            def_SaveName = RemoveLeftStr(dwResName, 1) ' create a temp filename
            Result = DisplayDialogSave("Bitmap", "Bitmap Files(*.bmp)|*.bmp|", RemoveLeftStr(DwResType, 1), def_SaveName)
            If Not Result Then MsgBox "There was an error while saveing the file.", vbCritical, frmmain.Caption
        Case "#14" ' Icon
           def_SaveName = RemoveLeftStr(dwResName, 1) ' create a temp filename
           Result = DisplayDialogSave("Bitmap", "Bitmap Files(*.bmp)|*.bmp|", RemoveLeftStr(DwResType, 1), def_SaveName)
           If Not Result Then MsgBox "There was an error while saveing the file.", vbCritical, frmmain.Caption

        Case "#24" ' Windows manifest files also known as XML files used to give windows that XP feel
            ' We make a temp name from the resources file name
            def_SaveName = RemoveFileExt(GetFileTitle(ResFileName)) ' Build a temp filename
            DisplayDialogSave "Manifest XP style", "Manifest files(*.manifest)|*.manifest|XML Files(*.xml)|*.xml|", 24, def_SaveName, txtText.Text
    End Select
    def_SaveName = "" ' Clear
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tv1.Height = (frmmain.ScaleHeight - StatusBar1.Height - tv1.Top)
    ' Line aboves resizes the treeview control to forms hieght
    Line1(0).X2 = frmmain.ScaleWidth: Line1(1).X2 = frmmain.ScaleWidth
    ' Line above adds a 3D lines along the top of the form

    PicBase.Width = (frmmain.ScaleWidth - PicBase.Left - 20)
    PicBase.Height = tv1.Height - PicBase.Top
    ' Lines above are used to reize the base picture box that will holder the view below
    If picView.Visible Then
        picView.Left = (PicBase.ScaleWidth - picView.ScaleWidth) \ 2
        picView.Top = (PicBase.ScaleHeight - picView.ScaleHeight) \ 2
        ' Position the view picture box in the center of the base picturebox above
    End If
    
    If WebView.Visible Then
        WebView.Left = 60
        WebView.Top = tv1.Top
        WebView.Width = (PicBase.Width - WebView.Left - 40) ' set Web Viewers width
        WebView.Height = (tv1.Height - PicBase.Top - cmdsave.Height * 1.5) ' set Web Viewers height makeing sure we can see the cmdsave button
    End If
    
    If MMPlayer.Visible Then
        MMPlayer.Left = (PicBase.ScaleWidth - MMPlayer.Width) \ 2
        MMPlayer.Top = (PicBase.ScaleHeight - MMPlayer.Height) \ 2
    End If
    
    If txtText.Visible Then ' check if the txtText is Visible
        txtText.Top = 60 ' set txtText top value
        txtText.Width = (PicBase.Width - txtText.Left - 40) ' set txtText width
        txtText.Height = (tv1.Height - PicBase.Top - cmdsave.Height * 1.5) ' set txtText height makeing sure we can see the cmdsave button
    End If
    
    cmdsave.Top = (PicBase.ScaleHeight - cmdsave.Height) ' Position the save bitmap button
    cmdBut2.Top = cmdsave.Top
End Sub

Private Sub Form_Terminate()
    CheckForDialogWnd ' Closes the resources dialog if is found to be open
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean up before we exit
    RemoveTempFile sTempFileName  ' Delete any temp files left behine
    GlobalCleanUp
    Set ClsMain = Nothing
    Set frmTempMenu = Nothing
    Set frmabout = Nothing
    Set frmmain = Nothing
End Sub


Private Sub MMPlayer_EndOfStream(ByVal Result As Long)
    If cmdBut2.Caption = "&Stop" Then cmdBut2.Caption = "&Play"
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, frmmain 'show this programs about box
End Sub

Private Sub mnuexit_Click()
On Error Resume Next
    Unload frmmain 'unload this program
End Sub

Private Sub mnuopen_Click()
On Error GoTo CanErr:
Dim FileExt As String * 3
Dim Result As Long

    With CDLG
        .CancelError = True ' turn Cancel error on
        .DialogTitle = "Open File" ' update dialogs title
        .Filter = "Appliaction Files(*.exe)|*.exe|Dynamic Link Files(*.dll)|*.dll|Ocx Files(*.ocx)|*.ocx|Icon Libraries(*.icl)|*.icl|" ' Update filetypes
        .ShowOpen ' show open dialog
        If Len(.FileName) = 0 Then Exit Sub ' stop if filename len = 0
        FileExt = LCase(GetFileExt(.FileName)) ' get the files ext
        
        If Not (FileExt = "exe" Or FileExt = "dll" Or FileExt = "ocx" Or FileExt = "icl") Then ' check for vaild file types
            MsgBox "Unsopported file type" _
            & vbCrLf & vbCrLf & "Only files types of ocx,dll,exe,icl are allowed", vbCritical, frmmain.Caption
            ' user selected other file type than supported for display error
            Exit Sub ' stop
        End If
        
        Result = HasLoaded(.FileName) ' returns 1 if loaded otherwise 0
        If Result = 0 Then ' if not loaded then stop
            GlobalCleanUp
            MsgBox "There was an error while loading the file", vbCritical, frmmain.Caption
            StatusBar1.Panels(1).Text = "" ' clear statusbar text
            StatusBar1.Panels(2).Text = "" ' clear statusbar text
            Exit Sub
        Else
            GlobalCleanUp
            cmdsave.Visible = False 'hide the save button
            StatusBar1.Panels(1).Text = .FileName ' update status bar with the filename
            ResFileName = .FileName ' store the filename for later use.
        End If
    End With
    
    Exit Sub ' stop
    
CanErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Dim CanShow As Boolean, Result As Long, ipos As Integer, iResStrData As String, sCaption As String

    If tv1.Nodes.Count = 0 Then Exit Sub ' do nothing if nothing is in the treeview
    If Node.Children <> 0 Then
        StatusBar1.Panels(2).Text = Node.Children ' update the status bar panel 2 with the number of child nodes
        dwResName = ""
        DwResType = ""
        MMPlayer.FileName = ""
        HideShowOBj False, False, False, False, False ' Hide all the controls
    Else
        StatusBar1.Panels(2).Text = "ID :" & Node.Key ' update the status bar panel 2 with the number of child key
    End If
    
    DwResType = GetResTypeFromStr(Node.Parent) ' Get the resources type
    ipos = InStr(1, Node.Key, ":", vbTextCompare)
    If ipos > 0 Then dwResName = Mid(Node.Key, 1, ipos - 1) ' Extract the resource ID
    
    sCaption = "&Save " & StrConv(GetRealTypeName(Val(Right(DwResType, Len(DwResType) - 1))), vbProperCase)
    ' Line above builds the save buttons caption base on the resource type selected
    SaveOption = DwResType ' Lets us know in the save button which resource type to save to
    
    CheckForDialogWnd ' Destroy the dialog resource if found
    Unload frmTempMenu ' unload the menu form

    
    Select Case DwResType
        Case "#AVI" ' Movie type resource
            Erase DataBlob ' Erase any data in the byte array
            DataBlob() = GetResourceBlob(dwResName, DwResType)
            If Not UBound(DataBlob) <> 0 Then Exit Sub
            
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button

            RemoveTempFile sTempFileName ' Delete the temp file
            sTempFileName = GetTempPathA() & RemoveLeftStr(dwResName, 1) & ".avi" ' Build the temp anme
            WriteByteArrayToFile sTempFileName, DataBlob() ' Write the resources contents to a temp file
            MMPlayer.FileName = sTempFileName
            Result = 1

            HideShowOBj False, False, False, False, True ' Hide or show controls

        Case "#REGINST" ' Registry resource data
            iResStrData = StrConv(GetResourceBlob(dwResName, DwResType), vbUnicode)
            Result = Len(iResStrData) <> 0
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button
            HideShowOBj False, True, False, False, False ' Hide or show controls
            txtText.Text = iResStrData ' Update the textbox with the Tips Text
            iResStrData = ""
        Case "#REGISTRY" ' Same as above
            iResStrData = StrConv(GetResourceBlob(dwResName, DwResType), vbUnicode)
            Result = Len(iResStrData) <> 0
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            HideShowOBj False, True, False, False, False ' Hide or show controls
            txtText.Text = iResStrData ' Update the textbox with the Tips Text
            iResStrData = ""
        Case "#TEXT" ' TipsTexts ' this is something I came accross in winamp thought be handy to add in
            iResStrData = StrConv(GetResourceBlob(dwResName, DwResType), vbUnicode)
            Result = Len(iResStrData) <> 0
            cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            HideShowOBj False, True, False, False, False ' Hide or show controls
            txtText.Text = iResStrData ' Update the textbox with the Tips Text
            iResStrData = ""
        Case "#1", "#3" ' hardware Cursor or Icons
            HideShowOBj True, False, False, False, False ' Hide or show controls
            Result = DisplayPicture(dwResName, DwResType, vbPicTypeIcon)
            picView.Picture = ViewPicture
         Case "#2"
            HideShowOBj True, False, False, False, False ' Hide or show controls
            Result = DisplayPicture(dwResName, DwResType, vbPicTypeBitmap)
            picView.Picture = ViewPicture
        Case "#12" 'cursor group
            HideShowOBj True, False, False, False, False ' Hide or show controls
            Result = DisplayPicture(dwResName, DwResType, 12)
            picView.Picture = ViewPicture
        Case "#4" ' Resource Menu
            Result = DisplayMenu(dwResName, frmTempMenu)
            frmTempMenu.Visible = Result <> 0 ' show menu if result is greator than 0
        Case "#5" ' Load and display a dialog from a resource file
            cmdBut2.Caption = "&Close Dialog" ' Update buttons to caption
            HideShowOBj False, True, False, True, False ' Hide or show controls
            Result = DisplayDialog(dwResName, DwResType, picView)
            If Result <> 0 Then txtText.Text = BuildVBForm
        Case "#14" ' icon group
            HideShowOBj True, False, False, False, False ' Hide or show controls
            Result = DisplayPicture(dwResName, DwResType, 14)
            picView.Picture = ViewPicture
        Case "#24" ' Manifest or other custom data
            HideShowOBj False, True, False, False, False ' Hide or show controls
            iResStrData = StrConv(GetResourceBlob(dwResName, DwResType), vbUnicode) ' get the resources string data
            Result = LenB(iResStrData) <> 0 ' Return a boolean value based of the len of the incomming string zero means at error
            txtText.Text = iResStrData ' update the textbox with the data found
            iResStrData = "" ' Clean string buffer
        Case "#6" ' String Table
            HideShowOBj False, True, False, False, False ' Hide or show controls
            txtText.Text = GetStringTable(dwResName, DwResType)
            ' Load and assign string table to the TextBox
            Result = 1
        Case "#2110", "#23" ' This seems to be a resource for HTML, JPEG, GIFS I found this in Outlook express
            Erase DataBlob
            DataBlob() = GetResourceBlob(dwResName, DwResType)  ' Get resource data stream
            
            cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button
            sCaption = "&Save " & StrConv(RemoveLeftStr(dwResName, 1), vbProperCase)
            ' Set the save buttons caption to the resource name
            vDlgProp = GetDescFromFileExt(RemoveLeftStr(dwResName, 1)) ' Get information for the save dialog box
            RemoveTempFile sTempFileName ' Delete the temp file
            sTempFileName = GetTempPathA() & RemoveLeftStr(dwResName, 1) ' Build the temp anme
            WriteByteArrayToFile sTempFileName, DataBlob() ' Write the resources contents to a temp file
            Result = UBound(DataBlob) <> 0
            
            If Result Then
                Select Case GetFileExt(sTempFileName) ' we now need to work out what con trols to show and hide
                    Case "GIF", "JPEG", "JPG", "JPE", "JFIF", "BMP", "DIB", "EMF", "WMF" ' Supported picture types allowed
                        HideShowOBj True, False, False, False, False ' Hide or show controls
                        picView.Picture = LoadPicture(sTempFileName) ' Update the picture box with picture
                    Case "HTM", "HTML", "CSS", "ASP", "SHTML", "SHTM", "XML", "TXT", "PHP", "PHP3", "PHP4", "PHP5" ' Supported web documents allowed
                        HideShowOBj False, False, True, False, False ' Hide or show controls
                        WebView.Navigate sTempFileName ' Load the webpage data
                End Select
            End If
    End Select
    
    cmdsave.Caption = sCaption
    cmdsave.Width = TextWidth(sCaption) * 1.5 ' resize the Save button
    cmdBut2.Left = cmdsave.Width + 100
    CanShow = Node.Parent <> "" ' check if a node is selected
    
    If CanShow Then
        Form_Resize ' Resize and update the form
        cmdsave.Visible = True ' show the cmdsave button
        LastStatus Result ' returns any error message returned from the value of result
    Else
        cmdsave.Visible = False ' hide the cmdsave button
        cmdBut2.Visible = False
        Set picView.Picture = Nothing ' destroy picview picture
        txtText.Visible = False
    End If
End Sub

