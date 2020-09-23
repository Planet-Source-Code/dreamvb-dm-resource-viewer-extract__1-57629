Attribute VB_Name = "ModDlg2VB"
' Well whats in this mod then
' Ok this was realy an extra bouns to my Resource Extractor what I tryed to do
' is allow you to convert a resources dialog to a Visual Basic form
' Please note that not all forms will work and this was only ment as a test

' at the moment this only picks the follwing controls of dialogs
' Buttons
' Lables
' Textboxes
' Listboxes
' checkboxes
' If anyone knows how to get the style of a window that will be handy I can inprove on this then
' reason is becuase some controls such as group boxes come in as buttons

'anyway that about all I am saying hope you create some forms with it
' and if you can't just remmber it was only a bouns add-on so you lost nothing

' O a little note you may need to ajust the leftness of the controls as there not perfect as yet but not to bad

Private Type WndInfo
    WndCaption As String
    WndHeight As Long
    WndWidth As Long
    WndLeft As Long
    WndTop As Long
End Type

Private WindowInfo As WndInfo
Private StrFinal As String
Private Const GWL_STYLE = (-16)

Private Function BuildFormFooter() As String
Dim StrA As String
    ' This functions builds the bottom of the form
    StrA = vbNewLine
    StrA = StrA & "Attribute VB_Name = " & Chr(34) & "Form1" & Chr(34) & vbNewLine
    StrA = StrA & "Attribute VB_GlobalNameSpace = False" & vbNewLine
    StrA = StrA & "Attribute VB_Creatable = False" & vbNewLine
    StrA = StrA & "Attribute VB_PredeclaredId = True" & vbNewLine
    StrA = StrA & "Attribute VB_Exposed = False" & vbNewLine
    StrA = StrA & "'This Form was generated with DM Resource Extractor"
    BuildFormFooter = StrA
    StrA = ""
End Function

Private Function BuildControls() As String
Dim ClsName As String
Dim clsWnd As Long, WndEx As Long, tWnd() As Long, cName As String
Dim CtrlCount(4) As Integer
Dim CtrInfo As Variant
Dim CrlWidth As Long, CtrHeight As Long, CtrLeft As Long, _
CtrTop As Long, CtrCaption As String
Dim sControlStrBuff As String

On Error Resume Next
    
    cName = String(255, Chr(0)) ' create a buffer for our class name
    If DialogHangle = 0 Then Exit Function ' we exit if we cannot find the resources dialogs hangle

    For i = 1 To 255 ' We just loop 255 in this version of the code
        Erase CtrInfo
        ReDim Preserve tWnd(i) ' Resize the window hangle array we use this to store the window Hwnds
        tWnd(i) = FindWindowEx(DialogHangle, tWnd(i - 1), vbNullString, vbNullString) ' Get the Hwnd of the window and store them
        iRet = GetClassName(tWnd(i), cName, 255)  ' Get the windows classname of the window
        ClsName = Left(cName, iRet) ' trim classname
        
        If tWnd(i) = 0 Then Exit For ' exit if we cannot find a window hangle
        
        CtrlCount(0) = CtrlCount(0) + 1 ' keep a count of the controls
        CtrCaption = GetWndText(tWnd(i)) ' Get the controls caption
        CtrInfo = GetWindowPosition(tWnd(i)) ' Get the windows ifnormation size and position
        CtrLeft = (CtrInfo(0) * Screen.TwipsPerPixelX)  ' Get left position of the control
        CtrTop = (CtrInfo(1) * Screen.TwipsPerPixelY)  ' Get the top position of the control
        CrlWidth = (CtrInfo(2) * Screen.TwipsPerPixelX)  ' get the controls width
        CtrHeight = (CtrInfo(3) * Screen.TwipsPerPixelY) ' get the controls height

        
        Select Case UCase(ClsName)
            Case "BUTTON" ' Command Button Control was found
                If (GetWindowStyle(tWnd(i)) And &H3&) = &H3& Then
                    'The line above checks for checkbox style widnow.
                    ' I am not sure if this is the correct way to check like
                    sControlStrBuff = "Begin VB.CheckBox Check" & CtrlCount(0) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Caption      =   " & Chr(34) & CtrCaption & Chr(34) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Enabled      =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(0)))) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Height       =   " & CtrHeight & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Left         =   " & CtrLeft & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      TabIndex     =   " & "1" & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Top          =   " & CtrTop & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Width        =   " & CrlWidth & vbNewLine
                    sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                    StrFinal = StrFinal & sControlStrBuff
                Else
                    sControlStrBuff = "Begin VB.CommandButton Command" & CtrlCount(0) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Caption      =   " & Chr(34) & CtrCaption & Chr(34) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Enabled      =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(0)))) & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Height       =   " & CtrHeight & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Left         =   " & CtrLeft & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      TabIndex     =   " & "1" & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Top          =   " & CtrTop & vbNewLine
                    sControlStrBuff = sControlStrBuff & "      Width        =   " & CrlWidth & vbNewLine
                    sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                    StrFinal = StrFinal & sControlStrBuff
                    sControlStrBuff = ""
               End If

            Case "STATIC" ' Label Control was found
                CtrlCount(1) = CtrlCount(1) + 1 ' keep a count of the controls
                ' Build a collection of lables
                sControlStrBuff = "Begin VB.Label Label" & CtrlCount(1) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      BackStyle      =   " & "0" & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Caption        =   " & Chr(34) & CtrCaption & Chr(34) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Enabled        =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(1)))) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Height         =   " & CtrHeight & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Left           =   " & CtrLeft & vbNewLine
                sControlStrBuff = sControlStrBuff & "      TabIndex       =   " & "0" & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Top            =   " & CtrTop & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Width          =   " & CrlWidth & vbNewLine
                sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                StrFinal = StrFinal & sControlStrBuff
                sControlStrBuff = ""
            
            Case "EDIT" ' Textbox Control was found
                ' Build collection of text boxes
                CtrlCount(2) = CtrlCount(2) + 1 ' keep a count of the controls
                sControlStrBuff = "Begin VB.TextBox Text" & CtrlCount(2) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Enabled        =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(2)))) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Height         =   " & CtrHeight & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Left           =   " & CtrLeft & vbNewLine
                sControlStrBuff = sControlStrBuff & "      TabIndex       =   " & "0" & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Text           =   " & Chr(34) & Chr(34) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Top            =   " & CtrTop & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Width          =   " & CrlWidth & vbNewLine
                sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                StrFinal = StrFinal & sControlStrBuff
                sControlStrBuff = ""
            Case "COMBOBOX" ' Combo Box Control was found
                ' Build a collection of combo boxes
                CtrlCount(3) = CtrlCount(3) + 1 ' keep a count of the controls
                sControlStrBuff = "Begin VB.ComboBox Combo" & CtrlCount(1) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Enabled        =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(3)))) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Height         =   " & CtrHeight & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Left           =   " & CtrLeft & vbNewLine
                sControlStrBuff = sControlStrBuff & "      TabIndex       =   " & "0" & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Text           =   " & Chr(34) & Chr(34) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Top            =   " & CtrTop & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Width          =   " & CrlWidth & vbNewLine
                sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                StrFinal = StrFinal & sControlStrBuff
                sControlStrBuff = ""
            Case "LISTBOX" ' List box was found
                ' Build a collection of listbxoes
                CtrlCount(4) = CtrlCount(4) + 1 ' keep a count of the controls
                sControlStrBuff = "Begin VB.ListBox List" & CtrlCount(4) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Enabled        =   " & CBool(IsWindowEnabled(tWnd(CtrlCount(3)))) & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Height         =   " & CtrHeight & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Left           =   " & CtrLeft & vbNewLine
                sControlStrBuff = sControlStrBuff & "      TabIndex       =   " & "0" & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Top            =   " & CtrTop & vbNewLine
                sControlStrBuff = sControlStrBuff & "      Width          =   " & CrlWidth & vbNewLine
                sControlStrBuff = sControlStrBuff & "End" & vbNewLine
                StrFinal = StrFinal & sControlStrBuff
                sControlStrBuff = ""
        End Select
    Next
    
    BuildControls = StrFinal & vbNewLine & "End" ' Send the data back makeing sure we add the end keyword first
    'DestroyWindow DialogHangle ' Destroy the dialog window if open
    
    ' Clear up
    sControlStrBuff = ""
    StrFinal = ""
    cName = ""
    ClsName = ""
    CtrCaption = ""
    clsWnd = 0
    WndEx = 0
    CtrLeft = 0
    CtrHeight = 0
    CrlWidth = 0
    Erase tWnd
    Erase CtrlCount

End Function

Private Function GetWindowStyle(WndHand As Long) As Long
    ' This is used to get the style of a window
    GetWindowStyle = GetWindowLong(WndHand, GWL_STYLE)
End Function

Private Function BuildVBFormTop() As String
Dim sData As String, WndInfo As Variant
    ' This is the Top part of the form based on the information we collected from the resource dialog
    
    WndInfo = GetWindowPosition(DialogHangle)
    WindowInfo.WndCaption = GetWndText(DialogHangle)
    ' convert width and height to correct values
    WindowInfo.WndWidth = (WndInfo(2) * Screen.TwipsPerPixelX - 120)
    WindowInfo.WndHeight = (WndInfo(3) * Screen.TwipsPerPixelX - 120)
    WindowInfo.WndLeft = 0 ' We just keep this to zero since it not that inportant
    WindowInfo.WndTop = 0  ' see above
    
    sData = "VERSION 5.00" & vbNewLine
    sData = sData & "Begin VB.Form Form1 " & vbNewLine
    sData = sData & "   Caption = " & Chr(34) & WindowInfo.WndCaption & Chr(34) & vbNewLine
    sData = sData & "   ClientHeight = " & WindowInfo.WndHeight & vbNewLine
    sData = sData & "   ClientLeft = " & "60" & vbNewLine
    sData = sData & "   ClientTop = " & "345" & vbNewLine
    sData = sData & "   ClientWidth = " & WindowInfo.WndWidth & vbNewLine
    sData = sData & "   LinkTopic = " & Chr(34) & WindowInfo.WndCaption & Chr(34) & vbNewLine
    sData = sData & "   ScaleHeight = " & "1" & vbNewLine
    sData = sData & "   ScaleMode = " & "1" & vbNewLine
    sData = sData & "   ScaleWidth = " & "1" & vbNewLine
    sData = sData & "   StartUpPosition = " & "3" & vbNewLine
    
    BuildVBFormTop = sData
    sData = ""
    
End Function

Private Function GetWndText(WndHangle As Long) As String
Dim WndStrLen As Long, WndStrBuff As String
    ' this function is used to return text from a window. eg the caption of a window. or captions of textboxes, lables etc
    WndStrBuff = Space(128) ' create a buffer to hold the data
    WndStrLen = GetWindowText(WndHangle, WndStrBuff, Len(WndStrBuff)) ' get the data
    GetWndText = Left(WndStrBuff, WndStrLen) ' strip away junk and return
End Function

Public Function BuildVBForm() As String
Dim sFormHead As String, FormBody As String
    ' Lets Build up the form
    sFormHead = BuildVBFormTop  ' Lets get the top of the form first
    FormBody = BuildControls
    BuildVBForm = sFormHead & FormBody & BuildFormFooter
End Function

