VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private h_Form As Form
Public NodeCnt As Integer ' Holds the number of nodes in a treeview
Public ResourcesFound As Integer
Public CanAdd As Boolean

Function TreeViewIndex(ResType As String, Optional ResName As String) As Integer
' All this function does is returns an index for the image to use in the treeviw
' this is based of the type of resource found

    Select Case ResType
        Case "AVI" ' Movie resource
            TreeViewIndex = 17
        Case "TEXT" ' Text Tips
            TreeViewIndex = 16
        Case "REGISTRY", "REGINST" ' Registry file
            TreeViewIndex = 11
        Case 1, 12 'cursor
            TreeViewIndex = 4
        Case 2 'bitmap
            TreeViewIndex = 3
        Case 4 'menu
            TreeViewIndex = 6
        Case 3, 14 'icon
            TreeViewIndex = 5
        Case 5 'dialog
            TreeViewIndex = 7
        Case 6 'string
            TreeViewIndex = 8
        Case 23 'html
             TreeViewIndex = 10
        Case 24 ' Manifest file or other custom resource types
            TreeViewIndex = 9
        Case 2110
            Select Case UCase(GetFileExt(ResName))
                Case "JPG"
                    TreeViewIndex = 13
                Case "GIF"
                    TreeViewIndex = 14
                Case "HTM", "HTML"
                    TreeViewIndex = 10
                Case "CSS"
                    TreeViewIndex = 15
                Case Else
                    TreeViewIndex = 12
            End Select
        Case Else
            TreeViewIndex = 12
    End Select
    
End Function

Public Sub ObjCallMeBack(ByVal mHangle As Long, ByVal ResName As String, ByVal ResType As String, isWorking As Boolean)
' I use this function to add all my code
Dim MyTreeView As TreeView
Dim lResType As String, lResName As String
On Error Resume Next

    If (CLng(ResType) And &HFFFF0000) = 0 Then 'is it equal to zero
        lResType = ResType ' Yes so we keep it the same
    Else
        lResType = PtrToStr(CLng(ResType)) ' Ok here we need to get the restype name from the pointer
    End If
    
    ' same as above just using the resource name
    If (CLng(ResName) And &HFFFF0000) = 0 Then
        lResName = ResName
    Else
        lResName = PtrToStr(CLng(ResName))
    End If
    
    Set MyTreeView = TFORM.tv1 'Set MyTreeView as the forms treeview
    
    If AddResType(lResType) And Not isWorking Then ' Add the top level resource type folders
        MyTreeView.Nodes.Add , tvwFirst, "#" & lResType, GetRealTypeName(lResType), 1, 2
        ' Line above adds the resource type and it's real resource name to the treeview
        NodeCnt = MyTreeView.Nodes.Count ' Keep a count of how many notes in the treeview we have
        MyTreeView.Nodes(NodeCnt).Bold = True
        MyTreeView.Nodes(NodeCnt).ForeColor = vbBlue
    End If
    
    ResourcesFound = ResourcesFound + 1 ' Keep track of how many resources we have
    
    If AddResType(lResType) And Len(lResName) <> 0 Then
        MyTreeView.Nodes.Add NodeCnt, tvwChild, "#" & lResName & ":" & ResourcesFound, "Resource [" & lResName & "]", TreeViewIndex(lResType, lResName), TreeViewIndex(lResType, lResName)
        'line above adds the node to the treeviww
    End If

End Sub

Public Property Get TFORM() As Form
    Set TFORM = h_Form
End Property

Public Property Let TFORM(ByVal vNewValue As Form)
    Set h_Form = vNewValue
End Property

Private Function PtrToStr(StrPtr As Long) As String
    Dim StrPtrLen As Long
  ' Ok this is a function that converts a pointer in memory to a string
  ' note the comments may be correct or not I not realy played with pointer till now
  ' main reason because some resource names at a pointer were pointing to there string name.
  ' such as htm
  
    StrPtrLen = lstrlenPtr(StrPtr) ' Get the length of the pointer
    PtrToStr = Space(StrPtrLen) ' Create space for our new string
    
    CopyMemoryToStr PtrToStr, ByVal StrPtr, StrPtrLen ' Copy pointer to String
End Function
