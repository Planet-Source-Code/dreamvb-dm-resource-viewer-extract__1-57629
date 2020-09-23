VERSION 5.00
Begin VB.Form frmTempMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmTempMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    If MenuHangle <> 0 Then DestroyMenu MenuHangle
    Unload frmTempmenu
End Sub

Private Sub Form_Paint()
    SetMenu hwnd, MenuHangle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTempmenu = Nothing
End Sub
