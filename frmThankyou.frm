VERSION 5.00
Begin VB.Form frmThankyou 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6555
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmThankyou.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4500
      Left            =   0
      Top             =   6120
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmThankyou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
'Unload Me
    'login.Show
End Sub
