VERSION 5.00
Begin VB.Form LoginS 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -135
   ClientWidth     =   20490
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   72
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sign up "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   72
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   14160
      TabIndex        =   0
      Top             =   9600
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "LoginS.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "LoginS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Load Registration
Registration.Show
End Sub

Private Sub Label2_Click()
Load MainLogin
MainLogin.Show
End Sub

