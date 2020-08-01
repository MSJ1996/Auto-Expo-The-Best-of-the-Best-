VERSION 5.00
Begin VB.Form WelcomeScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Welcome Screen"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -405
   ClientWidth     =   20490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   105
      Left            =   480
      Top             =   11040
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   11040
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "  Welcomes You !!!!!   "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   11400
      TabIndex        =   3
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Expo"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   13560
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Best"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   16200
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The Best Of "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   975
      Left            =   11640
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Image SplashImg 
      Height          =   11505
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "WelcomeScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SplashImg.Width = Me.Width
    SplashImg.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
    Unload Me
    Load LoginS
    LoginS.Show
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Right(Label5.Caption, 1) + Left(Label5.Caption, Len(Label5.Caption) - 1)
End Sub
