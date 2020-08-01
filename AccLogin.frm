VERSION 5.00
Begin VB.Form AccLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   7515
   ClientTop       =   4035
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSign 
      BackColor       =   &H00808000&
      Caption         =   "Sign in"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox TxtPass 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      TabIndex        =   1
      Text            =   "Password"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Txtuser 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Text            =   "Username"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   0
      Picture         =   "AccLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "AccLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public conn As ADODB.Connection

Private Sub SignIn()
Dim rs As New ADODB.Recordset
rs.Open "SELECT Password FROM Admin_L WHERE Username = '" & Txtuser.Text & "'", conn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
    MsgBox "Username is Invalid. Please try again.", vbInformation
    Txtuser.SetFocus
Exit Sub
Else
    If TxtPass.Text = rs!Password Then
       Load bookingfrm
       bookingfrm.Show
       Unload Me
        
    Exit Sub
    Else
        MsgBox "Password is Invalid. Please try again.", vbInformation
        TxtPass.SetFocus
    Exit Sub
    End If
End If
Set rs = Nothing
End Sub


Private Sub CmdSign_Click()
If Txtuser.Text = "" Then
MsgBox "Username is Empty.", vbInformation
Txtuser.SetFocus
Exit Sub
ElseIf TxtPass.Text = "" Then
MsgBox "Password is Empty"
TxtPass.SetFocus
Exit Sub
Else
Call SignIn
End If
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\AdminLog.mdb;persist security info=false"
End Sub

Private Sub TxtPass_Change()
    TxtPass.PasswordChar = "*"
End Sub

Private Sub TxtPass_Click()
    TxtPass.Text = ""
End Sub


Private Sub TxtPass_LostFocus()
    If TxtPass.Text = "" Then
        TxtPass.Text = "Password"
        TxtPass.PasswordChar = ""
    End If
End Sub

