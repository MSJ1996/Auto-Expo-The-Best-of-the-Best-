VERSION 5.00
Begin VB.Form MainLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login "
   ClientHeight    =   11520
   ClientLeft      =   -210
   ClientTop       =   -210
   ClientWidth     =   20490
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtPass 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15600
      TabIndex        =   3
      Text            =   "Password"
      ToolTipText     =   "Please Enter a Password"
      Top             =   9120
      Width           =   2775
   End
   Begin VB.TextBox Txtuser 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   15600
      TabIndex        =   2
      Text            =   "Username"
      ToolTipText     =   "Please Enter a User Name"
      Top             =   8400
      Width           =   3375
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Login"
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
      Left            =   16440
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   18360
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   20040
      TabIndex        =   5
      Top             =   11160
      Width           =   375
   End
   Begin VB.Label restpass 
      BackStyle       =   0  'Transparent
      Caption         =   "I forgot my password"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   15960
      TabIndex        =   4
      Top             =   9720
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   15000
      Picture         =   "Login.frx":42D2
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   15000
      Picture         =   "Login.frx":4FB5
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login "
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   15600
      TabIndex        =   0
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -120
      Picture         =   "Login.frx":1B36C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20715
   End
End
Attribute VB_Name = "MainLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public conn As ADODB.Connection

Private Sub cmdlogin_Click()
If Txtuser.Text = "" Then
MsgBox "Username is Empty.", vbInformation
Txtuser.SetFocus
Exit Sub
ElseIf TxtPass.Text = "" Then
MsgBox "Password is Empty"
TxtPass.SetFocus
Exit Sub
Else
Call Login
End If
End Sub

Private Sub Login()
Dim rs As New ADODB.Recordset
rs.Open "SELECT password FROM Logintab WHERE Username = '" & Txtuser.Text & "'", conn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
    MsgBox "Username is Invalid. Please try again.", vbInformation
    Txtuser.SetFocus
Exit Sub
Else
    If TxtPass.Text = rs!Password Then
        MsgBox "Login Successfully"
        Unload Me
        Load MainForm
        MainForm.Show
        MainLogin.Visible = False
    Exit Sub
    Else
        MsgBox "Password is Invalid. Please try again.", vbInformation
        TxtPass.SetFocus
    Exit Sub
    End If
End If
Set rs = Nothing
End Sub


Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Logindb.mdb;persist security info=false"
End Sub

Private Sub Image4_Click()
If Image4.Visible Then
    TxtPass.PasswordChar = ""
    Else
    TxtPass.PasswordChar = "*"
    End If
End Sub

Private Sub Image4_DblClick()
If Image4.Visible Then
    TxtPass.PasswordChar = "*"
    Else
    TxtPass.PasswordChar = ""
    End If
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub restpass_Click()
Load Reset
Reset.Show
End Sub

Private Sub txtpass_Change()
    TxtPass.PasswordChar = "*"
End Sub

Private Sub txtpass_Click()
    TxtPass.Text = ""
End Sub
Private Sub Txtuser_Click()
Txtuser.Text = ""
End Sub
Private Sub Txtuser_lostFocus()
If Txtuser.Text = "" Then
    Txtuser.Text = "Username"
    End If
End Sub
Private Sub txtpass_LostFocus()
    If TxtPass.Text = "" Then
        TxtPass.Text = "Password"
        TxtPass.PasswordChar = ""
    End If
End Sub
