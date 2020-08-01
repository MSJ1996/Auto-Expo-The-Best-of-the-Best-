VERSION 5.00
Begin VB.Form AdminLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Please Enter Password"
      Top             =   6480
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "LOGIN ADMIN USER"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   3855
   End
   Begin VB.TextBox Txtuser 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   8760
      TabIndex        =   1
      ToolTipText     =   "Please Enter Username"
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox Txtmail 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   8760
      TabIndex        =   0
      ToolTipText     =   "Please Enter Admin Mail Id"
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Image pass2 
      Height          =   495
      Left            =   11880
      Picture         =   "Admin Login.frx":0000
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   6
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   19920
      TabIndex        =   5
      Top             =   11040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   10200
      Width           =   3135
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   12975
      Left            =   0
      Picture         =   "Admin Login.frx":42D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "AdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public conn As ADODB.Connection
Dim s As String

Private Sub Login()
Dim rs As New ADODB.Recordset
rs.Open "SELECT password FROM ALogin WHERE Username = '" & Txtuser(1).Text & "'", conn, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
    MsgBox "Username is Invalid. Please try again.", vbInformation
    Txtuser(1).SetFocus
Exit Sub
Else
    If TxtPass.Text = rs!Password Then
        MsgBox "Login Sucessfully"
        Unload Me
        Load MDIForm1
        MDIForm1.Show
        AdminLogin.Visible = False
    Exit Sub
    Else
        MsgBox "Password is Invalid. Please try again.", vbInformation
        TxtPass.SetFocus
    Exit Sub
    End If
End If
Set rs = Nothing
End Sub


Private Sub Command1_Click()
If Txtuser(1).Text = "" Then
MsgBox "Username is Empty.", vbInformation
Txtuser(1).SetFocus
Exit Sub
ElseIf TxtPass.Text = "" Then
MsgBox "Password is Empty"
TxtPass.SetFocus
Exit Sub
Else
Call Login
End If
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\AdminLogin.mdb;persist security info=false"
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub pass2_Click()
If pass2.Visible Then
    TxtPass.PasswordChar = ""
    Else
    TxtPass.PasswordChar = "*"
    End If
End Sub

Private Sub pass2_DblClick()
If pass2.Visible Then
    TxtPass.PasswordChar = "*"
    Else
    TxtPass.PasswordChar = ""
    End If
End Sub

Private Sub Txtmail_LostFocus(Index As Integer)
s = Txtmail(0).Text
Dim intAt, intDot As Integer
intAt = InStr(1, s, "@", vbTextCompare)
intDot = InStr(intAt + 1, s, ".", vbTextCompare)
If (intAt = 0) Or (intDot = 0) Or (InStr(intAt + 1, s, "@")) Or (InStr(intDot + 1, s, "@")) Then
MsgBox ("Please input '@' and/or '.' in Mail Id")
Txtmail(0).SetFocus
Else
MsgBox ("Mail Id Validated")
End If
End Sub
