VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Accfrm 
   BorderStyle     =   0  'None
   Caption         =   "Account Details"
   ClientHeight    =   12900
   ClientLeft      =   -120
   ClientTop       =   0
   ClientWidth     =   23040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12900
   ScaleWidth      =   23040
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker Dat 
      Height          =   495
      Left            =   10560
      TabIndex        =   21
      Top             =   7800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      Format          =   122880001
      CurrentDate     =   43411
   End
   Begin VB.TextBox TxtWord 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      TabIndex        =   19
      Top             =   6000
      Width           =   5055
   End
   Begin VB.CommandButton CmdSummit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Submit"
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
      Left            =   10560
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8400
      Width           =   5055
   End
   Begin VB.TextBox TxtDeMob 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      MaxLength       =   10
      TabIndex        =   17
      ToolTipText     =   "Please Enter Depositor Mobile Number"
      Top             =   7200
      Width           =   5055
   End
   Begin VB.TextBox TxtDeName 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      TabIndex        =   16
      ToolTipText     =   "Please Enter Depositor Name"
      Top             =   6600
      Width           =   5055
   End
   Begin VB.TextBox TxtAmtNo 
      BackColor       =   &H0080C0FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      TabIndex        =   15
      Top             =   5400
      Width           =   5055
   End
   Begin VB.TextBox TxtBrName 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   10560
      TabIndex        =   14
      Text            =   "Pune Station Branch"
      Top             =   4800
      Width           =   5055
   End
   Begin VB.TextBox TxtDName 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      TabIndex        =   13
      Text            =   "HDFC Bank Pvt Ltd"
      Top             =   4200
      Width           =   5055
   End
   Begin VB.TextBox TxtCode 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      MaxLength       =   11
      TabIndex        =   12
      Text            =   "ASBWZ187925"
      Top             =   3600
      Width           =   5055
   End
   Begin VB.TextBox TxtAccNo 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      MaxLength       =   15
      TabIndex        =   11
      Text            =   "414163800001638"
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox TxtAccN 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10560
      TabIndex        =   10
      Text            =   "Auto Expo Pvt Ltd"
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      Picture         =   "Accfrm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Amount (in Words) :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Depositer Mob No :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Depositer Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Acc Payment"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   8520
      TabIndex        =   7
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Date :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Branch Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "IFSC Code :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   405
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Amount (in Rs.) :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Bank Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Acc No :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   405
      Left            =   7080
      TabIndex        =   1
      Top             =   3000
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Acc Holder Name :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   450
      Left            =   7080
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   12945
      Left            =   -120
      Picture         =   "Accfrm.frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23145
   End
End
Attribute VB_Name = "Accfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String
Dim amount As String
Dim amount1 As String

Public Sub Load_customer(ByRef customer_data As customer)
TxtDeName.Text = customer_data.Name
TxtDeMob.Text = customer_data.MobileNumber
End Sub

Private Sub CmdSummit_Click()

Invoice.Load_payment TxtAmtNo.Text, "Bank Pay"

Dim An As New AccDetails
An.Acc_Holder_Name = TxtAccN.Text
An.Acc_No = TxtAccNo.Text
An.IFSC_Code = TxtCode.Text
An.Drawn_Name = TxtDName.Text
An.Branch_Name = TxtBrName.Text
An.Amt_no = TxtAmtNo.Text
An.Amt_Words = TxtWord.Text
An.Depositor_Name = TxtDeName.Text
An.Depositor_Mob_No = TxtDeMob.Text
An.Acc_Dat = Dat.Value

Call An.Save
Invoice.Show
Accfrm.Visible = False
End Sub
' Amount in No
Public Sub Load_Amount(Price)
    amount = Price
    TxtAmtNo.Text = amount
End Sub
' Amount in Word
Public Sub Load_amount1(pword)
amount1 = pword
TxtWord.Text = amount1
End Sub

Private Sub Image2_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub TxtAccN_Change()
If IsNumeric(TxtAccN.Text) = True Then
MsgBox ("Text Only")
TxtAccN.Text = ""
TxtAccN.SetFocus
End If
End Sub

Private Sub TxtAccNo_Change()
If IsNumeric(TxtAccNo.Text) = False Then
MsgBox ("Digits Only")
TxtAccNo.Text = ""
TxtAccNo.SetFocus
End If
End Sub

Private Sub TxtAccNo_LostFocus()
Dim l As Integer
l = Len(TxtAccNo.Text)
If l < 15 Then
MsgBox ("Input Valid Number")
TxtAccNo.SetFocus
End If
End Sub

Private Sub TxtAccNo_GotFocus()
If (TxtAccN.Text = "") Then
MsgBox ("Enter a valid Bank Acc Name")
TxtAccN.SetFocus
End If
End Sub

Private Sub TxtBrName_Change()
If IsNumeric(TxtBrName.Text) = True Then
MsgBox ("Text Only")
TxtBrName.Text = ""
TxtBrName.SetFocus
End If
End Sub

Private Sub TxtBrName_GotFocus()
If (TxtDName.Text = "") Then
MsgBox ("Enter a valid Drawer Name")
TxtDName.SetFocus
End If
End Sub

Private Sub TxtCode_Change()
If IsNumeric(TxtCode.Text) = False Then
MsgBox ("Digits Only")
TxtCode.Text = ""
TxtCode.SetFocus
End If
End Sub

Private Sub TxtCode_GotFocus()
If (TxtAccNo.Text = "") Then
MsgBox ("Enter a valid Bank Acc No")
TxtAccNo.SetFocus
End If
End Sub

Private Sub TxtDeMob_LostFocus()
Dim l As Integer
l = Len(TxtDeMob.Text)
If l < 10 Then
MsgBox ("Input Valid Number")
TxtDeMob.SetFocus
End If
End Sub

Private Sub TxtDeMob_Change()
If IsNumeric(TxtDeMob.Text) = False Then
MsgBox ("Digits Only")
TxtDeMob.Text = ""
TxtDeMob.SetFocus
End If
End Sub

Private Sub TxtDeMob_GotFocus()
If (TxtDeName.Text = "") Then
MsgBox ("Enter a valid Depositor Name")
TxtDeName.SetFocus
End If
End Sub

Private Sub TxtDeName_Change()
If IsNumeric(TxtDeName.Text) = True Then
MsgBox ("Text Only")
TxtDeName.Text = ""
TxtDeName.SetFocus
End If
End Sub

Private Sub TxtDeName_GotFocus()
If (TxtBrName.Text = "") Then
MsgBox ("Enter a valid Branch Name")
TxtBrName.SetFocus
End If
End Sub

Private Sub TxtDName_Change()
If IsNumeric(TxtDName.Text) = True Then
MsgBox ("Text Only")
TxtDName.Text = ""
TxtDName.SetFocus
End If
End Sub

Private Sub TxtDName_GotFocus()
If (TxtCode.Text = "") Then
MsgBox ("Enter a valid IFSC Code")
TxtCode.SetFocus
End If
End Sub

