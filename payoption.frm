VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Chqfrm 
   BorderStyle     =   0  'None
   Caption         =   "Cheque Option"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -210
   ClientWidth     =   20490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker ChqD 
      Height          =   615
      Left            =   10320
      TabIndex        =   17
      Top             =   7080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   114884609
      CurrentDate     =   43440
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      TabIndex        =   3
      Text            =   "Select IFSC Code"
      ToolTipText     =   "Please Select IFSC Code"
      Top             =   3600
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      TabIndex        =   2
      Text            =   "Select Bank Name"
      ToolTipText     =   "Please Select Bank Name"
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox Txtwords 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   6
      Top             =   5640
      Width           =   5175
   End
   Begin VB.CommandButton Submit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Submit and Proceed"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   5175
   End
   Begin VB.TextBox TxtCn 
      BackColor       =   &H000080FF&
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
      Height          =   615
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   7
      ToolTipText     =   "Please Enter the Cheque Number"
      Top             =   6360
      Width           =   5175
   End
   Begin VB.TextBox TxtAno 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   5
      Top             =   4920
      Width           =   5175
   End
   Begin VB.TextBox TxtHol 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   10320
      TabIndex        =   4
      ToolTipText     =   "Please Enter the Cheque Holder Name"
      Top             =   4200
      Width           =   5175
   End
   Begin VB.TextBox Txtpay 
      BackColor       =   &H000080FF&
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
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Text            =   "Auto Expo Pvt Ltd"
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "payoption.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Amount(in Words) :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   5760
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Master"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   7680
      TabIndex        =   14
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Dated :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque no :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   6480
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Amount(in No) :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Holder Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay To :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "payoption.frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Chqfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cn As New ADODB.Connection
Dim conn As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String
Dim amount As String
Dim amount1 As String

Public Sub Load_customer(ByRef customer_data As customer)
TxtHol.Text = customer_data.Name
End Sub

Private Sub Combo1_Click()
Combo2.Clear
If Combo1.Text = "Axis Bank" Then
Combo2.AddItem "UTIB0000037"
ElseIf Combo1.Text = "Baroda Bank" Then
Combo2.AddItem "FDRL0001335"
ElseIf Combo1.Text = "Cosmos Bank" Then
Combo2.AddItem "COSB0000916"
ElseIf Combo1.Text = "Dena Bank" Then
Combo2.AddItem "BKDN0CIRCLE"
ElseIf Combo1.Text = "HDFC Bank" Then
Combo2.AddItem "HDFC0001794"
ElseIf Combo1.Text = "Indusland Bank" Then
Combo2.AddItem "INDB0000843"
ElseIf Combo1.Text = "ICICI Bank" Then
Combo2.AddItem "ICIC0000985"
ElseIf Combo1.Text = "Kotak Mahindra Bank" Then
Combo2.AddItem "KKBK0001767"
ElseIf Combo1.Text = "Oriental Bank" Then
Combo2.AddItem "ORBC0100002"
ElseIf Combo1.Text = "Punjab National Bank" Then
Combo2.AddItem "PUNB0126800"
ElseIf Combo1.Text = "State Bank of India" Then
Combo2.AddItem "SBIN0050918"
ElseIf Combo1.Text = "Saraswat Bank" Then
Combo2.AddItem "SRCB0000001"
ElseIf Combo1.Text = "Union Bank" Then
Combo2.AddItem "UBIN0558389"
ElseIf Combo1.Text = "Yes Bank" Then
Combo2.AddItem "YESB0000657"
Else
End If
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Customers.mdb;persist security info=false"
Combo1.AddItem "Axis Bank"
Combo1.AddItem "Baroda Bank"
Combo1.AddItem "Cosmos Bank"
Combo1.AddItem "Dena Bank"
Combo1.AddItem "HDFC Bank"
Combo1.AddItem "Indusland Bank"
Combo1.AddItem "ICICI Bank"
Combo1.AddItem "Kotak Mahindra Bank"
Combo1.AddItem "Oriental Bank"
Combo1.AddItem "Punjab National Bank"
Combo1.AddItem "State Bank of India"
Combo1.AddItem "Saraswat Bank"
Combo1.AddItem "Union Bank"
Combo1.AddItem "Yes Bank"
End Sub

Private Sub Image2_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub Submit_Click()
Dim ch As New ChqDetails
'If ch.Bank_Name = Combo1.Text Then
ch.Pay_to = TxtPay.Text
ch.Bank_Name = Combo1.Text
ch.IFSC = Combo2.Text
ch.Cheq_Holder_Name = TxtHol.Text
ch.Cheq_Amt_No = TxtAno.Text
ch.Cheq_Amt_Words = Txtwords.Text
ch.Cheq_No = TxtCn.Text
ch.Cheq_Dated = ChqD.Value
Call ch.SaveD
Invoice.Load_payment TxtAno.Text, "Cheque Pay"
Invoice.Show
'Else
'MsgBox "Select Option"

Chqfrm.Visible = False
'End If
End Sub

Public Sub Load_Amount(Price)
    amount = Price
    TxtAno.Text = amount
End Sub

Public Sub Load_amount1(pword)
amount1 = pword
Txtwords.Text = amount1
End Sub

Private Sub TxtCn_LostFocus()
Dim l As Integer
l = Len(TxtCn.Text)
If l < 6 Then
MsgBox ("Cheque Number must be 6 digit")
TxtCn.SetFocus
End If
End Sub

Private Sub TxtCn_Change()
If IsNumeric(TxtCn.Text) = False Then
MsgBox ("Digits Only")
TxtCn.Text = ""
TxtCn.SetFocus
End If
End Sub
