VERSION 5.00
Begin VB.Form Selfrom 
   BorderStyle     =   0  'None
   Caption         =   "Payment Option"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   19560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Proceed to Pay"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   4815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000080FF&
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14760
      TabIndex        =   3
      Top             =   4560
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000080FF&
      Caption         =   "Bank ACC"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      Picture         =   "payment.frx":0000
      Stretch         =   -1  'True
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Select Payment Mode"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   38.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   38.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   7680
      TabIndex        =   0
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "payment.frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Selfrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim amount As String
Dim amount1 As String
Dim model_id As Integer
Dim model_name As String
Dim category As String
Dim brand As String


Private Sub Command1_Click()
If Option1.Value = True Then
Selfrom.Visible = False
Accfrm.Visible = True
Chqfrm.Visible = False
Else
Selfrom.Visible = False
Accfrm.Visible = False
Chqfrm.Visible = True
End If
End Sub

Private Sub Command2_Click()
Invoice.Show
End Sub

Public Sub Load_data(Price, pword, car_category, car_brand, car_model_id, m_name, ByRef customer_object As customer)
    amount = Price
    amount1 = pword
    category = car_category
    brand = car_brand
    model_id = car_model_id
    model_name = m_name
    
    Load Invoice
    Invoice.Load_data car_model_id, model_name, category, brand, customer_object
    
    Accfrm.Load_customer customer_object
    Chqfrm.Load_customer customer_object
    feedbackfrm.Load_customer customer_object
    
    Load Accfrm
    Accfrm.Load_Amount Price
    Accfrm.Load_amount1 pword
    
    
    Load Chqfrm
    Chqfrm.Load_Amount Price
    Chqfrm.Load_amount1 pword
End Sub

Private Sub Image2_Click()
Load MainForm
MainForm.Show
End Sub
