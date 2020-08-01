VERSION 5.00
Begin VB.Form feedbackfrm 
   BorderStyle     =   0  'None
   Caption         =   "Feedbackfrm"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -210
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo10 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   11
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   10320
      Width           =   3975
   End
   Begin VB.ComboBox Combo9 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   10
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   9720
      Width           =   3975
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   9
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   9120
      Width           =   3975
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   8
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   8520
      Width           =   3975
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   7
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   7920
      Width           =   3975
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   6
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   7320
      Width           =   3975
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16320
      TabIndex        =   5
      Text            =   "Select Your Choice"
      ToolTipText     =   "Please Select Your Choice"
      Top             =   6720
      Width           =   3975
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   18240
      TabIndex        =   4
      Text            =   "Choice"
      ToolTipText     =   "Please Select Your Annual Income"
      Top             =   4920
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16920
      TabIndex        =   2
      Text            =   "Select Gender"
      ToolTipText     =   "Please Select Gender"
      Top             =   3720
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
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
      Left            =   18240
      TabIndex        =   3
      Text            =   "Select Age"
      ToolTipText     =   "Please Select Your Age group"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   110
      Left            =   120
      Top             =   11040
   End
   Begin VB.TextBox TxtQ8 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   14880
      TabIndex        =   12
      ToolTipText     =   "Please give a Suggestion"
      Top             =   10920
      Width           =   3255
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00C0C0C0&
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
      Left            =   16920
      TabIndex        =   1
      ToolTipText     =   "Please Enter Your Name"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H00FFC0C0&
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19680
      TabIndex        =   29
      Top             =   10920
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   0
      Picture         =   "feedback.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   18000
      Picture         =   "feedback.frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "             Feedback             "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   28
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "8-Any room for improvements/Suggestions"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   8400
      TabIndex        =   27
      Top             =   10920
      Width           =   6375
   End
   Begin VB.Label Sex 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   15120
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   15120
      TabIndex        =   25
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1-The Reception"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   0
      Left            =   13560
      TabIndex        =   24
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "5-Sharing your experience with others"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   10200
      TabIndex        =   23
      Top             =   9120
      Width           =   5895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "7-Our services throughtout the expo"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Index           =   0
      Left            =   10440
      TabIndex        =   22
      Top             =   10320
      Width           =   5535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "6-Attendance provided to you by Sales Person"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      Top             =   9720
      Width           =   6855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4-Vehicle prices and payment policy "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   10560
      TabIndex        =   20
      Top             =   8520
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "3-Fullfillness of our commitments"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   11040
      TabIndex        =   19
      Top             =   7920
      Width           =   5175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "2-Consideration of your time "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   11760
      TabIndex        =   18
      Top             =   7320
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Rate our Team on the basis of :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   10560
      TabIndex        =   17
      Top             =   6120
      Width           =   6255
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Annual Income :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   15120
      TabIndex        =   16
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Age :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   15120
      TabIndex        =   15
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Please tell us a bit about  you ..... ! ! !"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"feedback.frx":5482
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   2415
      Left            =   7440
      TabIndex        =   0
      Top             =   3360
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   -240
      Picture         =   "feedback.frx":5561
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20805
   End
End
Attribute VB_Name = "feedbackfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String

Public Sub Load_customer(ByRef customer_data As customer)
txtname.Text = customer_data.Name
End Sub

Private Sub CmdSubmit_Click()
Dim fee As New Feedback
fee.Name = txtname.Text
fee.Sex = Combo2.Text
fee.Age = Combo1.Text
fee.Income = Combo3.Text
fee.Q1 = Combo4.Text
fee.Q2 = Combo5.Text
fee.Q3 = Combo6.Text
fee.Q4 = Combo7.Text
fee.Q5 = Combo8.Text
fee.Q6 = Combo9.Text
fee.Q7 = Combo10.Text
fee.Q8 = TxtQ8.Text
Call fee.SaveData
Load MainForm
MainForm.Show
feedbackfrm.Visible = False
feedbackfrm.Visible = False
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Feeback.mdb;persist security info=false"

'Age Section
Combo1.AddItem "18-23"
Combo1.AddItem "24-32"
Combo1.AddItem "33-47"
Combo1.AddItem "50-61"
Combo1.AddItem "62-71"
Combo1.AddItem "72-81"
Combo1.AddItem "82-101"
Combo1.AddItem "101-110"
'Gender Section
Combo2.AddItem "Male"
Combo2.AddItem "Female"
Combo2.AddItem "Other"
'Salary Section
Combo3.AddItem "<1 Lk"
Combo3.AddItem "10-20 Lk"
Combo3.AddItem "30-45 Lk"
Combo3.AddItem "50-70 Lk"
Combo3.AddItem ">1.5 Cr"
'Rate Our Services Section
Combo4.AddItem "Poor"
Combo4.AddItem "Satisfactory"
Combo4.AddItem "Moderate"
Combo4.AddItem "Good"
Combo4.AddItem "Excellent"
Combo4.AddItem "Superb"
Combo4.AddItem "Need To Improve"

Combo5.AddItem "Poor"
Combo5.AddItem "Satisfactory"
Combo5.AddItem "Moderate"
Combo5.AddItem "Good"
Combo5.AddItem "Excellent"
Combo5.AddItem "Superb"
Combo5.AddItem "Need To Improve"

Combo6.AddItem "Poor"
Combo6.AddItem "Satisfactory"
Combo6.AddItem "Moderate"
Combo6.AddItem "Good"
Combo6.AddItem "Excellent"
Combo6.AddItem "Superb"
Combo6.AddItem "Need To Improve"

Combo7.AddItem "Poor"
Combo7.AddItem "Satisfactory"
Combo7.AddItem "Moderate"
Combo7.AddItem "Good"
Combo7.AddItem "Excellent"
Combo7.AddItem "Superb"
Combo7.AddItem "Need To Improve"

Combo8.AddItem "Yes"
Combo8.AddItem "No"

Combo9.AddItem "Poor"
Combo9.AddItem "Satisfactory"
Combo9.AddItem "Moderate"
Combo9.AddItem "Good"
Combo9.AddItem "Excellent"
Combo9.AddItem "Superb"
Combo9.AddItem "Need To Improve"

Combo10.AddItem "Poor"
Combo10.AddItem "Satisfactory"
Combo10.AddItem "Moderate"
Combo10.AddItem "Good"
Combo10.AddItem "Excellent"
Combo10.AddItem "Superb"
Combo10.AddItem "Need To Improve"
End Sub

Private Sub Image3_Click()
Load MainForm
MainForm.Show
End Sub


Private Sub Label1_Click()
 Load Thankufrm
 Thankufrm.Show
 feedbackfrm.Visible = False
End Sub

Private Sub Timer1_Timer()
Label16.Caption = Right(Label16.Caption, 1) + Left(Label16.Caption, Len(Label16.Caption) - 1)
End Sub

