VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bookingfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DOP 
      Height          =   495
      Left            =   12360
      TabIndex        =   34
      Top             =   7560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      Format          =   115015681
      CurrentDate     =   43440
   End
   Begin MSComCtl2.DTPicker DOB 
      Height          =   495
      Left            =   5880
      TabIndex        =   33
      Top             =   7560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393216
      Format          =   115015681
      CurrentDate     =   43440
   End
   Begin VB.TextBox Txtzipcode 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12360
      MaxLength       =   6
      TabIndex        =   5
      ToolTipText     =   "Please Enter Zip Code"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Txtarea 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Please Enter Area Name"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12360
      TabIndex        =   3
      Text            =   "Select City"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      TabIndex        =   2
      Text            =   "Select State"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   120
      Top             =   11040
   End
   Begin VB.TextBox Txt_CusMail 
      BackColor       =   &H00FFFF80&
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
      Left            =   5880
      TabIndex        =   10
      ToolTipText     =   "Please Enter Customer Mail Id"
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox Txt_DelMid 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      TabIndex        =   11
      ToolTipText     =   "Please Enter Dealer Mail Id"
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox Txt_DelMob 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   12360
      MaxLength       =   10
      TabIndex        =   9
      ToolTipText     =   "Please Enter Dealer Mobile Number"
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Txt_DelNam 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      TabIndex        =   7
      ToolTipText     =   "Please Enter Dealer Name"
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox Brand_Txt 
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
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
      Left            =   12360
      TabIndex        =   29
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox TxtAdd 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      ToolTipText     =   "Please Enter Address"
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox TxtMob 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "Please Enter Mobile Number"
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox TxtCompany 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      TabIndex        =   1
      ToolTipText     =   "Please Enter Company Name"
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      ToolTipText     =   "Please Enter Name"
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Model_Txt 
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
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
      Left            =   5880
      TabIndex        =   28
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00FFFF00&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   19920
      TabIndex        =   32
      Top             =   11040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Picture         =   "Booking .frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Area :- "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   31
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust_Mail :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   30
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_Mail :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   27
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_Mobile :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   26
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal_name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   25
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "DOP :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   24
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Model_id :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Zipcode :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   21
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "City :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   20
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "State :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   18
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Com_Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10440
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "           ""Customer Details""             "
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
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Image Batman 
      Height          =   11520
      Left            =   0
      Picture         =   "Booking .frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20610
   End
End
Attribute VB_Name = "bookingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim s As String
Dim model_id_selected As Integer
Dim brand_selected As String
Dim model_price As String
Dim car_pword As String
Dim car_category As String
Dim model_name As String
Dim cst As New customer

Private Sub CmdSave_Click()
cst.Name = txtname.Text
cst.Company = TxtCompany.Text
cst.MobileNumber = TxtMob.Text
cst.DOB = DOB.Value
cst.Address = TxtAdd.Text
cst.Area = Txtarea.Text
cst.State = Combo1.Text
cst.City = Combo2.Text
cst.Zip = Txtzipcode.Text
cst.model_id = Model_Txt.Text
cst.brand = Brand_Txt.Text
cst.DOP = DOP.Value
cst.Dealer_Name = Txt_DelNam.Text
cst.Dealer_Mob = Txt_DelMob.Text
cst.Dealer_Mid = Txt_DelMid.Text
cst.Cus_Mid = Txt_CusMail.Text
Call cst.SaveData
Load Selfrom
Selfrom.Load_data model_price, car_pword, car_category, brand_selected, model_id_selected, model_name, cst
Selfrom.Show
bookingfrm.Visible = False
End Sub

Public Sub Load_Selected_Data(car_model_id, brand, Price, pword, model, category)
    model_id_selected = car_model_id
    brand_selected = brand
    model_price = Price
    model_name = model
    car_pword = pword
    car_category = category
    
    Model_Txt.Text = model_id_selected
    Brand_Txt.Text = brand_selected
End Sub

Private Sub Combo1_Click()
Combo2.Clear
If Combo1.Text = "Andhra Pradesh" Then
Combo2.AddItem "Anantapur"
Combo2.AddItem "Bhimavaram"
Combo2.AddItem "Eluru"
Combo2.AddItem "Guntur"
Combo2.AddItem "Kakinda"
Combo2.AddItem "Khamman"
Combo2.AddItem "Kurnool"
Combo2.AddItem "Nellore"
Combo2.AddItem "Ongole"
Combo2.AddItem "Rajammudry"
Combo2.AddItem "Srikakulam"
Combo2.AddItem "Tirupati"
Combo2.AddItem "Vijayawada"
Combo2.AddItem "Vishakhapatnam"
ElseIf Combo1.Text = "Andaman" Then
Combo2.AddItem "Port Blair"
ElseIf Combo1.Text = "Arunachal Pradesh" Then
Combo2.AddItem "Itanagar"
ElseIf Combo1.Text = "Assam" Then
Combo2.AddItem "Bongaigaon"
Combo2.AddItem "Dibrugarh"
Combo2.AddItem "Guwahati"
Combo2.AddItem "Jorhat"
Combo2.AddItem "Nagoan"
Combo2.AddItem "North Lakhimpur"
Combo2.AddItem "Sibsagar"
Combo2.AddItem "Silchar"
Combo2.AddItem "Tezpur"
Combo2.AddItem "Tinsukia"
ElseIf Combo1.Text = "Bihar" Then
Combo2.AddItem "Begusarai"
Combo2.AddItem "Bhagalpur"
Combo2.AddItem "Darbhanga"
Combo2.AddItem "Gaya"
Combo2.AddItem "Gopalganu"
Combo2.AddItem "Hajipur"
Combo2.AddItem "Muzaffapur"
Combo2.AddItem "Patna"
Combo2.AddItem "Purnea"
Combo2.AddItem "Sasaram"
ElseIf Combo1.Text = "Chhattisgarh" Then
Combo2.AddItem "Ambikapur"
Combo2.AddItem "Bhilai"
Combo2.AddItem "Bilaspur"
Combo2.AddItem "Durg"
Combo2.AddItem "Jagdalpur"
Combo2.AddItem "Korda"
Combo2.AddItem "Raigarh"
Combo2.AddItem "Raipur"
ElseIf Combo1.Text = "Chandigarh" Then
Combo2.AddItem "Chandigarh"
ElseIf Combo1.Text = "Dadra & Nagar Hevali" Then
Combo2.AddItem "Silvassa"
ElseIf Combo1.Text = "Delhi" Then
Combo2.AddItem "Delhi"
ElseIf Combo1.Text = "Goa" Then
Combo2.AddItem "Goa"
ElseIf Combo1.Text = "Gujarat" Then
Combo2.AddItem "Ahmedabad"
Combo2.AddItem "Amreli"
Combo2.AddItem "Jamnagar"
Combo2.AddItem "Navasari"
Combo2.AddItem "Patan"
Combo2.AddItem "Porbandar"
Combo2.AddItem "Rajkot"
Combo2.AddItem "Surat"
Combo2.AddItem "Vadodara"
Combo2.AddItem "Valsad"
Combo2.AddItem "Vapi"
ElseIf Combo1.Text = "Haryana" Then
Combo2.AddItem "Ambala"
Combo2.AddItem "Bahadurgarh"
Combo2.AddItem "Ballabhgarh"
Combo2.AddItem "Faridabad"
Combo2.AddItem "Gurgaon"
Combo2.AddItem "Itnd"
Combo2.AddItem "Kundli"
Combo2.AddItem "palwal"
Combo2.AddItem "Panipat"
Combo2.AddItem "Sirsa"
Combo2.AddItem "sohna"
Combo2.AddItem "Yamunanagar"
ElseIf Combo1.Text = "Himachal Pradesh" Then
Combo2.AddItem "Hamirpur"
Combo2.AddItem "Mandz"
Combo2.AddItem "Nagrota"
Combo2.AddItem "Shimla"
Combo2.AddItem "Solan"
ElseIf Combo1.Text = "Jammu & Kashmir" Then
Combo2.AddItem "Anantnag"
Combo2.AddItem "Jammu"
Combo2.AddItem "Leh"
Combo2.AddItem "Udhampur"
Combo2.AddItem "Srinagar"
ElseIf Combo1.Text = "Jharkhand" Then
Combo2.AddItem "Bokaro"
Combo2.AddItem "Deoghar"
Combo2.AddItem "Dhanbad"
Combo2.AddItem "Hazaribagh"
Combo2.AddItem "Jamshedpur"
Combo2.AddItem "Ramgarh"
Combo2.AddItem "Ranchi"
ElseIf Combo1.Text = "Karnataka" Then
Combo2.AddItem "Banglore"
Combo2.AddItem "Belgaum"
Combo2.AddItem "Bijapur"
Combo2.AddItem "Gulbarga"
Combo2.AddItem "Hassan"
Combo2.AddItem "Hubali"
Combo2.AddItem "Mangalore"
Combo2.AddItem "Mysore"
Combo2.AddItem "Shimoga"
Combo2.AddItem "Tumkur"
Combo2.AddItem "Udupi"
ElseIf Combo1.Text = "Kerala" Then
Combo2.AddItem "Allepey"
Combo2.AddItem "Cochin"
Combo2.AddItem "Kottayam"
Combo2.AddItem "malappuram"
Combo2.AddItem "Muvattapuzah"
Combo2.AddItem "Palakkad"
Combo2.AddItem "Quilon"
Combo2.AddItem "Thalassery"
Combo2.AddItem "Trivandrum"
Combo2.AddItem "Thirissur"
ElseIf Combo1.Text = "Madhya Pradesh" Then
Combo2.AddItem "Bhopal"
Combo2.AddItem "Dewas"
Combo2.AddItem "Guna"
Combo2.AddItem "Gwalior"
Combo2.AddItem "Indore"
Combo2.AddItem "Jabalpur"
Combo2.AddItem "Ratlam"
Combo2.AddItem "Sagar"
Combo2.AddItem "Stna"
Combo2.AddItem "Shahdol"
Combo2.AddItem "Ujjain"
ElseIf Combo1.Text = "Maharashtra" Then
Combo2.AddItem "Ahmednagar"
Combo2.AddItem "Akola"
Combo2.AddItem "Amravati"
Combo2.AddItem "Aurangabad"
Combo2.AddItem "Baramati"
Combo2.AddItem "Jalgoan"
Combo2.AddItem "Kalyan"
Combo2.AddItem "Kharghar"
Combo2.AddItem "Kolhapur"
Combo2.AddItem "Mumbai"
Combo2.AddItem "Nagpur"
Combo2.AddItem "Nanded"
Combo2.AddItem "Pune"
Combo2.AddItem "Panvel"
Combo2.AddItem "Sangali"
Combo2.AddItem "Sholapur"
Combo2.AddItem "Thane"
Combo2.AddItem "Vasai"
ElseIf Combo1.Text = "Manipur" Then
Combo2.AddItem "Imphaal"
ElseIf Combo1.Text = "Meghalaya" Then
Combo2.AddItem "Shillong"
ElseIf Combo1.Text = "Mizoram" Then
Combo2.AddItem "Aizwal"
ElseIf Combo1.Text = "Nagaland" Then
Combo2.AddItem "Dimapur"
Combo2.AddItem "Kohima"
ElseIf Combo1.Text = "Odisha" Then
Combo2.AddItem "Angul"
Combo2.AddItem "Balasore"
Combo2.AddItem "Berhampur"
Combo2.AddItem "Bhubaneshwar"
Combo2.AddItem "Cuttack"
Combo2.AddItem "Jeypore"
Combo2.AddItem "Rourela"
Combo2.AddItem "Samdhalpur"
ElseIf Combo1.Text = "Pondicherry" Then
Combo2.AddItem "Pondicherry"
ElseIf Combo1.Text = "Punjab" Then
Combo2.AddItem "Amritsar"
Combo2.AddItem "Barnala"
Combo2.AddItem "Bhatinda"
Combo2.AddItem "Jallandar"
Combo2.AddItem "Ludhiana"
Combo2.AddItem "Mohali"
Combo2.AddItem "Pathankot"
Combo2.AddItem "Patiala"
Combo2.AddItem "Sangrur"
Combo2.AddItem "Zirakpur"
ElseIf Combo1.Text = "Rajasthan" Then
Combo2.AddItem "Ajmer"
Combo2.AddItem "Banswara"
Combo2.AddItem "Bharatpur"
Combo2.AddItem "Bhiwadi"
Combo2.AddItem "Chittorgarh"
Combo2.AddItem "Jaipur"
Combo2.AddItem "Jodhpur"
Combo2.AddItem "Kota"
Combo2.AddItem "Pali"
Combo2.AddItem "Udaipur"
ElseIf Combo1.Text = "Sikkim" Then
Combo2.AddItem "Gangtok"
ElseIf Combo1.Text = "Tamil Nadu" Then
Combo2.AddItem "Chennai"
Combo2.AddItem "Coimbatore"
Combo2.AddItem "Madurai"
Combo2.AddItem "Pollachi"
Combo2.AddItem "Salem"
Combo2.AddItem "Tiruneveli"
Combo2.AddItem "Tirupur"
Combo2.AddItem "Tuticorin"
Combo2.AddItem "Vellore"
ElseIf Combo1.Text = "Telangana" Then
Combo2.AddItem "Hyderabad"
Combo2.AddItem "Karimnagar"
Combo2.AddItem "Mahabubnagar"
Combo2.AddItem "Nalgonda"
Combo2.AddItem "Secunderabad"
Combo2.AddItem "Waarangal"
ElseIf Combo1.Text = "Tripura" Then
Combo2.AddItem "Agartala"
ElseIf Combo1.Text = "Uttrakhand" Then
Combo2.AddItem "Dehradun"
Combo2.AddItem "Halpwant"
Combo2.AddItem "Hardwar"
Combo2.AddItem "Kashipur"
Combo2.AddItem "Roorkee"
Combo2.AddItem "Rudrapur"
ElseIf Combo1.Text = "Uttar Pradesh" Then
Combo2.AddItem "Agra"
Combo2.AddItem "Aligarh"
Combo2.AddItem "Allahabad"
Combo2.AddItem "Barelly"
Combo2.AddItem "Faizabad"
Combo2.AddItem "Ghaziabad"
Combo2.AddItem "Gorakhpur"
Combo2.AddItem "Greter Nodia"
Combo2.AddItem "Jhansi"
Combo2.AddItem "Kanpur"
Combo2.AddItem "Luknow"
Combo2.AddItem "Meerut"
Combo2.AddItem "Mirzapur"
Combo2.AddItem "Moradabad"
Combo2.AddItem "Muzzafarinagar"
Combo2.AddItem "Noida"
Combo2.AddItem "Varanasi"
ElseIf Combo1.Text = "West Bengal" Then
Combo2.AddItem "Asansol"
Combo2.AddItem "Barasai"
Combo2.AddItem "Coochbewar"
Combo2.AddItem "Durgapur"
Combo2.AddItem "Howarah"
Combo2.AddItem "Kalyani"
Combo2.AddItem "Kharagar"
Combo2.AddItem "Kolkata"
Combo2.AddItem "Malda"
Combo2.AddItem "Serampore"
Combo2.AddItem "Siliurz"
Else
End If
End Sub

Private Sub Combo1_GotFocus()
If (TxtCompany.Text = "") Then
MsgBox ("Enter valid Company Name")
TxtCompany.SetFocus
End If
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Customers.mdb;persist security info=false"
Combo1.AddItem "Andhra Pradesh"
Combo1.AddItem "Andaman"
Combo1.AddItem "Arunachal Pradesh"
Combo1.AddItem "Assam"
Combo1.AddItem "Bihar"
Combo1.AddItem "Chhattisgarh"
Combo1.AddItem "Chandigarh"
Combo1.AddItem "Dadra & Nagar Hevali"
Combo1.AddItem "Delhi"
Combo1.AddItem "Goa"
Combo1.AddItem "Gujarat"
Combo1.AddItem "Haryana"
Combo1.AddItem "Himachal Pradesh"
Combo1.AddItem "Jammu & Kashmir"
Combo1.AddItem "Jharkhand"
Combo1.AddItem "Karnataka"
Combo1.AddItem "Kerala"
Combo1.AddItem "Madhya Pradesh"
Combo1.AddItem "Maharashtra"
Combo1.AddItem "Manipur"
Combo1.AddItem "Meghalaya"
Combo1.AddItem "Mizoram"
Combo1.AddItem "Nagaland"
Combo1.AddItem "Odisha"
Combo1.AddItem "Pondicherry"
Combo1.AddItem "Punjab"
Combo1.AddItem "Rajasthan"
Combo1.AddItem "Sikkim"
Combo1.AddItem "Tamil Nadu"
Combo1.AddItem "Telangana"
Combo1.AddItem "Tripura"
Combo1.AddItem "Uttrakhand"
Combo1.AddItem "Uttar Pradesh"
Combo1.AddItem "West Bengal"
End Sub

Private Sub Image1_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub Label18_Click()
End
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Right(Label1.Caption, 1) + Left(Label1.Caption, Len(Label1.Caption) - 1)
End Sub

'Private Sub Txt_CusMail_GotFocus()
'If (TxtZip.Text = "") Then
'MsgBox ("Enter valid Zipcode")
'TxtZip.SetFocus
'End If
'End Sub

Private Sub Txt_CusMail_LostFocus()
Dim s2 As String
s2 = Txt_CusMail.Text
Dim intAt, intDot As Integer
intAt = InStr(1, s2, "@", vbTextCompare)
intDot = InStr(intAt + 1, s2, ".", vbTextCompare)
If (intAt = 0) Or (intDot = 0) Or (InStr(intAt + 1, s2, "@")) Or (InStr(intDot + 1, s2, "@")) Then
MsgBox ("Please input '@' and/or '.' in Mail Id")
Txt_DelMob.SetFocus
Else
MsgBox ("Mail Id Validated")
End If
End Sub

Private Sub Txt_CusMail_GotFocus()
If (Txt_DelMob.Text = "") Then
MsgBox ("Enter valid Dealer Mobile No")
Txt_DelMob.SetFocus
End If
End Sub

'Private Sub Txt_DelMid LostFocus()
'If (InStr(str1, "@") < 0) Then
'MsgBox ("mail id valid")
'End If
'End Sub

Private Sub Txt_DelMid_LostFocus()
s = Txt_DelMid.Text
Dim intAt, intDot As Integer
intAt = InStr(1, s, "@", vbTextCompare)
intDot = InStr(intAt + 1, s, ".", vbTextCompare)
If (intAt = 0) Or (intDot = 0) Or (InStr(intAt + 1, s, "@")) Or (InStr(intDot + 1, s, "@")) Then
MsgBox ("Please input '@' and/or '.' in Mail Id")
Txt_DelMid.SetFocus
Else
MsgBox ("Mail Id Validated")
End If
End Sub

Private Sub Txt_DelMid_GotFocus()
If (Txt_CusMail.Text = "") Then
MsgBox ("Enter valid Customer Mail")
Txt_CusMail.SetFocus
End If
End Sub

Private Sub Txt_DelMob_Change()
If IsNumeric(Txt_DelMob.Text) = False Then
MsgBox ("Digits Only")
Txt_DelMob.Text = ""
Txt_DelMob.SetFocus
End If
End Sub

Private Sub Txt_DelMob_LostFocus()
Dim l2 As Integer
l2 = Len(Txt_DelMob.Text)
If l2 < 10 Then
MsgBox ("Dealer Mobile Number Must be 10 digit")
TxtAdd.SetFocus
End If
End Sub

Private Sub Txt_DelMob_GotFocus()
If (TxtAdd.Text = "") Then
MsgBox ("Enter valid Address")
TxtAdd.SetFocus
End If
End Sub

Private Sub Txt_DelNam_Change()
If IsNumeric(Txt_DelNam.Text) = True Then
MsgBox ("Text Only")
Txt_DelNam.Text = ""
Txt_DelNam.SetFocus
End If
End Sub

Private Sub Txt_DelNam_GotFocus()
If (TxtMob.Text = "") Then
MsgBox ("Enter valid Mobile No")
TxtMob.SetFocus
End If
End Sub

Private Sub TxtAdd_GotFocus()
If (Txt_DelNam.Text = "") Then
MsgBox ("Enter valid Dealer Name")
Txt_DelNam.SetFocus
End If
End Sub

'Private Sub Txt_DelNam_GotFocus()
'If (TxtCity.Text = "") Then
'MsgBox ("Enter valid City Name")
'TxtCity.SetFocus
'End If
'End Sub

'Private Sub TxtAdd_GotFocus()
'If (TxtMob.Text = "") Then
'MsgBox ("Enter valid Mobile No")
'TxtMob.SetFocus
'End If
'End Sub

'Private Sub TxtCity_Change()
'If IsNumeric(TxtCity.Text) = True Then
'MsgBox ("Text Only")
'TxtCity.Text = ""
'TxtCity.SetFocus
'End If
'End Sub

'Private Sub TxtCity_GotFocus()
'If (TxtState.Text = "") Then
'MsgBox ("Enter valid State Name")
'TxtState.SetFocus
'End If
'End Sub

Private Sub TxtCompany_Change()
If IsNumeric(TxtCompany.Text) = True Then
MsgBox ("Text Only")
TxtCompany.Text = ""
TxtCompany.SetFocus
End If
End Sub

Private Sub TxtCompany_GotFocus()
If (txtname.Text = "") Then
MsgBox ("Enter valid Name")
txtname.SetFocus
End If
End Sub

Private Sub TxtMob_Change()
If IsNumeric(TxtMob.Text) = False Then
MsgBox ("Digits Only")
TxtMob.Text = ""
TxtMob.SetFocus
End If
End Sub
Private Sub TxtMob_LostFocus()
Dim l As Integer
l = Len(TxtMob.Text)
If l < 10 Then
MsgBox ("Mobile No must be 10 digit")
Txtzipcode.SetFocus
End If
End Sub

Private Sub TxtMob_GotFocus()
If (Txtzipcode.Text = "") Then
MsgBox ("Enter valid Zipcode")
Txtzipcode.SetFocus
End If
End Sub

Private Sub TxtName_Change()
If IsNumeric(txtname.Text) = True Then
MsgBox ("Text Only")
txtname.Text = ""
txtname.SetFocus
End If
End Sub

Private Sub TxtName_LostFocus()
txtname.Text = UCase(txtname.Text)
End Sub

Private Sub Txtzipcode_Change()
If IsNumeric(Txtzipcode.Text) = False Then
MsgBox ("Digits Only")
Txtzipcode.Text = ""
Txtzipcode.SetFocus
End If
End Sub


Private Sub Txtzipcode_LostFocus()
Dim l3 As Integer
l3 = Len(Txtzipcode.Text)
If l3 < 6 Then
MsgBox ("Zipcode must be 6 digit")
Txtzipcode.SetFocus
End If
End Sub

