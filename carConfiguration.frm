VERSION 5.00
Begin VB.Form carConfiguration 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Car Configuration"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   0
      Top             =   11040
   End
   Begin VB.TextBox Txt_words 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   49
      Top             =   10800
      Width           =   7935
   End
   Begin VB.TextBox Txt_Des 
      BackColor       =   &H0080C0FF&
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
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Top             =   9480
      Width           =   9135
   End
   Begin VB.TextBox Txt_Pol 
      BackColor       =   &H0080C0FF&
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
      Left            =   12240
      TabIndex        =   45
      Top             =   10080
      Width           =   2175
   End
   Begin VB.TextBox Txt_Bui 
      BackColor       =   &H0080C0FF&
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
      Left            =   12240
      TabIndex        =   43
      Top             =   9360
      Width           =   2175
   End
   Begin VB.TextBox Txt_Whee 
      BackColor       =   &H0080C0FF&
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
      Height          =   405
      Left            =   17520
      TabIndex        =   40
      Top             =   10200
      Width           =   2655
   End
   Begin VB.TextBox Txt_Dri 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   39
      Top             =   9600
      Width           =   2655
   End
   Begin VB.TextBox Txt_Stw 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   38
      Top             =   9000
      Width           =   2655
   End
   Begin VB.TextBox Txt_Tank 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   37
      Top             =   8400
      Width           =   2655
   End
   Begin VB.TextBox Txt_Size 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   36
      Top             =   7800
      Width           =   2655
   End
   Begin VB.CommandButton CmdNex 
      BackColor       =   &H000080FF&
      Caption         =   "Next"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton CmdPrev 
      BackColor       =   &H000080FF&
      Caption         =   "Previous "
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton CmdFeedb 
      BackColor       =   &H000080FF&
      Caption         =   "Feedback"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton CmdBook 
      BackColor       =   &H000080FF&
      Caption         =   "Book"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton CmdMain 
      BackColor       =   &H000080FF&
      Caption         =   "Main"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Txt_Cost 
      BackColor       =   &H0080C0FF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """Rs."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
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
      Height          =   375
      Left            =   17520
      TabIndex        =   24
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox Txt_FuelType 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   23
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox Txt_Torque 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   22
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox Txt_Abs 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   21
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Txt_Airbags 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   20
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Txt_Speed 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   19
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Txt_Power 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   18
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Txt_Transmission 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   17
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Txt_Engine 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   16
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Txt_Model 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   15
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Txt_Brand 
      BackColor       =   &H0080C0FF&
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
      Height          =   375
      Left            =   17520
      TabIndex        =   14
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Txt_Model_ID 
      BackColor       =   &H0080C0FF&
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
      Height          =   420
      Left            =   17520
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image ImageDisplay 
      Height          =   8055
      Left            =   0
      Picture         =   "carConfiguration.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost In Words :-"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   48
      Top             =   10800
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   47
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Polution Check :-"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   44
      Top             =   10080
      Width           =   2535
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Built Quality :-"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   42
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Certification"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10920
      TabIndex        =   41
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Wheels Cover :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   35
      Top             =   10200
      Width           =   2655
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Driving Mode :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   34
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Steering Wheel :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   33
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Tank (ltr) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   32
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tyre Size :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   31
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost (Rs) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   13
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Type(G/D) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   12
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Torque (NM) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   11
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ABS :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   10
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Airbags :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   9
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed (KPH) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Power (BHP) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   7
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transmission :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Engine (cc) :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Model :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Model_Id :-"
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
      Height          =   375
      Left            =   14640
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "             Configuration             "
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
      Height          =   375
      Left            =   15240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "carConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim model_id As Integer
Dim brand As String
Dim model_price As String
Dim car_pword As String
Dim car_category_list
Dim car_category As String
Dim model_name As String

'Dim myVal As Currency

'Connect with datbase
Dim conn As New ADODB.Connection
Dim car As New ADODB.Recordset
Dim pictures As New ADODB.Recordset

'set picture in Confi
Dim pic_context() As String
Dim display_pic As String
Dim current_loaded_pic_index As Integer

Private Sub CmdBook_Click()
    Load bookingfrm
    bookingfrm.Load_Selected_Data model_id, brand, model_price, car_pword, car_category, model_name
    bookingfrm.Show
    carConfiguration.Visible = False
    MainForm.Visible = False
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdFeedb_Click()
Load feedbackfrm
feedbackfrm.Show
carConfiguration.Visible = False
MainForm.Visible = False
End Sub

Private Sub CmdMain_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub CmdNex_Click()
    If current_loaded_pic_index < UBound(pic_context) - 1 Then
        current_loaded_pic_index = current_loaded_pic_index + 1
        ImageDisplay.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
        'ImageDisplay1.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
    Else
        current_loaded_pic_index = -1
        ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
        'ImageDisplay1.Picture = LoadPicture("E:\project\images\" & display_pic)
    End If
End Sub

Private Sub CmdPrev_Click()
    If current_loaded_pic_index > 0 Then
        current_loaded_pic_index = current_loaded_pic_index - 1
        ImageDisplay.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
        'ImageDisplay1.Picture = LoadPicture("E:\project\images\Extra\" & model_id & "\" & pic_context(current_loaded_pic_index))
    Else
        current_loaded_pic_index = -1
        ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
        'ImageDisplay1.Picture = LoadPicture("E:\project\images\" & display_pic)
    End If
End Sub

Private Sub Form_Load()
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\cars3.mdb;persist security info=false"
End Sub

Public Sub Intilize_form()
    car_category_list = Array("Sports", "Vintage", "Luxury", "Hybrid", "Concept")
    
    If model_id > 0 Then
        query = "SELECT * FROM cars WHERE model_id = " & model_id
        car.Open query, conn, adUseClient, adLockOptimistic, adCmdText
        
        If car.RecordCount > 0 Then
            car_category = car_category_list(car!category)
            
            Txt_Model_ID.Text = car!model_id
            
            brand = car!brand
            Txt_Brand.Text = car!brand
            
            Txt_Model.Text = car!model
            model_name = car!model
            
            Txt_Engine.Text = car!engine
            Txt_Transmission.Text = car!transmission
            Txt_Power.Text = car!power
            Txt_Speed.Text = car!speed
            Txt_Airbags.Text = car!airbags
            Txt_ABS.Text = car!Abs
            Txt_Torque.Text = car!torque
            Txt_FuelType.Text = car!fuel_type
            
            model_price = car!cost
            Txt_Cost.Text = car!cost
          
            car_pword = car!word
            Txt_Words.Text = car!word
        
            Txt_Size.Text = car!Tyre_Size
            Txt_Tank.Text = car!Fuel_tank
            Txt_Stw.Text = car!Steering_Wheel
            Txt_Dri.Text = car!Drive_mode
            Txt_Whee.Text = car!Wheels_Cover
            Txt_Bui.Text = car!Built
            Txt_Pol.Text = car!Polution
            Txt_Des.Text = car!Description
            
            
            'Load images from pictures database
            
            pic_query = "SELECT Pic FROM cars2 WHERE model_id = " & model_id
            pictures.Open pic_query, conn, adUseClient, adLockOptimistic, adCmdText
            
            If pictures.RecordCount > 0 Then
                ReDim pic_context(pictures.RecordCount) As String
                For i = 0 To pictures.RecordCount - 1
                    pic_context(i) = pictures!pic
                    pictures.MoveNext
                Next i
            End If
            
            If Not IsNull(car!display_pic) Then
                display_pic = car!display_pic
                ImageDisplay.Picture = LoadPicture("E:\project\images\" & display_pic)
                'ImageDisplay1.Picture = LoadPicture("E:\project\images\" & display_pic)
                current_loaded_pic_index = -1
            End If
        Else
            MsgBox ("Model not found")
        End If
        
    Else
        MsgBox ("Invalid Model ID")
    End If
    car.Close
    pictures.Close
End Sub
Public Sub Add_model_id(ByVal model As Integer)
    model_id = model
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Right(Label3.Caption, 1) + Left(Label3.Caption, Len(Label3.Caption) - 1)
End Sub
