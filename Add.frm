VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Add 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23040
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   23040
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid Car_Details 
      Height          =   4935
      Left            =   6240
      TabIndex        =   58
      Top             =   5760
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Cardata 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6240
      ScaleHeight     =   270
      ScaleWidth      =   1755
      TabIndex        =   59
      Top             =   10920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FF80FF&
      Caption         =   "EXIT "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   10920
      Width           =   1575
   End
   Begin VB.CommandButton CmdP 
      BackColor       =   &H00FF80FF&
      Caption         =   "PREVIOUS "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   10920
      Width           =   1695
   End
   Begin VB.CommandButton CmdN 
      BackColor       =   &H00FF80FF&
      Caption         =   " NEXT "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   10920
      Width           =   2055
   End
   Begin VB.CommandButton CmdLast 
      BackColor       =   &H00FFC0FF&
      Caption         =   "LAST "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   10920
      Width           =   2055
   End
   Begin VB.CommandButton CmdFirst 
      BackColor       =   &H00FFC0FF&
      Caption         =   "FIRST "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   10920
      Width           =   2055
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00FFC0FF&
      Caption         =   "DELETE "
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
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ADD "
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Txt_Des 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Description"
      DataSource      =   "Cardata"
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
      Left            =   14880
      TabIndex        =   49
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Txt_V 
      BackColor       =   &H00FFC0FF&
      DataField       =   "video"
      DataSource      =   "Cardata"
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
      Left            =   14880
      TabIndex        =   48
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox TxtDisplay 
      BackColor       =   &H00FFC0FF&
      DataField       =   "display_pic"
      DataSource      =   "Cardata"
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
      Left            =   14880
      TabIndex        =   47
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Txt_Words 
      BackColor       =   &H00FFC0FF&
      DataField       =   "word"
      DataSource      =   "Cardata"
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
      Left            =   14880
      TabIndex        =   46
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Txt_Cost 
      BackColor       =   &H00FFC0FF&
      DataField       =   "cost"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   45
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Txt_Polu 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Polution"
      DataSource      =   "Cardata"
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
      Left            =   14880
      TabIndex        =   44
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Txt_Built 
      BackColor       =   &H00FFC0FF&
      DataField       =   "Built"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   43
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Txt_Speed 
      BackColor       =   &H00FFC0FF&
      DataField       =   "speed"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   42
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Txt_Pw 
      BackColor       =   &H00FFC0FF&
      DataField       =   "power"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   41
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Txt_Trans 
      BackColor       =   &H00FFC0FF&
      DataField       =   "transmission"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   40
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Txt_P 
      BackColor       =   &H00FFC0FF&
      DataField       =   "engine"
      DataSource      =   "Cardata"
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
      Left            =   9120
      TabIndex        =   39
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Txt_ABS 
      BackColor       =   &H00FF8080&
      DataField       =   "abs"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   38
      Top             =   10680
      Width           =   2415
   End
   Begin VB.TextBox Txt_WC 
      BackColor       =   &H00FF8080&
      DataField       =   "Wheels_Cover"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   37
      Top             =   9960
      Width           =   2415
   End
   Begin VB.TextBox Txt_DM 
      BackColor       =   &H00FF8080&
      DataField       =   "Drive_mode"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   36
      Top             =   9240
      Width           =   2415
   End
   Begin VB.TextBox Txt_SW 
      BackColor       =   &H00FF8080&
      DataField       =   "Steering_Wheel"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   35
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox Txt_FT 
      BackColor       =   &H00FF8080&
      DataField       =   "Fuel_tank"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   34
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Txt_TS 
      BackColor       =   &H00FF8080&
      DataField       =   "Tyre_Size"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   33
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox Txt_Air 
      BackColor       =   &H00FF8080&
      DataField       =   "Airbags"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   32
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox Txt_T 
      BackColor       =   &H00FFC0FF&
      DataField       =   "torque"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   31
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Txt_Model 
      BackColor       =   &H00FFC0FF&
      DataField       =   "model"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   30
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Txt_Fuel 
      BackColor       =   &H00FFC0FF&
      DataField       =   "fuel_type"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   29
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Txt_C_ID 
      BackColor       =   &H00FF8080&
      DataField       =   "car_id"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Txt_Brand 
      BackColor       =   &H00FF8080&
      DataField       =   "brand"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   27
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Txt_C 
      BackColor       =   &H00FF8080&
      DataField       =   "category"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   26
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Txt_ID 
      BackColor       =   &H00FF8080&
      DataField       =   "model_id"
      DataSource      =   "Cardata"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Car Entry"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   57
      Top             =   240
      Width           =   10335
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Polution"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   24
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Built"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   22
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Vedio"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   21
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Display_Pic"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   20
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost_in_Words"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   19
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Transimission "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Engine"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "ABS"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   10800
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Wheels_Cover"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   10080
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Drive_Mode"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sterring_Wheels"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel_Tank"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tyre_Size"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Airbags"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Torque"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel_Type"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Car_id"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Model_Id"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   27000
      Left            =   0
      Picture         =   "Add.frx":0000
      Top             =   0
      Width           =   43200
   End
End
Attribute VB_Name = "Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim s As String

Private Sub CmdAdd_Click()
Txt_ID.Text = ""
Txt_C.Text = ""
Txt_Brand.Text = ""
Txt_C_ID.Text = ""
Txt_Fuel.Text = ""
Txt_Model.Text = ""
Txt_T.Text = ""
Txt_Air.Text = ""
Txt_TS.Text = ""
Txt_FT.Text = ""
Txt_SW.Text = ""
Txt_DM.Text = ""
Txt_WC.Text = ""
Txt_ABS.Text = ""
Txt_P.Text = ""
Txt_Trans = ""
Txt_Pw.Text = ""
Txt_Speed.Text = ""
Txt_Built.Text = ""
Txt_Cost.Text = ""
Txt_Words.Text = ""
TxtDisplay.Text = ""
Txt_V.Text = ""
Txt_Des.Text = ""
Txt_Polu.Text = ""
Txt_ID.SetFocus
End Sub

Private Sub CmdDelete_Click()
On Error GoTo errmsg
Cardata.Recordset.Delete
MsgBox "Sucessfully Deleted"
Exit Sub
errmsg:
MsgBox "Cannot delete this records"
End Sub

Private Sub CmdExit_Click()
Load MDIForm1
MDIForm1.Show
Add.Visible = False
End Sub

Private Sub CmdFirst_Click()
Cardata.Recordset.MoveFirst
End Sub

Private Sub CmdLast_Click()
Cardata.Recordset.MoveLast
End Sub

Private Sub CmdN_Click()
Cardata.Recordset.MoveNext
End Sub

Private Sub CmdP_Click()
Cardata.Recordset.MovePrevious
End Sub

Private Sub Form_Load()
c.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\cars3.mdb;persist security info=false"
s = "select * from cars"
r.Open s, c, adOpenDynamic, adLockOptimistic
If Not r.BOF And r.EOF Then
Txt_ID.Text = r.Fields(0).Value
Txt_C.Text = r.Fields(1).Value
Txt_Brand.Text = r.Fields(2).Value
Txt_C_ID.Text = r.Fields(3).Value
Txt_Fuel.Text = r.Fields(4).Value
Txt_Model.Text = r.Fields(5).Value
Txt_T.Text = r.Fields(6).Value
Txt_Air.Text = r.Fields(7).Value
Txt_TS.Text = r.Fields(8).Value
Txt_FT.Text = r.Fields(9).Value
Txt_SW.Text = r.Fields(10).Value
Txt_DM.Text = r.Fields(11).Value
Txt_WC.Text = r.Fields(12).Value
Txt_ABS.Text = r.Fields(13).Value
Txt_P.Text = r.Fields(14).Value
Txt_Trans = r.Fields(15).Value
Txt_Pw.Text = r.Fields(16).Value
Txt_Speed.Text = r.Fields(17).Value
Txt_Built.Text = r.Fields(18).Value
Txt_Cost.Text = r.Fields(19).Value
Txt_Words.Text = r.Fields(20).Value
TxtDisplay.Text = r.Fields(21).Value
Txt_V.Text = r.Fields(22).Value
Txt_Des.Text = r.Fields(23).Value
Txt_Polu.Text = r.Fields(24).Value
End If
End Sub
