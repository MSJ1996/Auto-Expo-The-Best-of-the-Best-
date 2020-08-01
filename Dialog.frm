VERSION 5.00
Begin VB.Form Selection 
   BorderStyle     =   0  'None
   Caption         =   "Choice Selection"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -135
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Close 
      BackColor       =   &H00808080&
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
      Height          =   495
      Left            =   19920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11040
      Width           =   615
   End
   Begin VB.CommandButton Search 
      BackColor       =   &H00404080&
      Caption         =   "Search"
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
      Left            =   14400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10080
      Width           =   1335
   End
   Begin VB.ComboBox PriceRange 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   8400
      TabIndex        =   4
      Text            =   "Price Range"
      Top             =   10560
      Width           =   3495
   End
   Begin VB.ComboBox CarCategory 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   8400
      TabIndex        =   3
      Text            =   "Car Category"
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "Dialog.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Range"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   10440
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Choice Selection ! ! !"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "Dialog.frx":2D60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim price_range(2) As Double

Dim conn As New ADODB.Connection
'Dim cars As New ADODB.Connection
Dim carBrands() As Integer

Dim brandNames() As String

Private Sub Close_Click()
Load MainForm
MainForm.Show
Selection.Visible = False
End Sub

Private Sub Form_Load()
    CarCategory.AddItem "Sports", 0
    CarCategory.AddItem "Vintage", 1
    CarCategory.AddItem "Luxury", 2
    CarCategory.AddItem "Hybrid", 3
    CarCategory.AddItem "Evision", 4
    
    PriceRange.AddItem "Bellow 20 lac", 0
    price_range(0) = 2000000
    
    PriceRange.AddItem "Between 30 lac - 70 lac", 1
    price_range(1) = 7000000
    
    PriceRange.AddItem "Above 3 cr", 2
    price_range(2) = 30000000
    
    'Connect to database
    ConnectDatabase "E:\project\assets\cars3.mdb"
       
End Sub

Private Sub Image2_Click()
Load MainForm
MainForm.Show
End Sub

Private Sub Search_Click()
    Dim category As Integer
    Dim range As Integer
    Dim total_cars As Integer
    Dim i As Integer
    Dim filter_query As String
    
    category = CarCategory.ListIndex
    range = PriceRange.ListIndex
    
    '' debug code for empty input
    
    filter_query = " cost > " & price_range(range)
    
    Select Case range
        Case 0
            filter_query = " cost < " & price_range(range)
        Case 1
            filter_query = " cost > " & price_range(0) & " AND cost < " & price_range(1)
        Case 2
            filter_query = " cost > " & price_range(range)
    End Select
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
    'Dim cars As New ADODB.Connection
        cars.Open "SELECT DISTINCT car_id, brand FROM cars WHERE category = " & category & " AND " & filter_query, conn, adUseClient, adLockOptimistic, adCmdText
        
        '"SELECT DISTINCT car_id, brand FROM cars WHERE category = " & category & " AND " & filter_query
       'cars.Open filter_query, conn, adUseClient, adLockOptimistic, adCmdText
     'cars.Open "Select distinct car_id, brand from cars where category =" & category & "AND " & filter_query, conn, adOpenDynamic, adLockOptimistic
    'Append carlist
        
    'Total cars in Database
    
    total_cars = cars.RecordCount
    ReDim carBrands(total_cars) As Integer
    ReDim brandNames(total_cars) As String
    
    
    
    If total_cars > 0 Then
    
        Load MainForm
        
        For i = 0 To total_cars - 1
            'Check for empty value
            MainForm.carList.AddItem cars.Fields(1).Value, i
            carBrands(i) = cars.Fields(0).Value
            brandNames(i) = cars.Fields(1).Value
            cars.MoveNext
        Next i
        
        
        MainForm.loadFilters filter_query, category
        
        MainForm.loadBrands carBrands, brandNames, total_cars
        
        MainForm.Show
    Else
        MsgBox ("No Cars Found")
    End If
    
        'cars.Close

    Selection.Visible = False
    
End Sub

Private Sub ConnectDatabase(ByVal database_path As String)
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =" & database_path & ";persist security info=false"
    'conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\cars3.mdb;persist security info=false"
    'conn.Close
End Sub
