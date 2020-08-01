VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   120
      Left            =   0
      Top             =   11040
   End
   Begin VB.CommandButton CmdSel 
      BackColor       =   &H00FFFF00&
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
      Height          =   615
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10800
      Width           =   1335
   End
   Begin VB.CommandButton ResetFilters 
      BackColor       =   &H00C0C000&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10800
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0C000&
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
      Height          =   615
      Left            =   19800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10800
      Width           =   615
   End
   Begin VB.ComboBox brandType 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Text            =   "Select Category"
      Top             =   240
      Width           =   2865
   End
   Begin VB.ListBox carModel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   885
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   2865
   End
   Begin VB.CommandButton loadcarBtn 
      BackColor       =   &H00C0C000&
      Caption         =   "View Configuration"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   11040
      UseMaskColor    =   -1  'True
      Width           =   4305
   End
   Begin VB.ListBox carList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   885
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2865
   End
   Begin VB.CommandButton playVdoControl 
      BackColor       =   &H00FFFF00&
      Cancel          =   -1  'True
      Caption         =   "Play Adv"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10800
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "  ""The Best of the Best""   "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   42
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1215
      Left            =   8640
      TabIndex        =   7
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Car Model "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image BGpic 
      Height          =   11535
      Left            =   0
      Picture         =   "mainform.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21135
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim conn As New ADODB.Connection
Dim currentVdo As String
Dim currentCar As String
Dim currentCategory As Integer
Dim currentModels() As Integer
Dim current_model As Integer

'Store car brand ids

Dim carBrands() As Integer

Dim filter_query As String

Private Sub CmdClose_Click()
End
End Sub

Private Sub CmdSel_Click()
Load Selection
Selection.Show
End Sub

Private Sub Form_Load()
    filter_query = ""
    
    Me.WindowState = 2
    BGpic.Top = Me.Top
    BGpic.Left = Me.Left
    BGpic.Height = Me.Height
    BGpic.Width = Me.Width
    
    currentVdo = ""
    loadcarBtn.Enabled = False
    
    
    brandType.AddItem "Sports", 0
    brandType.AddItem "Vintage", 1
    brandType.AddItem "Luxury", 2
    brandType.AddItem "Hybrid", 3
    brandType.AddItem "Concept", 4
    
    
    'Connect to database
    ConnectDatabase "E:\project\assets\cars3.mdb"
    
    playVdoControl.Enabled = False
End Sub

' Brand Type: sports, vintage, ...
Private Sub brandType_Click()
    Index = brandType.ListIndex
    
    'Sets the current category on clipboard
    currentCategory = Index
    
    'Load Carlist
     Dim cars As New ADODB.Recordset
     query = "SELECT DISTINCT car_id, brand FROM cars WHERE category = " & Index
     
     cars.Open query, conn, adUseClient, adLockOptimistic, adCmdText
     

    'Append carlist
        
    'Total cars in Database
    total_cars = cars.RecordCount
    ReDim carBrands(total_cars) As Integer
    
    carList.Clear
    carList.Refresh
    
    If total_cars > 0 Then
    
        For i = 0 To total_cars - 1
            'Check for empty value
            carList.AddItem cars.Fields(1).Value, i
            carBrands(i) = cars.Fields(0).Value
            cars.MoveNext
        Next i
    End If
    
    'carList.AddItem cars.Fields(1)
    cars.Close
End Sub

Private Sub carModel_Click()
    loadcarBtn.Enabled = True
    Index = carModel.ListIndex
    model_id = currentModels(Index)
    current_model = model_id
    
    Dim car As New ADODB.Recordset
    
    q = "SELECT * FROM cars WHERE model_id = " & model_id
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    If IsNull(car!display_pic) Then
        MsgBox ("No Pics found for this car")
    Else
        BGpic.Picture = LoadPicture("E:\project\images\" & car!display_pic)
    End If
    
    currentCar = car!brand & "(" & car!model & ")"
    
    If IsNull(car!video) Then
    
    MsgBox ("No Vdo found for this car")
    
       playVdoControl.Enabled = False
    Else
        playVdoControl.Enabled = True
        currentVdo = car!video
    End If
    
    car.Close
End Sub

Private Sub Form_Resize()
    BGpic.Height = Me.Height
    BGpic.Width = Me.Width
End Sub

Private Sub frontView_Click()
    Dim pic As New ADODB.Recordset
        
    Index = carModel.ListIndex
    model_id = currentModels(Index)
    
    Dim car As New ADODB.Recordset
    q = "SELECT pic FROM pictures WHERE model_id = " & model_id & " AND view = 'front'"
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    
    If IsNull(car!pic) Then
        MsgBox ("No Pictures found for this car")
    Else
        BGpic.Picture = LoadPicture("E:\project\images\" & car!pic)
    End If
    
End Sub

Private Sub Label2_Click()
Load AdminLogin
AdminLogin.Show
MainForm.Visible = False
End Sub

Private Sub loadcarBtn_Click()
    Load carConfiguration
    carConfiguration.Add_model_id current_model
    carConfiguration.Intilize_form
    carConfiguration.Show
End Sub

Private Sub carList_Click()

    Index = carList.ListIndex
    
    Dim car As New ADODB.Recordset
    
    
    q = "SELECT model_id, model FROM cars WHERE car_id = " & carBrands(Index) & " AND category = " & currentCategory
    
    If Not filter_query = "" Then
        q = q & " AND " & filter_query
    End If
    
    car.Open q, conn, adUseClient, adLockOptimistic, adCmdText
    total_models = car.RecordCount
    
    ReDim currentModels(total_models) As Integer
    carModel.Clear
    carModel.Refresh
    
    For i = 0 To total_models - 1
        currentModels(i) = car.Fields(0).Value
        carModel.AddItem car.Fields(1).Value, i
        car.MoveNext
    Next i
    car.Close
End Sub



'************************
' Database subrouteines (Functions)

Private Sub ConnectDatabase(ByVal database_path As String)
    conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =" & database_path & ";persist security info=false"
End Sub

Private Sub playVdoControl_Click()
    Load vdoPlayerDlg
    vdoPlayerDlg.Show
    vdoPlayerDlg.Play_Video "E:\project\vdo\" & currentVdo, currentCar
End Sub

Public Sub loadBrands(ByRef brands() As Integer, ByRef names() As String, total As Integer)
    ReDim carBrands(total) As Integer
    carBrands = brands
    
    carList.Clear
    carList.Refresh
    
    For i = 0 To total - 1
        carList.AddItem names(i), i
    Next i
End Sub

Public Sub loadFilters(ByRef query As String, category As Integer)
    filter_query = query
    brandType.ListIndex = category
End Sub

Private Sub ResetFilters_Click()
    filter_query = ""
    Unload MainForm
    Set MainForm = Nothing
    Load MainForm
    MainForm.Show
End Sub
