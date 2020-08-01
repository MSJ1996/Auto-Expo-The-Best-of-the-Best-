VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cName As String
Private cMobileNumber As String
Private cAddress As String
Private cCompany As String
Private cArea As String
Private cState As String
Private cCity As String
Private cZip As String
Private cDOB As String
Private cModel_id As String
Private cBrand As String
Private cDOP As String
Private cDealer_Name As String
Private cDealer_Mob As String
Private cDealer_Mid As String
Private cCus_Mid As String

Public Property Let Name(customerName As String)
    cName = customerName
End Property

Public Property Let MobileNumber(customerContact As String)
    cMobileNumber = customerContact
End Property

Public Property Let Address(customerAddress As String)
    cAddress = customerAddress
End Property

Public Property Let Company(customerCompany As String)
    cCompany = customerCompany
End Property

Public Property Let Area(customerArea As String)
    cArea = customerArea
End Property

Public Property Let State(customerState As String)
    cState = customerState
End Property

Public Property Let City(customerCity As String)
    cCity = customerCity
End Property

Public Property Let Zip(customerZip As String)
    cZip = customerZip
End Property

Public Property Let DOB(customerDOB As String)
    cDOB = customerDOB
End Property

Public Property Let model_id(customerModel_id As String)
    cModel_id = customerModel_id
End Property

Public Property Let brand(customerBrand As String)
    cBrand = customerBrand
End Property

Public Property Let DOP(customerDOP As String)
    cDOP = customerDOP
End Property
Public Property Let Dealer_Name(customerDealer_Name As String)
    cDealer_Name = customerDealer_Name
End Property

Public Property Let Dealer_Mob(customerDealer_Mob As String)
    cDealer_Mob = customerDealer_Mob
End Property

Public Property Let Dealer_Mid(customerDealer_Mid As String)
    cDealer_Mid = customerDealer_Mid
End Property

Public Property Let Cus_Mid(customerCus_Mid As String)
    cCus_Mid = customerCus_Mid
End Property



Public Property Get Name() As String
    Name = cName
End Property

Public Property Get MobileNumber() As String
    MobileNumber = cMobileNumber
End Property

Public Property Get Address() As String
    Address = cAddress
End Property

Public Property Get Company() As String
    Company = cCompany
End Property

Public Property Get Area() As String
    Area = cArea
End Property

Public Property Get State() As String
    State = cState
End Property

Public Property Get City() As String
    City = cCity
End Property

Public Property Get Zip() As String
    Zip = cZip
End Property

Public Property Get DOB() As String
    DOB = cDOB
End Property

Public Property Get model_id() As String
    model_id = cModel_id
End Property

Public Property Get brand() As String
    brand = cBrand
End Property

Public Property Get DOP() As String
    DOP = cDOP
End Property

Public Property Get Dealer_Name() As String
    Dealer_Name = cDealer_Name
End Property

Public Property Get Dealer_Mob() As String
    Dealer_Mob = cDealer_Mob
End Property

Public Property Get Dealer_Mid() As String
    Dealer_Mid = cDealer_Mid
End Property

Public Property Get Cus_Mid() As String
    Cus_Mid = cCus_Mid
End Property

Public Sub SaveData()
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String
'connect
 conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Customers.mdb;persist security info=false"
'query
     s = "INSERT INTO cus_details (name, company, contact, DOB, address, area, state, city, pincode, Model_id, Brand, DOP, Deal_name, Deal_mobile, Deal_Mail, Cus_Mail) VALUES('" & Me.Name & "','" & Me.Company & "','" & Me.MobileNumber & "','" & Me.DOB & "','" & Me.Address & "','" & Me.Area & "','" & Me.State & "','" & Me.City & "','" & Me.Zip & "','" & Me.model_id & "','" & Me.brand & "','" & Me.DOP & "','" & Me.Dealer_Name & "','" & Me.Dealer_Mob & "','" & Me.Dealer_Mid & "','" & Me.Cus_Mid & "')"
     rec.Open s, conn, adOpenDynamic, adLockOptimistic
     conn.Close
     MsgBox "Customer Saved", vbInformation
End Sub
