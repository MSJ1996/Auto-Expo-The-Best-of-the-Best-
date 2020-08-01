VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChqDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private chPay_to As String
Private chBank_Name As String
Private chIFSC As String
Private chCheq_Holder_Name As String
Private chCheq_Amt_No As String
Private chCheq_Amt_Words As String
Private chCheq_No As String
Private chCheq_Dated As String


Public Property Let Pay_to(ChequePay_to As String)
    chPay_to = ChequePay_to
End Property

Public Property Let Bank_Name(ChequeBank_Name As String)
    chBank_Name = ChequeBank_Name
End Property

Public Property Let IFSC(ChequeIFSC As String)
    chIFSC = ChequeIFSC
End Property


Public Property Let Cheq_Holder_Name(ChequeCheq_Holder_Name As String)
    chCheq_Holder_Name = ChequeCheq_Holder_Name
End Property

Public Property Let Cheq_Amt_No(ChequeCheq_Amt_No As String)
    chCheq_Amt_No = ChequeCheq_Amt_No
End Property

Public Property Let Cheq_Amt_Words(ChequeCheq_Amt_Words As String)
    chCheq_Amt_Words = ChequeCheq_Amt_Words
End Property

Public Property Let Cheq_No(ChequeCheq_No As String)
    chCheq_No = ChequeCheq_No
End Property

Public Property Let Cheq_Dated(ChequeCheq_Dated As String)
    chCheq_Dated = ChequeCheq_Dated
End Property

Public Property Get Pay_to() As String
    Pay_to = chPay_to
End Property

Public Property Get Bank_Name() As String
    Bank_Name = chBank_Name
End Property

Public Property Get IFSC() As String
    IFSC = chIFSC
End Property

Public Property Get Cheq_Holder_Name() As String
    Cheq_Holder_Name = chCheq_Holder_Name
End Property

Public Property Get Cheq_Amt_No() As String
    Cheq_Amt_No = chCheq_Amt_No
End Property

Public Property Get Cheq_Amt_Words() As String
    Cheq_Amt_Words = chCheq_Amt_Words
End Property

Public Property Get Cheq_No() As String
    Cheq_No = chCheq_No
End Property

Public Property Get Cheq_Dated() As String
    Cheq_Dated = chCheq_Dated
End Property

Public Sub SaveD()
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String
'connect
 conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Cheque_Info.mdb;persist security info=false"
'query
   s = "INSERT INTO Cheque_Details (Pay_to, Bank_Name, IFSC, Cheq_Holder_Name, Cheq_Amt_No, Cheq_Amt_Words, Cheq_No, Cheq_Dated) VALUES('" & Me.Pay_to & "','" & Me.Bank_Name & "','" & Me.IFSC & "','" & Me.Cheq_Holder_Name & "','" & Me.Cheq_Amt_No & "','" & Me.Cheq_Amt_Words & "','" & Me.Cheq_No & "','" & Me.Cheq_Dated & "')"

   rec.Open s, conn, adOpenDynamic, adLockOptimistic
   conn.Close
   MsgBox "Cheque Details Updated", vbInformation
End Sub

