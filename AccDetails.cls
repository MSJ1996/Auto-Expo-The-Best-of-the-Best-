VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private AnAcc_Holder_Name As String
Private AnAcc_No As String
Private AnIFSC_Code As String
Private AnDrawn_Name As String
Private AnBranch_Name As String
Private AnAmt_no As String
Private AnAmt_Words As String
Private AnDepositor_Name As String
Private AnDepositor_Mob_No As String
Private AnAcc_Dat As String

Public Property Let Acc_Holder_Name(AccountAcc_Holder_Name As String)
    AnAcc_Holder_Name = AccountAcc_Holder_Name
End Property

Public Property Let Acc_No(AccountAcc_No As String)
    AnAcc_No = AccountAcc_No
End Property

Public Property Let IFSC_Code(AccountIFSC_Code As String)
    AnIFSC_Code = AccountIFSC_Code
End Property

Public Property Let Drawn_Name(AccountDrawn_Name As String)
    AnDrawn_Name = AccountDrawn_Name
End Property

Public Property Let Branch_Name(AccountBranch_Name As String)
    AnBranch_Name = AccountBranch_Name
End Property

Public Property Let Amt_no(AccountAmt_no As String)
    AnAmt_no = AccountAmt_no
End Property

Public Property Let Amt_Words(AccountAnAmt_Words As String)
    AnAmt_Words = AccountAnAmt_Words
End Property

Public Property Let Depositor_Name(Accountepositor_Name As String)
    AnDepositor_Name = Accountepositor_Name
End Property

Public Property Let Depositor_Mob_No(AccountDepositor_Mob_No As String)
    AnDepositor_Mob_No = AccountDepositor_Mob_No
End Property

Public Property Let Acc_Dat(AccountDat As String)
    AnAcc_Dat = AccountDat
End Property

Public Property Get Acc_Holder_Name() As String
    Acc_Holder_Name = AnAcc_Holder_Name
End Property

Public Property Get Acc_No() As String
    Acc_No = AnAcc_No
End Property

Public Property Get IFSC_Code() As String
    IFSC_Code = AnIFSC_Code
End Property

Public Property Get Drawn_Name() As String
    Drawn_Name = AnDrawn_Name
End Property

Public Property Get Branch_Name() As String
    Branch_Name = AnBranch_Name
End Property

Public Property Get Amt_no() As String
    Amt_no = AnAmt_no
End Property

Public Property Get Amt_Words() As String
    Amt_Words = AnAmt_Words
End Property

Public Property Get Depositor_Name() As String
    Depositor_Name = AnDepositor_Name
End Property

Public Property Get Depositor_Mob_No() As String
    Depositor_Mob_No = AnDepositor_Mob_No
End Property

Public Property Get Acc_Dat() As String
    Acc_Dat = AnAcc_Dat
End Property

Public Sub Save()
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String
'connect
 conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Acc_Info.mdb;persist security info=false"
'query
   s = "INSERT INTO Acc_Details (Acc_Holder_Name, Acc_no, IFSC_Code, Drawn_Name, Branch_Name, Amt_no, Amt_Words, Depositor_Name, Depositor_Mob_No, Acc_Dat) VALUES('" & Me.Acc_Holder_Name & "','" & Me.Acc_No & "','" & Me.IFSC_Code & "','" & Me.Drawn_Name & "','" & Me.Branch_Name & "','" & Me.Amt_no & "','" & Me.Amt_Words & "','" & Me.Depositor_Name & "','" & Me.Depositor_Mob_No & "','" & Me.Acc_Dat & "')"

   rec.Open s, conn, adOpenDynamic, adLockOptimistic
   conn.Close
   MsgBox "Account Details Updated", vbInformation
End Sub


