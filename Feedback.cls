VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private fName As String
Private fSex As String
Private fAge As String
Private fIncome As String
Private fQ1 As String
Private fQ2 As String
Private fQ3 As String
Private fQ4 As String
Private fQ5 As String
Private fQ6 As String
Private fQ7 As String
Private fQ8 As String

Public Property Let Name(FeedbackName As String)
    fName = FeedbackName
End Property

Public Property Let Sex(FeedbackSex As String)
    fSex = FeedbackSex
End Property

Public Property Let Age(FeedbackAge As String)
    fAge = FeedbackAge
End Property

Public Property Let Income(FeedbackIncome As String)
    fIncome = FeedbackIncome
End Property

Public Property Let Q1(FeedbackQ1 As String)
    fQ1 = FeedbackQ1
End Property

Public Property Let Q2(FeedbackQ2 As String)
    fQ2 = FeedbackQ2
End Property

Public Property Let Q3(FeedbackQ3 As String)
    fQ3 = FeedbackQ3
End Property

Public Property Let Q4(FeedbackQ4 As String)
    fQ4 = FeedbackQ4
End Property

Public Property Let Q5(FeedbackQ5 As String)
    fQ5 = FeedbackQ5
End Property

Public Property Let Q6(FeedbackQ6 As String)
    fQ6 = FeedbackQ6
End Property

Public Property Let Q7(FeedbackQ7 As String)
    fQ7 = FeedbackQ7
End Property

Public Property Let Q8(FeedbackQ8 As String)
    fQ8 = FeedbackQ8
End Property


Public Property Get Name() As String
    Name = fName
End Property

Public Property Get Sex() As String
    Sex = fSex
End Property

Public Property Get Age() As String
    Age = fAge
End Property

Public Property Get Income() As String
    Income = fIncome
End Property

Public Property Get Q1() As String
    Q1 = fQ1
End Property

Public Property Get Q2() As String
    Q2 = fQ2
End Property

Public Property Get Q3() As String
    Q3 = fQ3
End Property

Public Property Get Q4() As String
    Q4 = fQ4
End Property

Public Property Get Q5() As String
    Q5 = fQ5
End Property

Public Property Get Q6() As String
    Q6 = fQ6
End Property

Public Property Get Q7() As String
    Q7 = fQ7
End Property

Public Property Get Q8() As String
    Q8 = fQ8
End Property


Public Sub SaveData()
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim s As String
'connect
 conn.Open "Provider=Microsoft.jet.OLEDB.4.0;Data Source =E:\project\assets\Feeback.mdb;persist security info=false"
'query
     s = "INSERT INTO Feed_Details (Name, Sex, Age, Income, Q1, Q2, Q3, Q4,Q5,Q6,Q7,Q8) VALUES('" & Me.Name & "','" & Me.Sex & "','" & Me.Age & "','" & Me.Income & "','" & Me.Q1 & "','" & Me.Q2 & "','" & Me.Q3 & "','" & Me.Q4 & "','" & Me.Q5 & "','" & Me.Q6 & "','" & Me.Q7 & "','" & Me.Q8 & "')"

   rec.Open s, conn, adOpenDynamic, adLockOptimistic
   conn.Close
   MsgBox "Feedback Saved", vbInformation
End Sub

