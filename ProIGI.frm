VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C000&
   Caption         =   "MDIForm1"
   ClientHeight    =   10635
   ClientLeft      =   0
   ClientTop       =   765
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu Main_Form 
      Caption         =   "Main"
      Begin VB.Menu Display_Form 
         Caption         =   "Display"
      End
      Begin VB.Menu Close 
         Caption         =   "End"
      End
   End
   Begin VB.Menu Report_Form 
      Caption         =   "Report"
      Begin VB.Menu Car_Report 
         Caption         =   "Car Report"
      End
      Begin VB.Menu Cus_Report 
         Caption         =   "Cus Report"
      End
      Begin VB.Menu Feedback_Report 
         Caption         =   "FeedBack Report"
      End
   End
   Begin VB.Menu Payment_Frm 
      Caption         =   "Payment"
      Begin VB.Menu Cheque_Report 
         Caption         =   "Cheque Report"
      End
      Begin VB.Menu bank_Report 
         Caption         =   "Bank Report"
      End
   End
   Begin VB.Menu Feedback_Form 
      Caption         =   "Feedback"
   End
   Begin VB.Menu Exit_Form 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bank_Report_Click()
BankReport.Show
End Sub

Private Sub Car_Report_Click()
CarReport.Show
End Sub

Private Sub Cheque_Report_Click()
CheqReport.Show
End Sub


Private Sub Cus_Report_Click()
CusReport.Show
End Sub

Private Sub Display_Form_Click()
MainForm.Show
End Sub

Private Sub Exit_Form_Click()
Load Thankufrm
Thankufrm.Show
MDIForm1.Visible = False
End Sub

Private Sub Feedback_Form_Click()
feedbackfrm.Show
End Sub

Private Sub Feedback_Report_Click()
FeedReport.Show
End Sub
