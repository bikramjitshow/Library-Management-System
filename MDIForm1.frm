VERSION 5.00
Begin VB.MDIForm mdilibrary 
   BackColor       =   &H0080C0FF&
   Caption         =   "LIBRARY MANAGEMENT SYSTEM"
   ClientHeight    =   7005
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15075
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":000C
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu New 
      Caption         =   "New"
      Begin VB.Menu Issueinsert 
         Caption         =   "issueinsert"
         Begin VB.Menu Book 
            Caption         =   "Book"
         End
         Begin VB.Menu Student 
            Caption         =   "Student"
         End
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu Student1 
         Caption         =   "Student"
      End
      Begin VB.Menu Book1 
         Caption         =   "Book"
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Admin"
   End
   Begin VB.Menu Query 
      Caption         =   "Query"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
   Begin VB.Menu Issue 
      Caption         =   "Issue"
   End
End
Attribute VB_Name = "mdilibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub book_Click()
frmbadd.Show
End Sub

Private Sub Book1_Click()
frmvb.Show

End Sub


Private Sub insert_Click()
Insert.Show
End Sub

Private Sub Issue_Click()
issuereturn.Show
End Sub

Private Sub student_Click()
frmsadd.Show
End Sub

Private Sub tfc_Click()

End Sub

Private Sub Student1_Click()
frmvs.Show

End Sub
