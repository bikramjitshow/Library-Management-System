VERSION 5.00
Begin VB.Form frmfine 
   Caption         =   "Library Loan Info"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmfine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e As Integer


Private Sub cmdbo_Click()
a = CInt(t1.Text)
b = CInt(t2.Text)
c = CInt(t3.Text)
d = CInt(t4.Text)
e = a - b
MsgBox ("return date is" & e)



End Sub
