VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsadd 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD STUDENT"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14265
   Icon            =   "frmsadd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14265
   Begin VB.CommandButton resetsadd 
      BackColor       =   &H008080FF&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adodcs 
      Height          =   735
      Left            =   9960
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cmbsyear 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
   End
   Begin VB.ComboBox cmbsstream 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox txtsname 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton cmdsadd 
      BackColor       =   &H008080FF&
      Caption         =   "ADD STUDENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtsid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Student Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Student Stream"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Student Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmsadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdsadd_Click()
adodcs.Refresh
adodcs.Recordset.AddNew
adodcs.Recordset.Fields!sid = txtsid.Text
adodcs.Recordset.Fields!sname = txtsname.Text
adodcs.Recordset.Fields!sstream = cmbsstream.Text
adodcs.Recordset.Fields!syear = cmbsyear.Text
adodcs.Recordset.Update
MsgBox "yes"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
cmbsstream.AddItem "BCA"
cmbsstream.AddItem "BBA"
cmbsstream.AddItem "MM"
cmbsstream.AddItem "HN"
cmbsstream.Text = cmbsstream.List(0)
cmbsyear.AddItem "1st year"
cmbsyear.AddItem "2nd year"
cmbsyear.AddItem "3rd year"
cmbsyear.Text = cmbsyear.List(0)


End Sub

Private Sub resetsadd_Click()
If resetsadd.Index = 1 Then
frmsadd.Show
End If
End Sub
