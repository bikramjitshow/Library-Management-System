VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmbadd 
   BackColor       =   &H00808000&
   Caption         =   "ADD NEW BOOK"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15675
   FillColor       =   &H00808000&
   ForeColor       =   &H00808000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   15675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdbadd 
      BackColor       =   &H008080FF&
      Caption         =   "ADD BOOK"
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
      Left            =   10200
      MaskColor       =   &H0080FF80&
      MousePointer    =   3  'I-Beam
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox txtbid 
      BackColor       =   &H8000000A&
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
      Left            =   10080
      TabIndex        =   3
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtbname 
      BackColor       =   &H8000000A&
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
      Left            =   10080
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox txtbauthor 
      BackColor       =   &H8000000A&
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
      Left            =   10080
      TabIndex        =   1
      Top             =   3960
      Width           =   4455
   End
   Begin VB.TextBox txtbedition 
      BackColor       =   &H8000000A&
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
      Left            =   10080
      TabIndex        =   0
      Top             =   4800
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc adodcb 
      Height          =   855
      Left            =   3000
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BIKRAM\Desktop\yes i do\library.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\BIKRAM\Desktop\yes i do\library.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "book"
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
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Book Id"
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
      Left            =   7080
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Book Name"
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
      Left            =   7080
      TabIndex        =   7
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Author Name"
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
      Left            =   7080
      TabIndex        =   6
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Book Edition"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Add a New Book"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   6960
      TabIndex        =   4
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmbadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbadd_Click()
adodcb.Refresh
adodcb.Recordset.AddNew
adodcb.Recordset.Fields!bid = txtbid.Text
adodcb.Recordset.Fields!bname = txtbname.Text
adodcb.Recordset.Fields!bauthor = txtbauthor.Text
adodcb.Recordset.Fields!bedition = txtbedition.Text
adodcb.Recordset.Update
MsgBox "yes"

End Sub
