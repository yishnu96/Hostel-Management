VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form register 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton backbtn 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   17
      Top             =   8040
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc registerado 
      Height          =   330
      Left            =   9600
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\VB PROJECT\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\VB PROJECT\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
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
   Begin VB.CommandButton resetbtn 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   16
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton regbtn 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox txtphone 
      DataField       =   "contact"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox txtadd 
      DataField       =   "address"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   6120
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      DataField       =   "password"
      DataSource      =   "registerado"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtuser 
      DataField       =   "username"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtclass 
      DataField       =   "class"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtname 
      DataField       =   "name"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtroll 
      DataField       =   "rollno"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "PHONE NO:"
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "ADDRESS:"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "PASSWORD:"
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "USER NAME:"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "CLASS:"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "NAME :"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "ROLL NO:"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backbtn_Click()
welcome.Show
register.Hide
End Sub

Private Sub Form_Load()
registerado.Recordset.AddNew
End Sub

Private Sub regbtn_Click()
registerado.Recordset.Fields("rollno") = txtroll.Text
registerado.Recordset.Fields("name") = txtname.Text
registerado.Recordset.Fields("class") = txtclass.Text
registerado.Recordset.Fields("username") = txtuser.Text
registerado.Recordset.Fields("password") = txtpass.Text
registerado.Recordset.Fields("address") = txtadd.Text
registerado.Recordset.Fields("contact") = txtphone.Text
registerado.Recordset.Update
MsgBox "User Registration Successful, Please Login User name andf Password", vbInformation
login.Show
register.Hide
End Sub

Private Sub resetbtn_Click()
txtroll.Text = ""
txtname.Text = ""
txtclass.Text = ""
txtuser.Text = ""
txtpass.Text = ""
txtadd.Text = ""
txtphone.Text = ""
End Sub
