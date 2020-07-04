VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form2"
   ClientHeight    =   5880
   ClientLeft      =   7080
   ClientTop       =   3180
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   ScaleHeight     =   5880
   ScaleWidth      =   8235
   Begin MSAdodcLib.Adodc loginado 
      Height          =   330
      Left            =   6000
      Top             =   5400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from admin_loginpassword"
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
      BackColor       =   &H00FF8080&
      Caption         =   "Reset Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtpass 
      DataField       =   "password"
      DataSource      =   "loginado"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtuser 
      DataField       =   "user id"
      DataSource      =   "loginado"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Login Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset





Private Sub loginbtn_Click()
loginado.RecordSource = "select *from admin_loginpassword where userid='" + txtuser.Text + "' and password='" + txtpass.Text + "' "
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Login Failed...Please login with correct  password"
Else
MsgBox "Login Successful", vbInformation
ENTRY.Show
login.Hide
End If
End Sub
Private Sub cancel_Click()
welcome.Show
login.Hide
End Sub


Private Sub resetbtn_Click()
reset.Show
login.Hide
End Sub

