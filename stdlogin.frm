VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form stdlogin 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc stdado 
      Height          =   330
      Left            =   5880
      Top             =   4320
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from STUDENT_REGISTRATION"
      Caption         =   "stdado"
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
   Begin VB.CommandButton newuserbtn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NEW USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton loginbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Login"
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
      Left            =   1320
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox passtxt 
      DataField       =   "password"
      DataSource      =   "stdado"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox useridtxt 
      DataField       =   "userid"
      DataSource      =   "stdado"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Login Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "stdlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "Select *from new_user", con, adOpenDynamic, adLockPessimistic

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub cancel_Click()
    welcome.Show
    stdlogin.Hide
End Sub

Private Sub loginbtn_Click()
stdado.RecordSource = "select *from new_user where userid='" + useridtxt.Text + "' and password='" + passtxt.Text + "' "
stdado.Refresh
If stdado.Recordset.EOF Then
MsgBox "Login Failed"
Else
     MsgBox "Login Successful", vbInformation
    info.Show
    stdlogin.Hide
End If




'If rs.EOF Then
'    MsgBox "Login Failed...Please login with correct  password", vbCritical
'    stdlogin.Show
'    useridtxt.Text = ""
'    passtxt.Text = ""
'    End
'Else
 '   MsgBox "Login Successful", vbInformation
'    info.Show
 '   stdlogin.Hide
' End If
End Sub


Private Sub newuserbtn_Click()
student_signup.Show
stdlogin.Hide
End Sub
