VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form student_signup 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc signupado 
      Height          =   330
      Left            =   5760
      Top             =   4680
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "new_user"
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
   Begin VB.CommandButton backbtn 
      BackColor       =   &H00C0C000&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton savebtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox conpasstxt 
      Appearance      =   0  'Flat
      DataField       =   "confirm password"
      DataSource      =   "signupado"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox passtxt 
      Appearance      =   0  'Flat
      DataField       =   "password"
      DataSource      =   "signupado"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox useridtxt 
      Appearance      =   0  'Flat
      DataField       =   "userid"
      DataSource      =   "signupado"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Confirm password :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Choose a password :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "User Id :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Student Sign-up Pge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "student_signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub backbtn_Click()
stdlogin.Show
student_signup.Hide
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "Select *from new_user", con, adOpenDynamic, adLockPessimistic

End Sub


Private Sub savebtn_Click()
rs.Fields("userid") = useridtxt.Text

    If passtxt.Text = conpasstxt.Text Then
        rs.Fields("password") = passtxt.Text
        rs.Update
        MsgBox "User Registration Successful, Please Login User name andf Password", vbInformation
        stdlogin.Show
        student_signup.Hide
    Else
        MsgBox "PASSWORD not matched"
        passtxt.Text = ""
        conpasstxt.Text = ""
        passtxt.SetFocus
    End If

End Sub

