VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form complainbox 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8760
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "SUBMIT"
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   5655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Text            =   "SELECT TYPE OF COMPLAINT OR FEEDBACK"
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "BACK"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "Enter Room Number"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "IF YOU HAVE ANY SUGGESTION OR COMPLAIN PLEASE FEEL FREE TO SAY IT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
End
Attribute VB_Name = "complainbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Me.Hide
info.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs.Fields("ROOM").Value = Text2.Text
    rs.Fields("TYPE").Value = Combo1.Text
    rs.Fields("CONTN").Value = Text1.Text
    MsgBox "YOUR COMPLAINT/SUGGESTION HAS BEEN SEND", vbInformation
    rs.Update
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "Select * from FEEDBACK", con, adOpenDynamic, adLockPessimistic
Combo1.AddItem "MANAGEMENT"
Combo1.AddItem "FOOD"
Combo1.AddItem "ROOM"
Combo1.AddItem "PAYMENT"
Combo1.AddItem "WATER"
Combo1.AddItem "OTHER"


End Sub
