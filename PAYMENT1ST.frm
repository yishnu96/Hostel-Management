VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PAYMENT1ST 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   9360
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "PAYMENT DETAILS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080FF80&
         Caption         =   "MARCH"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FF80&
         Caption         =   "APRIL"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "MAY"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FF80&
         Caption         =   "JUNE"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         Caption         =   "JULY"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
         Caption         =   "AUGUST"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "SEPTEMBER"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "OCTOBER"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "NOVEMBER"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "DECEMBER"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   20
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "FEBUARY"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "JANUARY"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   3840
         X2              =   3840
         Y1              =   120
         Y2              =   5400
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
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
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9240
      Top             =   1080
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
      RecordSource    =   "select * from STUDENT_REGISTRATION"
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
   Begin VB.CommandButton detailsbtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "REMARK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ID"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "PAYMENT1ST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
ENTRY.Show
Me.Hide

End Sub

Private Sub Command2_Click()
info.Show
Me.Hide
End Sub

Private Sub detailsbtn_Click()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
con.Open
Set rs = New ADODB.Recordset
rs.Open "select * from payment where Roll like '" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic

Frame1.Visible = True
Text6.Visible = True
Text1.Visible = True
Label3.Visible = True
Label4.Visible = True

Text1.Text = rs.Fields("amount")
Text2.Text = rs.Fields("JANUARY")
Text3.Text = rs.Fields("FEBUARY")
Text10.Text = rs.Fields("MARCH")
Text9.Text = rs.Fields("APRIL")
Text8.Text = rs.Fields("MAY")
Text7.Text = rs.Fields("JUNE")
Text5.Text = rs.Fields("JULY")
Text16.Text = rs.Fields("AUGUST")
Text15.Text = rs.Fields("SEPTEMBER")
Text14.Text = rs.Fields("OCTOBER")
Text13.Text = rs.Fields("NOVEMBER")
Text12.Text = rs.Fields("DECEMBER")
End Sub

