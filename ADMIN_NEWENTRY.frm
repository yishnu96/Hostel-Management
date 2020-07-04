VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ADMIN_NEWENTRY 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   FillColor       =   &H00C0FFC0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10200
      Top             =   6720
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\VB PROJECT\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\VB PROJECT\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from STUDENT_REGISTRATION"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton BACKCMD 
      BackColor       =   &H008080FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Additional Details  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   47
      Top             =   8280
      Width           =   8175
      Begin VB.ComboBox Combo9 
         Height          =   330
         Left            =   5040
         TabIndex        =   57
         Text            =   "Combo9"
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox Combo8 
         Height          =   330
         Left            =   5040
         TabIndex        =   54
         Text            =   "AC OR NON AC"
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox Combo7 
         Height          =   330
         Left            =   1200
         TabIndex        =   53
         Text            =   "SELECT"
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Non-Veg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   52
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Veg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1200
         TabIndex        =   51
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Room No  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   56
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   50
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Room  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Food  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Contact Details "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   40
      Top             =   6840
      Width           =   8175
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   1560
         TabIndex        =   46
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   5520
         TabIndex        =   45
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   1560
         TabIndex        =   44
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Email  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate No  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   42
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Address Details "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   960
      TabIndex        =   25
      Top             =   4440
      Width           =   8175
      Begin VB.ComboBox Combo6 
         Height          =   330
         Left            =   1920
         TabIndex        =   58
         Text            =   "COUNTRY"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   5640
         TabIndex        =   38
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox Combo5 
         Height          =   330
         Left            =   1920
         TabIndex        =   36
         Text            =   "SELECT STATE"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ComboBox Combo4 
         Height          =   330
         Left            =   5640
         TabIndex        =   34
         Text            =   "DISTRICT"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1920
         TabIndex        =   30
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Residencial"
         Height          =   330
         Left            =   3360
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Permanent"
         Height          =   330
         Left            =   1920
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Country  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pin  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "State  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "District  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "City/Village  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Address  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Address Type  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Course Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   16
      Top             =   2640
      Width           =   8175
      Begin VB.ComboBox Combo3 
         Height          =   330
         Left            =   5280
         TabIndex        =   24
         Text            =   "SELECT SEMESTER"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Text            =   "SELECT DEPARTMENT"
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   1680
         TabIndex        =   19
         Text            =   "SELECT COURSE"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Roll No  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Semester  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Department  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Course  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   8175
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   2535
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Female"
         Height          =   255
         Left            =   6840
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Male"
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Format          =   127926273
         CurrentDate     =   43226
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Guardian's Name  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "D.O.B  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "NEW ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton CREATEBTN 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CREATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton uploadcmd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "UPLOAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   9360
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGESTRATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "ADMIN_NEWENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim confirm As Integer


Private Sub BACKCMD_Click()
ADMIN_NEWENTRY.Hide
ENTRY.Show
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text1.SetFocus
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False


End Sub

Private Sub Createbtn_Click()
rs.AddNew
        rs.Fields("S_name").Value = Text1.Text
        rs.Fields("FATHERNAME").Value = Text2.Text
        rs.Fields("DOB").Value = DTPicker1.Value
If Option3.Value = True Then
        rs.Fields("Gender").Value = "Male"
End If
If Option4.Value = True Then
        rs.Fields("Gender").Value = "Female"
End If
    
        rs.Fields("GUARDIAN'SNAME").Value = Text3.Text
        rs.Fields("COURSE").Value = Combo1.Text
        rs.Fields("DEPARTMENT").Value = Combo2.Text
        rs.Fields("SEMESTER").Value = Combo3.Text
        rs.Fields("ROLL").Value = Text4.Text
 If Option1.Value = True Then
        rs.Fields("ADDRESSTYPE").Value = Option1.Caption
End If

If Option2.Value = True Then
        rs.Fields("ADDRESSTYPE").Value = Option2.Caption
End If
        
        rs.Fields("ADDRESS").Value = Text5.Text

        
        rs.Fields("CITY").Value = Text6.Text
        rs.Fields("DISTRICT").Value = Combo4.Text
        rs.Fields("STATE").Value = Combo5.Text
        rs.Fields("COUNTRY").Value = Combo6.Text
        rs.Fields("PIN").Value = Text7.Text
        rs.Fields("MOBILE").Value = Text8.Text
        rs.Fields("ALTERNATE").Value = Text9.Text
        rs.Fields("EMAIL").Value = Text10.Text

        
If Option5.Value = True Then
        rs.Fields("Food").Value = Option5.Caption
End If
If Option6.Value = True Then
        rs.Fields("Food").Value = Option6.Caption
End If

        rs.Fields("ROOMTYPE").Value = Combo8.Text
        rs.Fields("ROOM").Value = Combo7.Text
        rs.Fields("ROOMNUMBER").Value = Combo9.Text
        
        rs.Fields("PICTURE").Value = str
        MsgBox "DATA IS SAVED SUCCESSFULLY", vbInformation
        rs.Update
End Sub





Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "Select *from STUDENT_REGISTRATION", con, adOpenDynamic, adLockPessimistic
Combo1.AddItem "B.TECH"
Combo1.AddItem "DIPLOMA"
Combo2.AddItem "CSE"
Combo2.AddItem "ECE"
Combo2.AddItem "CE"
Combo2.AddItem "EE"
Combo2.AddItem "ME"

Combo3.AddItem "SEMESTER-I"
Combo3.AddItem "SEMESTER-II"
Combo3.AddItem "SEMESTER-III"
Combo3.AddItem "SEMESTER-IV"
Combo3.AddItem "SEMESTER-V"
Combo3.AddItem "SEMESTER-VI"
Combo3.AddItem "SEMESTER-VII"
Combo3.AddItem "SEMESTER-VIII"

Combo4.AddItem "HOWRAH"
Combo4.AddItem "KOLKATA"

Combo5.AddItem "WEST BENGAL"
Combo5.AddItem "KALARA"

Combo6.AddItem "INDIA"
Combo6.AddItem "USA"

Combo7.AddItem "SINGLE BED"
Combo7.AddItem "DOUBLE BED"
Combo7.AddItem "TRIPLE BED"

Combo8.AddItem "AC"
Combo8.AddItem "NON-AC"

Combo9.AddItem "400"
Combo9.AddItem "401"
Combo9.AddItem "402"
Combo9.AddItem "403"
Combo9.AddItem "404"
End Sub





Private Sub uploadcmd_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Image1.Stretch = True
Image1.Picture = LoadPicture(str)
End Sub




