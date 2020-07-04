VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PAYMENT3RD 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   4800
      TabIndex        =   17
      Top             =   4800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   112394241
      CurrentDate     =   43228
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "SELECT PAYMENT NAME"
      Height          =   1695
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   6255
      Begin VB.OptionButton Option12 
         BackColor       =   &H0080FFFF&
         Caption         =   "DECEBER"
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H0080FFFF&
         Caption         =   "NOVEMBER "
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H0080FFFF&
         Caption         =   "OCTOBER"
         Height          =   495
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H0080FFFF&
         Caption         =   "SEPTEMBER"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H0080FFFF&
         Caption         =   "AUGUST"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080FFFF&
         Caption         =   "JULY"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H0080FFFF&
         Caption         =   "JUNE"
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080FFFF&
         Caption         =   "MAY"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080FFFF&
         Caption         =   "APRIL"
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080FFFF&
         Caption         =   "MARCH"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "FEBUARY"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "JANUARY"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8640
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "ENTER AMOUNT FOR THIS YEAR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "ENTER ROLL NUMBER"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "PAYMENT3RD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()


rs.Close
rs.Open "Select * from payment where Roll like '" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
Add
reload
Else
rs.AddNew
Add
End If

End Sub

Sub reload()
rs.Close
rs.Open "Select * from payment", con, adOpenDynamic, adLockPessimistic
End Sub



Sub Add()

    rs.Fields("Roll").Value = Text1.Text
    
    If Option1.Value = True Then
        rs.Fields("JANUARY").Value = DTPicker1.Value
        
    End If
    If Option2.Value = True Then
        rs.Fields("FEBUARY").Value = DTPicker1.Value
    End If
    If Option3.Value = True Then
        rs.Fields("MARCH").Value = DTPicker1.Value
    End If
    If Option4.Value = True Then
        rs.Fields("APRIL").Value = DTPicker1.Value
    End If
    If Option5.Value = True Then
        rs.Fields("MAY").Value = DTPicker1.Value
    End If
    If Option6.Value = True Then
        rs.Fields("JUNE").Value = DTPicker1.Value
    End If
    If Option7.Value = True Then
        rs.Fields("JULY").Value = DTPicker1.Value
    End If
    If Option8.Value = True Then
        rs.Fields("AUGUST").Value = DTPicker1.Value
    End If
    If Option9.Value = True Then
        rs.Fields("SEPTEMBER").Value = DTPicker1.Value
    End If
    If Option10.Value = True Then
        rs.Fields("OCTOBER").Value = DTPicker1.Value
    End If
    If Option11.Value = True Then
        rs.Fields("NOVEMBER").Value = DTPicker1.Value
    End If
    If Option12.Value = True Then
        rs.Fields("DECEMBER").Value = DTPicker1.Value
    End If
        
        MsgBox "DATA IS SAVED SUCCESSFULLY", vbInformation
        rs.Update
        

End Sub


Private Sub Command2_Click()
ENTRY.Show
Me.Hide
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "Select *from payment", con, adOpenDynamic, adLockPessimistic

End Sub


Private Sub Option1_Click()
Label3.Visible = True
        Text3.Visible = True
        If Text3.Text = "" Then
            MsgBox "PLEASE ENTER AMOUNT"
        End If
End Sub

