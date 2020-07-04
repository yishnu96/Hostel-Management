VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form studentprofile 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   3015
      Left            =   7080
      ScaleHeight     =   2955
      ScaleWidth      =   2475
      TabIndex        =   59
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7320
      Top             =   8640
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
      RecordSource    =   "select *from STUDENT_REGISTRATION"
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
   Begin VB.CommandButton gobtn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GO"
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   33
      Top             =   720
      Width           =   2175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Additional Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   48
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   47
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   46
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Label39"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   45
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label25 
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
         Left            =   3720
         TabIndex        =   29
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label24 
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
         Left            =   3720
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label23 
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
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label22 
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
         Left            =   3720
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depertment  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label20 
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label19 
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
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Contact Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   50
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   49
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Email  :"
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
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label17 
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
         Left            =   3600
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Guardian's Mobile No  :"
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
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label15 
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
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Address Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   57
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   56
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   55
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   54
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   53
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label14 
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         Left            =   3720
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4680
         TabIndex        =   40
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   39
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   37
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label7 
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
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton BACKCMD 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BACK"
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
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
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
      Left            =   480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4440
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "PHOTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7800
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT  PROFILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "studentprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub Command1_Click()
info.Show
Me.Hide
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\hostel management1\login.mdb;Persist Security Info=False"
rs.Open "select *from STUDENT_REGISTRATION", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub DISPLAY()

Label30.Caption = rs.Fields("S_name")
Label31.Caption = rs.Fields("DOB")
Label32.Caption = rs.Fields("FATHERNAME")
Label33.Caption = rs.Fields("GUARDIAN'SNAME")
Label34.Caption = rs.Fields("GENDER")
Label35.Caption = rs.Fields("COURSE")
Label36.Caption = rs.Fields("DEPARTMENT")
Label37.Caption = rs.Fields("SEMESTER")
Label38.Caption = rs.Fields("ROLL")
Label39.Caption = rs.Fields("FOOD")
Label40.Caption = rs.Fields("ROOM")
Label41.Caption = rs.Fields("ROOMTYPE")
Label42.Caption = rs.Fields("ROOMNUMBER")
Label43.Caption = rs.Fields("MOBILE")
Label44.Caption = rs.Fields("ALTERNATE")
Label45.Caption = rs.Fields("EMAIL")
Label46.Caption = rs.Fields("ADDRESSTYPE")
Label47.Caption = rs.Fields("ADDRESS")
Label49.Caption = rs.Fields("CITY")
Label50.Caption = rs.Fields("DISTRICT")
Label51.Caption = rs.Fields("STATE")
Label52.Caption = rs.Fields("PIN")
Label53.Caption = rs.Fields("COUNTRY")
Picture1.Picture = LoadPicture(rs.Fields("PICTURE"))

End Sub
Private Sub BACKCMD_Click()
ENTRY.Show
studentprofile.Hide
End Sub

Private Sub gobtn_Click()

Image1.Visible = True
Frame1.Visible = True
Frame2.Visible = True
Frame1.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = True
Picture1.Visible = True
Label28.Visible = True



rs.Close
rs.Open "select *from STUDENT_REGISTRATION where roll='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic

If rs.EOF Then
MsgBox "Student Data not found...Enter The Correct ROLL NUMBER!!!!!!"
Else
Adodc1.Caption = Adodc1.RecordSource
DISPLAY
End If
End Sub

