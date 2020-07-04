VERSION 5.00
Begin VB.Form welcome 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton stdloginbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Student"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton adminlogbtn 
      BackColor       =   &H00FFFF00&
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOG IN FORM"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub
Private Sub adminlogbtn_Click()
login.Show
welcome.Hide
stdlogin.Hide
End Sub

Private Sub stdloginbtn_Click()
stdlogin.Show
welcome.Hide
login.Hide
End Sub
