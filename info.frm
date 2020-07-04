VERSION 5.00
Begin VB.Form info 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000C0&
      Caption         =   "FEEDBACK/COMPLAINT"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "LOG OUT"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "VIEW PAYMENT INFORMATION"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
      Caption         =   "VIEW DETAILS"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    studentprofile.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
PAYMENT1ST.Show
Me.Hide
End Sub

Private Sub Command3_Click()
welcome.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Me.Hide
complainbox.Show
End Sub
