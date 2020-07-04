VERSION 5.00
Begin VB.Form ENTRY 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   3405
   ClientTop       =   3120
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9675
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "VIEW COMPLAINTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "New Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton LOGOUTBTN 
      BackColor       =   &H00FF8080&
      Caption         =   "LOG-OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sent Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Edit Student Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton stdprofile 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Student Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Newregbtn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "New Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
PAYMENT3RD.Show
Me.Hide
End Sub

Private Sub Command3_Click()
PAYMENT1ST.Show
Me.Hide
End Sub

Private Sub Command5_Click()
view_compt.Show
Me.Hide

End Sub

Private Sub LOGOUTBTN_Click()
getbutton = MsgBox("Are You Want To Exit?", vbYesNo, MESSAGE)
If getbutton = vbYes Then

login.Show
ENTRY.Hide
End If
End Sub

Private Sub Newregbtn_Click()
ADMIN_NEWENTRY.Show
ENTRY.Hide
End Sub

Private Sub stdprofile_Click()
studentprofile.Show
ENTRY.Hide
End Sub
