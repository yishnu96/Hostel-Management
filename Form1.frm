VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form spalsh 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "progess bar"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10905
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4920
      Top             =   2760
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2805
      Left            =   720
      Picture         =   "Form1.frx":0000
      Top             =   360
      Width           =   4185
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 10.1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO BUNK BOUTIQUE HOSTEL"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   5640
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "spalsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label4.Caption = "Loading...Please wait....."
Label5.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
welcome.Show
End If
End Sub
