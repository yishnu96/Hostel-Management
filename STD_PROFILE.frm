VERSION 5.00
Begin VB.Form STDPROFILE 
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   35
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   960
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   7320
      ScaleHeight     =   2355
      ScaleWidth      =   2115
      TabIndex        =   31
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label33 
      Caption         =   "ROLL NO  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   33
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "PICTURE"
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
      Left            =   7920
      TabIndex        =   32
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label30 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label29 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label28 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label27 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label26 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label25 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label24 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label23 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label22 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label21 
      Caption         =   "Label17"
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label20 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label19 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label18 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label16 
      Caption         =   "ROOM NO :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "ROOM TYPE :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "ROOM  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "FOOD :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "ROLL NO :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "SEMESTER :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "DEPERTMENT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "COURSE :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "EMAIL :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "FATHER NAME :  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "CONTACT NO :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "ADDRESS  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "GENDER  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "D.O.B  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "NAME  :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STUDENT PROFILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "STDPROFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
