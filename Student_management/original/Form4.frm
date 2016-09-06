VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000007&
   Caption         =   "Form4"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20130
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20130
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000007&
      Caption         =   "6-SUBJECT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12360
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000007&
      Caption         =   "5-SUBJECT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9480
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD MARKS"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   4
      Top             =   4680
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   510
      Left            =   9480
      TabIndex        =   3
      Text            =   "Select Department"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "NUMBER OF SUBJECTS :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "DEPARTMENT :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "STUDENT INFORMATION RECORD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Show
End Sub

Private Sub Form_Load()
Me.Visible = True
Combo1.SetFocus
Combo1.AddItem ("COMPUTER SCIENCE")
Combo1.AddItem ("ELECTRICAL")
Combo1.AddItem ("ELECTRONICS")
Combo1.AddItem ("CIVIL")
Combo1.AddItem ("MECHANICAL")
Combo1.AddItem ("AUTOMOBILE")
End Sub

