VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "ATTENDANCE--GENERATOR"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
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
      Left            =   9480
      TabIndex        =   4
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE FINAL YEAR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   3
      Top             =   6600
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE THIRD YEAR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   2
      Top             =   4800
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE SECOND YEAR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ATTENDANCE - GENERATOR "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form3.Show

End Sub

Private Sub Command3_Click()
Form4.Show

End Sub

Private Sub Command4_Click()
End
End Sub
