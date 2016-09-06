VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "exit"
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "change color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
Option1.Value = True
Form1.BackColor = RGB(0, 0, 255)
End Sub

Private Sub Command1_Click()
Form1.BackColor = RGB(0, 0, 255)
Form1.BackColor = RGB(255, 0, 255)
Form1.BackColor = RGB(0, 255, 255)
Form1.BackColor = RGB(255, 0, 0)
Command3_Click
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form1.BackColor = RGB(0, 0, 255)
Form1.BackColor = RGB(255, 0, 255)
Form1.BackColor = RGB(0, 255, 255)
Form1.BackColor = RGB(255, 0, 0)
Command4_Click
End Sub

Private Sub Command4_Click()
Form1.BackColor = RGB(0, 0, 255)
Form1.BackColor = RGB(255, 0, 255)
Form1.BackColor = RGB(0, 255, 255)
Form1.BackColor = RGB(255, 0, 0)
Command1_Click
End Sub
