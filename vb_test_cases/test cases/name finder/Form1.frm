VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "THARUN SOFTWARES"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "FIND ABBREVATION OF NAME"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   1
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   10080
      TabIndex        =   0
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "ENTER YOUR NAME :"
      BeginProperty Font 
         Name            =   "Andale Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "KNOW WHAT'S THE ABBREVATION OF                         YOUR NAME "
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   7095
   End
   Begin VB.Menu FILE_ 
      Caption         =   "FILE"
      Begin VB.Menu CLEAR 
         Caption         =   "CLEAR TEXT"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu EXIT_ 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CLEAR_Click()
Text1.Text = ""
End Sub

Private Sub Command1_Click()
A = Len(Text1.Text)
For i = 1 To A
b = Mid(Text1.Text, i, 1)
b = LCase(b)
Select Case (b)
Case "a"
MsgBox "A-->AGGRESIVE"
Case "b"
MsgBox "B-->BRILLIANT"
Case "c"
MsgBox "C-->COURAGEOUS"
Case "d"
MsgBox "D-->DASHING (OR) DIFFERENT"
Case "e"
MsgBox "E-->ENCOURAGEABLE"
Case "f"
MsgBox "F-->FANTASTIC"
Case "g"
MsgBox "G-->GENIUS"
Case "h"
MsgBox "H-->HONOURABLE"
Case "i"
MsgBox "I-->INTELLIGENT"
Case "j"
MsgBox "J-->JOVIAL"
Case "k"
MsgBox "K-->KIND"
Case "l"
MsgBox "L-->LOVEABLE"
Case "m"
MsgBox "M-->MYSTERIOUS"
Case "n"
MsgBox "N-->NOBLE"
Case "o"
MsgBox "O-->ORGANISER"
Case "p"
MsgBox "P-->PREETY"
Case "q"
MsgBox "Q-->QUITE"
Case "r"
MsgBox "R-->ROBUST"
Case "s"
MsgBox "S-->STRAIGHT FORWARD"
Case "t"
MsgBox "T-->THINKER"
Case "u"
MsgBox "U-->UNBEATABLE"
Case "v"
MsgBox "V-->VIGROUS"
Case "w"
MsgBox "W-->WORTHFULL"
Case "x"
MsgBox "X-->XLR8"
Case "y"
MsgBox "Y-->YOUNG"
Case "z"
MsgBox "Z-->zero"
End Select
Next
End Sub

Private Sub EXIT__Click()
End
End Sub

