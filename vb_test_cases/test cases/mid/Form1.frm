VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   3720
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x, y As Integer
x = 1
y = ((Len(Text1.Text) + 1) / 5)
For i = 1 To y
MsgBox (Mid(Text1.Text, x, 4))
x = x + 5
Next
End Sub
