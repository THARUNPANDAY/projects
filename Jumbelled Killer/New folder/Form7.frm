VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00800080&
   Caption         =   "Form7"
   ClientHeight    =   10950
   ClientLeft      =   -30
   ClientTop       =   300
   ClientWidth     =   20250
   LinkTopic       =   "Form9"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   11400
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   13560
      TabIndex        =   27
      Text            =   "0"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   13920
      TabIndex        =   26
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   13920
      TabIndex        =   25
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   13920
      TabIndex        =   24
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   13920
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   3120
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   525
      Left            =   1200
      TabIndex        =   22
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "7"
      Height          =   495
      Left            =   10320
      TabIndex        =   19
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "4"
      Height          =   495
      Left            =   9120
      TabIndex        =   18
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "5"
      Height          =   495
      Left            =   9720
      TabIndex        =   17
      Top             =   9480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "3"
      Height          =   495
      Left            =   8520
      TabIndex        =   16
      Top             =   9480
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "2"
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000001&
      Caption         =   "1"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   9480
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "8"
      Height          =   495
      Left            =   10920
      TabIndex        =   13
      Top             =   9480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800080&
      Caption         =   "/640"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18000
      TabIndex        =   21
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800080&
      Caption         =   "GOTO :"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   20
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   12960
      TabIndex        =   12
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   12960
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   12960
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   12960
      TabIndex        =   9
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "ONNOUYASM"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where ONNOUYASM = '" & Text1.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label5.BackColor = &HFF&
Text9.Text = 0
Else
Label5.BackColor = &HFF00&
Text9.Text = 1
End If
End Sub

Private Sub Command10_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form5.Show
End Sub

Private Sub Command11_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form8.Show
End Sub

Private Sub Command12_Click()
Text10.Text = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)
End Sub

Private Sub Command2_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where ONNOUYASM= '" & Text2.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label2.BackColor = &HFF&
Text8.Text = 0
Else
Label2.BackColor = &HFF00&
Text8.Text = 1
End If
End Sub

Private Sub Command3_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where ONNOUYASM = '" & Text3.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label3.BackColor = &HFF&
Text7.Text = 0
Else
Label3.BackColor = &HFF00&
Text7.Text = 1
End If
End Sub

Private Sub Command4_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where ONNOUYASM = '" & Text4.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label4.BackColor = &HFF&
Text6.Text = 0
Else
Label4.BackColor = &HFF00&
Text6.Text = 1
End If
End Sub

Private Sub Command5_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form9.Show
End Sub

Private Sub Command6_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form2.Show
End Sub

Private Sub Command7_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form3.Show
End Sub

Private Sub Command8_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form4.Show
End Sub

Private Sub Command9_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form6.Show
End Sub


Private Sub Form_Load()
Text5.Text = Form6.Text5.Text
End Sub

Private Sub Text4_Change()
Command4_Click
End Sub
Private Sub Text3_Change()
Command3_Click
End Sub
Private Sub Text2_Change()
Command2_Click
End Sub
Private Sub Text1_Change()
Command1_Click
End Sub

Private Sub Timer1_Timer()
Text5.Text = Val(Text5.Text) + 1
Command12_Click
If Text5.Text = 480 Then
Me.Hide
Form8.Show
End If
End Sub
