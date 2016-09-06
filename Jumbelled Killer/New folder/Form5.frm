VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form5"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   13440
      TabIndex        =   28
      Text            =   "0"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   11400
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   13560
      TabIndex        =   26
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   13560
      TabIndex        =   25
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   13560
      TabIndex        =   24
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   13560
      TabIndex        =   23
      Top             =   3120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   4080
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1440
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "6"
      Height          =   495
      Left            =   9960
      TabIndex        =   20
      Top             =   9120
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "8"
      Height          =   495
      Left            =   11280
      TabIndex        =   19
      Top             =   9120
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "7"
      Height          =   495
      Left            =   10680
      TabIndex        =   18
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "5"
      Height          =   495
      Left            =   9360
      TabIndex        =   17
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command7 
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
      Left            =   9720
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
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
      Left            =   9720
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
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
      Left            =   9720
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   9720
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   9120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   9120
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
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
      Height          =   615
      Left            =   2640
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17880
      TabIndex        =   21
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "GOTO :"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   12480
      TabIndex        =   4
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   12480
      TabIndex        =   3
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   12480
      TabIndex        =   2
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   12480
      TabIndex        =   1
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "EHISTECAT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form2.Show
End Sub

Private Sub Command10_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form9.Show
End Sub

Private Sub Command11_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form7.Show
End Sub

Private Sub Command12_Click()
Text10.Text = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)
End Sub

Private Sub Command2_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form4.Show
End Sub

Private Sub Command3_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form3.Show
End Sub

Private Sub Command4_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where EHISTECAT = '" & Text2.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label3.BackColor = &HFF&
Text9.Text = 0
Else
Label3.BackColor = &HFF00&
Text9.Text = 1
End If
End Sub

Private Sub Command5_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where EHISTECAT = '" & Text3.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label4.BackColor = &HFF&
Text8.Text = 0
Else
Label4.BackColor = &HFF00&
Text8.Text = 1
End If
End Sub

Private Sub Command6_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where EHISTECAT = '" & Text4.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label5.BackColor = &HFF&
Text7.Text = 0
Else
Label5.BackColor = &HFF00&
Text7.Text = 1
End If
End Sub

Private Sub Command7_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where EHISTECAT = '" & Text5.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label6.BackColor = &HFF&
Text6.Text = 0
Else
Label6.BackColor = &HFF00&
Text6.Text = 1
End If
End Sub

Private Sub Command8_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form6.Show
End Sub

Private Sub Command9_Click()
If Text1.Text = 640 Then
Form10.Show
End If
Form8.Show
End Sub

Private Sub Form_Load()
Text1.Text = Form4.Text5.Text

End Sub

Private Sub Text2_Change()
Command4_Click
End Sub

Private Sub Text3_Change()
Command5_Click
End Sub

Private Sub Text4_Change()
Command6_Click
End Sub

Private Sub Text5_Change()
Command7_Click
End Sub

Private Sub Timer1_Timer()
Text1.Text = Val(Text1.Text) + 1
Command12_Click
If Text1.Text = 320 Then
Me.Hide
Form6.Show
End If
End Sub
