VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   735
      Left            =   10920
      TabIndex        =   28
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   13320
      TabIndex        =   27
      Text            =   "0"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   13560
      TabIndex        =   26
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   13560
      TabIndex        =   25
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   13560
      TabIndex        =   24
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   13560
      TabIndex        =   23
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0000FFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   3360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   20280
      Top             =   10920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tharunpro"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
      Caption         =   "7"
      Height          =   495
      Left            =   9480
      TabIndex        =   20
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "8"
      Height          =   495
      Left            =   10080
      TabIndex        =   19
      Top             =   9240
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00000000&
      Caption         =   "6"
      Height          =   495
      Left            =   8880
      TabIndex        =   17
      Top             =   9240
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   8760
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Top             =   9240
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   9120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      Height          =   495
      Left            =   8040
      TabIndex        =   13
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9840
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   9840
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   9840
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   9840
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
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
      Left            =   2640
      TabIndex        =   29
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18480
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Caption         =   "GOTO :"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   12360
      TabIndex        =   12
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   12360
      TabIndex        =   11
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   12360
      TabIndex        =   10
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   12360
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "BEREBAVAIT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where BEREBAVAIT = '" & Text1.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label3.BackColor = &HFF&
Text9.Text = 0
Else
Label3.BackColor = &HFF00&
Text9.Text = 1
End If
End Sub

Private Sub Command10_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form7.Show
End Sub

Private Sub Command11_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form9.Show
End Sub

Private Sub Command12_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form8.Show
End Sub

Private Sub Command2_Click()
connect
cmd.Close
cmd.Open "select * from tharunpro where BEREBAVAIT = '" & Text2.Text & "'", con, adOpenDynamic, adLockPessimistic
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
cmd.Open "select * from tharunpro where BEREBAVAIT = '" & Text3.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label4.BackColor = &HFF&
Text7.Text = 0
Else
Label4.BackColor = &HFF00&
Text7.Text = 1
End If
End Sub

Private Sub Command4_Click()
''If Text4.Text = "BEER" Then Label5.BackColor = &HFF00& Else Label5.BackColor = &HFF&
connect
cmd.Close
cmd.Open "select * from tharunpro where BEREBAVAIT = '" & Text4.Text & "'", con, adOpenDynamic, adLockPessimistic
If cmd.EOF Then
Label5.BackColor = &HFF&
Text6.Text = 0
Else
Label5.BackColor = &HFF00&
Text6.Text = 1
End If

End Sub

Private Sub Command5_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form6.Show
End Sub

Private Sub Command6_Click()
Text10.Text = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)
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
Form5.Show
End Sub

Private Sub Command9_Click()
If Text5.Text = 640 Then
Form10.Show
End If
Form4.Show
End Sub

Private Sub Form_Load()
Text5.Text = Form1.Text2.Text
End Sub


Private Sub Text1_Change()
Command1_Click
End Sub

Private Sub Text2_Change()
Command2_Click
End Sub

Private Sub Text3_Change()
Command3_Click
End Sub

Private Sub Text4_Change()
Command4_Click
End Sub

Private Sub Timer1_Timer()
Text5.Text = Val(Text5.Text) + 1
Command6_Click
If Text5.Text = 80 Then
Me.Hide
Form3.Show
End If
End Sub
