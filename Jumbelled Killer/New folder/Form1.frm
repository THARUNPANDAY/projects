VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.OptionButton Option2 
      Caption         =   "Expert"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000007&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   51.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1185
      Left            =   8160
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   8760
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   3840
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
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
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   20040
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00B7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "PROFESSIONAL"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "/ 640"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "JUMBBLED    KILLER"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   12495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "NAME :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Adodc1_Click()

End Sub

Private Sub Command1_Click()
connect
rreg.Close
rreg.Open "select * from reg", con, adOpenDynamic, adLockPessimistic
rreg.AddNew
If Text1.Text <> Empty Then
rreg("name") = InputBox("ENTER YOUR NAME HERE :")
rreg.Update
End If
End Sub


Private Sub Form_Load()
Command1_Click
End Sub

Private Sub Option1_Click()
Form2.Show
End Sub

Private Sub Option2_Click()
Form2.Show
End Sub

Private Sub Text1_Change()
Command1_Click
End Sub

Private Sub Timer1_Timer()
Text2.Text = Val(Text2.Text) + 1
If Text2.Text = 620 Then
Me.Hide
Form10.Show
End If
End Sub
