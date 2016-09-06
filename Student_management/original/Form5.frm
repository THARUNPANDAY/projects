VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   Caption         =   "Form5"
   ClientHeight    =   10950
   ClientLeft      =   -90
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   13680
      TabIndex        =   35
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
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
      Left            =   4320
      TabIndex        =   33
      Top             =   9480
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   5400
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   31
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   30
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   29
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   28
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   27
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   25
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   24
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   23
      Top             =   7560
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   20
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   8400
      TabIndex        =   19
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   18
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000B&
      Caption         =   "FIND"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14520
      TabIndex        =   17
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   16
      Top             =   9480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   15
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000007&
      Caption         =   "INTERNAL :"
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
      Left            =   11640
      TabIndex        =   34
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      Caption         =   "STUDENT INFORMATION SYSTEM"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   7440
      TabIndex        =   32
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000007&
      Caption         =   "MARK 6 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   13
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      Caption         =   "MARK 5 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   12
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      Caption         =   "MARK 4 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Caption         =   "MARK 3 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12000
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "MARK 2 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Caption         =   "MARK 1 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12000
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000001&
      Caption         =   "SUBJECT CODE 6 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "SUBJECT CODE 5 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   "SUBJECT CODE 4 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "SUBJECT CODE 3 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "SUBJECT CODE 2 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "SUBJECT CODE 1 :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "BATCH* :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "REGISTER NUMBER* :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Form4.Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_internal_marks", con, adOpenDynamic, adLockPessimistic
rs.AddNew
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Added Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
End If
End Sub

Private Sub Command2_Click()
clear
Me.Visible = True
Text1.SetFocus
End Sub

Private Sub Command3_Click()
If Form4.Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
ElseIf Form4.Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_internal_marks", con, adOpenDynamic, adLockPessimistic
If Form4.Option1.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
ElseIf Form4.Option2.Value = True Then
With rs
.Fields("register_number").Value = Text13.Text
.Fields("batch").Value = Text14.Text
.Fields("sub_code1").Value = Text2.Text
.Fields("sub_code2").Value = Text3.Text
.Fields("sub_code3").Value = Text1.Text
.Fields("sub_code4").Value = Text4.Text
.Fields("sub_code5").Value = Text5.Text
.Fields("sub_code6").Value = Text6.Text
.Fields("mark1").Value = Text7.Text
.Fields("mark2").Value = Text8.Text
.Fields("mark3").Value = Text9.Text
.Fields("mark4").Value = Text12.Text
.Fields("mark5").Value = Text11.Text
.Fields("mark6").Value = Text10.Text
.Fields("internal").Value = Text15.Text
End With
MsgBox "Data Has Been Updated Successfully!!!", vbInformation, "MESSAGE"
rs.Update
End If
rs.Close
End If
End Sub

Private Sub Command4_Click()
If Form4.Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
ElseIf Form4.Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
ElseIf Form4.Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
ElseIf Form4.Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
ElseIf Form4.Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
ElseIf Form4.Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_internal_marks where register_number='" + Text13.Text + "' and batch='" + Text14.Text + "' and internal='" + Text15.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("sub_code1").Value
Text3.Text = rs.Fields("sub_code2").Value
Text1.Text = rs.Fields("sub_code3").Value
Text4.Text = rs.Fields("sub_code4").Value
Text5.Text = rs.Fields("sub_code5").Value
Text6.Text = rs.Fields("sub_code6").Value
Text7.Text = rs.Fields("mark1").Value
Text8.Text = rs.Fields("mark2").Value
Text9.Text = rs.Fields("mark3").Value
Text12.Text = rs.Fields("mark4").Value
Text11.Text = rs.Fields("mark5").Value
Text10.Text = rs.Fields("mark6").Value
End If
rs.Close
End Sub

Private Sub Command5_Click()
con.Close
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
Me.Visible = True
Text13.SetFocus
If Form4.Option1.Value = True Then
Label8.Visible = False
Label14.Visible = False
Text6.Visible = False
Text10.Visible = False
ElseIf Form4.Option2.Value = True Then
Label8.Visible = True
Label14.Visible = True
Text6.Visible = True
Text10.Visible = True
End If
End Sub
Private Sub clear()
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
End Sub
