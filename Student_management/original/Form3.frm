VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   645
   ClientWidth     =   20085
   LinkTopic       =   "Form3"
   ScaleHeight     =   10515
   ScaleWidth      =   20085
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   22
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   21
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   19
      Top             =   9840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   840
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1005
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from cse_student_info"
      Caption         =   " <--PREVIOUS    NEXT-->"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   13
      Text            =   "select department"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   9360
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   25
      Top             =   6480
      Width           =   4335
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "EMAIL :"
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
      Left            =   8160
      TabIndex        =   24
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000007&
      Caption         =   "BATCH :"
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
      Height          =   615
      Left            =   1560
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000007&
      Caption         =   "DEPARTMENT :"
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
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   1200
      Width           =   2895
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
      Left            =   7920
      TabIndex        =   14
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   12
      Top             =   8880
      Width           =   4335
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "PHONE NUMBER :"
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
      Height          =   615
      Left            =   6720
      TabIndex        =   11
      Top             =   8880
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   10
      Top             =   8040
      Width           =   4335
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "ADDRESS :"
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
      Height          =   615
      Left            =   7680
      TabIndex        =   9
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   8
      Top             =   7200
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "GENDER :"
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
      Height          =   615
      Left            =   7920
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   6
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "DOB :"
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
      Height          =   615
      Left            =   8520
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   4
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "NAME :"
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
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10320
      TabIndex        =   2
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "REGISTER NUMBER :"
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
      Height          =   615
      Left            =   6120
      TabIndex        =   1
      Top             =   3960
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If Combo1.Text = "COMPUTER SCIENCE" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from cse_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from cse_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
ElseIf Combo1.Text = "ELECTRICAL" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from eee_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from eee_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
ElseIf Combo1.Text = "ELECTRONICS" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from ece_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from ece_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
ElseIf Combo1.Text = "CIVIL" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from civil_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from civil_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
ElseIf Combo1.Text = "MECHANICAL" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from mech_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from mech_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
ElseIf Combo1.Text = "AUTOMOBILE" And Not Text1.Text = "" Then
Adodc1.Recordset.Close
rs.Close
Adodc1.Recordset.Open "select * from auto_student_info where batch='" + Text1.Text + "'"
rs.Open "select * from auto_student_info where batch='" + Text1.Text + "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
If rs.EOF = True Then
MsgBox "please check the Information correctly", vbCritical + vbInformation, "ERROR:402 data wrong"
Form3.Visible = True
Combo1.SetFocus
Else
display
End If
Else
MsgBox "Please Enter Correct Department and Batch", vbCritical + vbQuestion, "Error"
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveFirst
rs.MoveFirst
display
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MovePrevious
rs.MovePrevious
If rs.BOF Then
MsgBox "First Record"
Else
display
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveNext
rs.MoveNext
If rs.EOF Then
MsgBox "End Of Records"
Else
display
End If
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
rs.MoveLast
display
End Sub

Private Sub Form_Load()
Me.Visible = True
rs.Open "select * from admin", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False", adOpenDynamic, adLockPessimistic
Combo1.SetFocus
Combo1.AddItem ("COMPUTER SCIENCE")
Combo1.AddItem ("ELECTRICAL")
Combo1.AddItem ("ELECTRONICS")
Combo1.AddItem ("CIVIL")
Combo1.AddItem ("MECHANICAL")
Combo1.AddItem ("AUTOMOBILE")
End Sub
Private Sub display()
Label3.Caption = Adodc1.Recordset.Fields("register_number").Value
Label5.Caption = Adodc1.Recordset.Fields("name").Value
Label7.Caption = Adodc1.Recordset.Fields("dob").Value
Label15.Caption = Adodc1.Recordset.Fields("email").Value
Label9.Caption = Adodc1.Recordset.Fields("gender").Value
Label11.Caption = Adodc1.Recordset.Fields("address").Value
Label13.Caption = Adodc1.Recordset.Fields("phone_no").Value
Picture1.Picture = LoadPicture(Adodc1.Recordset.Fields("photo").Value)
End Sub
