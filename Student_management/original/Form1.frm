VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20070
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   20070
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   1440
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      RecordSource    =   "select * from admin"
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
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "ADMIN"
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
      Height          =   735
      Left            =   9488
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "OTHER'S"
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
      Height          =   735
      Left            =   11528
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   9008
      TabIndex        =   5
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   9008
      TabIndex        =   4
      Top             =   3240
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   9120
      TabIndex        =   3
      Text            =   "Select Department"
      Top             =   4440
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "NAME:"
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
      Height          =   615
      Left            =   6488
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "PASSWORD :"
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
      Height          =   735
      Left            =   6488
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "LOGIN AS:"
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
      Height          =   855
      Left            =   6495
      TabIndex        =   9
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "DEPARTMENT :"
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
      Left            =   6368
      TabIndex        =   8
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "FORGET PASSWORD !!!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "STUDENT AND  STAFF  INFORMATION SYSTEM"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please check the data given", vbQuestion, "HELP:RECORRECT"
End If
If Option1.Value = True Then
rs.Close
rs.Open "select * from admin where name ='" + LCase(Text1.Text) + "' and password='" + Text2.Text + "'", con, adOpenDynamic, adLockPessimistic
If rs.EOF Then
MsgBox "Please Check the UserName and Password", vbCritical, "ERROR:Uid/Password Wrong"
Else
entry_log_admin
MsgBox "super"
End If
ElseIf Option2.Value = True Then
rs.Close
rs.Open "select * from admin where name='" + LCase(Text1.Text) + "' and password='" + Text2.Text + "' and department= '" + Combo1.Text + "'"
If rs.EOF Then
MsgBox "Please Check the UserName and Password", vbCritical, "ERROR:Uid/Password Wrong"
Else
entry_log
Form2.Show
End If
End If
End Sub

Private Sub Form_Load()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
rs.Open "select * from admin", con, adOpenDynamic, adLockPessimistic
Text2.PasswordChar = "-"
Combo1.AddItem ("COMPUTER SCIENCE")
Combo1.AddItem ("ELECTRICAL")
Combo1.AddItem ("ELECTRONICS")
Combo1.AddItem ("CIVIL")
Combo1.AddItem ("MECHANICAL")
Combo1.AddItem ("AUTOMOBILE")
Me.Visible = True
Text1.SetFocus
End Sub

Private Sub Label5_Click()
MsgBox "please contact your admin for further procedure", vbExclamation, "HELP:FORGOT PASSWORD"
End Sub

Private Sub Option1_Click()
Label6.Visible = False
Combo1.Visible = False
End Sub

Private Sub Option2_Click()
Label6.Visible = True
Combo1.Visible = True
End Sub

Private Sub entry_log()
rs.Close
rs.Open "select * from log", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("Name").Value = Text1.Text
.Fields("Department").Value = Combo1.Text
.Fields("Date").Value = Date$
.Fields("Time").Value = Time$
End With
rs.Update
End Sub
Private Sub entry_log_admin()
rs.Close
rs.Open "select * from log", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("Name").Value = Text1.Text
.Fields("Date").Value = Date$
.Fields("Time").Value = Time$
End With
rs.Update
End Sub
