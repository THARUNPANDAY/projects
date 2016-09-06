VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   Caption         =   "Form9"
   ClientHeight    =   10950
   ClientLeft      =   -90
   ClientTop       =   240
   ClientWidth     =   20250
   LinkTopic       =   "Form9"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check9 
      BackColor       =   &H80000012&
      Caption         =   "BAD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12360
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H80000012&
      Caption         =   "FAIR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10560
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H80000012&
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8760
      TabIndex        =   17
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000012&
      Caption         =   "BAD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12360
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H80000012&
      Caption         =   "FAIR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10560
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000012&
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000012&
      Caption         =   "BAD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12360
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000012&
      Caption         =   "FAIR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10560
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000012&
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   8520
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   6240
      Width           =   4815
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
      Left            =   10320
      TabIndex        =   3
      Text            =   "SELECT DEPARTMENT"
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "ACTIVITIES :"
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
      Left            =   6240
      TabIndex        =   10
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "STUDIES :"
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
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "DICIPLINE :"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Caption         =   "REASON :"
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
      Left            =   6960
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   7080
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
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
      Top             =   1200
      Width           =   3855
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
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
rs.Open "select * from cse_feedback", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
If Check1.Enabled = True Then
.Fields("dicipline").Value = Check1.Caption
ElseIf Check2.Enabled = True Then
.Fields("dicipline").Value = Check2.Caption
ElseIf Check3.Enabled = True Then
.Fields("dicipline").Value = Check3.Caption
End If
If Check4.Enabled = True Then
.Fields("studies").Value = Check4.Caption
ElseIf Check5.Enabled = True Then
.Fields("studies").Value = Check5.Caption
ElseIf Check6.Enabled = True Then
.Fields("studies").Value = Check6.Caption
End If
If Check7.Enabled = True Then
.Fields("activities").Value = Check7.Caption
ElseIf Check8.Enabled = True Then
.Fields("activities").Value = Check8.Caption
ElseIf Check9.Enabled = True Then
.Fields("activities").Value = Check9.Caption
End If
.Fields("reason").Value = Text2.Text
End With
rs.Update
rs.Close
con.Close
End Sub

Private Sub Form_Load()
Form9.Visible = True
With Combo1
.AddItem ("COMPUTER SCIENCE")
.AddItem ("ELECTRONICS")
.AddItem ("ELECTRICAL")
.AddItem ("CIVIL")
.AddItem ("MECHANICAL")
.AddItem ("AUTOMOBILE")
End With
End Sub
