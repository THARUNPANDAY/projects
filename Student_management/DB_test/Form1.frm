VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   15750
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4320
      Top             =   5160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   7800
      TabIndex        =   2
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
rs.AddNew
rs.Fields("name") = Text1.Text
rs.Fields("password") = Text2.Text
rs.Update
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\DB_test\ERP_try.mdb;Persist Security Info=false"
rs.Open "select * from ERP_try_add_update", con, adOpenDynamic, adLockPessimistic
End Sub
