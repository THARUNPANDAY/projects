VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   120
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9000
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=BanK"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "BanK"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bankoffice"
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
      Caption         =   "Clear"
      Height          =   495
      Left            =   9480
      TabIndex        =   22
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "DEPOSITS"
      Height          =   855
      Left            =   6600
      TabIndex        =   21
      Top             =   10200
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "WITHDRAWLS"
      Height          =   855
      Left            =   6600
      TabIndex        =   20
      Top             =   8880
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
      Height          =   855
      Left            =   6480
      TabIndex        =   19
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "END"
      Height          =   975
      Left            =   6600
      TabIndex        =   18
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT"
      Height          =   735
      Left            =   6360
      TabIndex        =   17
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREVIOUS"
      Height          =   735
      Left            =   6360
      TabIndex        =   16
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      Height          =   735
      Left            =   6480
      TabIndex        =   15
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   735
      Left            =   6240
      TabIndex        =   14
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "balance"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   3360
      TabIndex        =   13
      Top             =   9360
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   855
      Left            =   3360
      TabIndex        =   12
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   3120
      TabIndex        =   10
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   3000
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "balance"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "acc type"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "des"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "add"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "AGE"
      DataField       =   "AGE"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "c name"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "id"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rss As New ADODB.Recordset

Private Sub cmdadd_Click()
rss(0) = Text1.Text
rss(1) = Text2.Text
rss(2) = Text3.Text
rss(3) = Text4.Text
rss(4) = Text5.Text
rss(5) = Text6.Text
rss(6) = Text7.Text
rss.AddNew
MsgBox "Bank Saved"
End Sub

Private Sub cmdupdate_Click()
rss.Update
End Sub

Private Sub Command2_Click()
rss.search
End Sub

Private Sub Command3_Click()
rss.MovePrevious
End Sub

Private Sub Command4_Click()
rss.MoveNext
End Sub

Private Sub Command5_Click()
Close
End Sub

Private Sub Command6_Click()
rss.Delete "select * from bankoffice where id=2"
End Sub

Private Sub Form_Load()
con.Open "dsn=BanK"
rss.Open "select * from bankoffice", con, adOpenKeyset, adLockOptimistic
End Sub
