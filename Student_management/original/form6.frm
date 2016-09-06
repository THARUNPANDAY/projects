VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000012&
   ClientHeight    =   10950
   ClientLeft      =   -90
   ClientTop       =   240
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   7680
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
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   9480
      TabIndex        =   32
      Text            =   "Select Department"
      Top             =   3480
      Width           =   5055
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H80000012&
      Caption         =   "HOSTAL"
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
      Height          =   495
      Left            =   12480
      TabIndex        =   30
      Top             =   5640
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000012&
      Caption         =   "DAYSCHOLAR"
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
      Height          =   495
      Left            =   9360
      TabIndex        =   29
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   495
      Left            =   9360
      TabIndex        =   28
      Top             =   5640
      Width           =   5055
   End
   Begin VB.OptionButton Option10 
      BackColor       =   &H80000012&
      Caption         =   "OC"
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
      Left            =   13560
      TabIndex        =   27
      Top             =   4800
      Width           =   1695
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H80000012&
      Caption         =   "BC/MBC"
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
      Left            =   11280
      TabIndex        =   26
      Top             =   4800
      Width           =   2415
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H80000012&
      Caption         =   "SC/ST"
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
      Left            =   9360
      TabIndex        =   25
      Top             =   4800
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "GOVERMENT"
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
      Left            =   9360
      TabIndex        =   24
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   735
      Left            =   9360
      TabIndex        =   23
      Top             =   4800
      Width           =   5775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "MANAGEMENT"
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
      Left            =   12360
      TabIndex        =   22
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   735
      Left            =   9360
      TabIndex        =   21
      Top             =   3960
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
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
      Left            =   13680
      TabIndex        =   19
      Top             =   9840
      Width           =   1335
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
      Left            =   10800
      TabIndex        =   18
      Top             =   9840
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
      Left            =   7800
      TabIndex        =   17
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4920
      TabIndex        =   16
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      Left            =   9360
      TabIndex        =   15
      Top             =   8640
      Width           =   2415
   End
   Begin VB.TextBox Text5 
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
      Left            =   9360
      TabIndex        =   14
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text4 
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
      Left            =   9360
      TabIndex        =   13
      Top             =   7200
      Width           =   2415
   End
   Begin VB.TextBox Text3 
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
      Left            =   9360
      TabIndex        =   12
      Top             =   6480
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000012&
      Caption         =   "BC/MBC"
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
      Height          =   495
      Left            =   11280
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text2 
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
      Left            =   9480
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9480
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6360
      TabIndex        =   31
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label10 
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
      Left            =   6600
      TabIndex        =   20
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "HOSTEL FEE :"
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
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   " MODE OF TRANSPORT :"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "BUS FEE :"
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
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "BOOK FEE :"
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
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "TUTION FEE :"
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
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "CASTE :"
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
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "QUOTA :"
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
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "BATCH :"
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
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label rr 
      BackColor       =   &H80000012&
      Caption         =   "REGISTER NUMBER :"
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
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
