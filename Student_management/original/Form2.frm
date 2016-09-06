VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20190
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20190
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   27
      Top             =   7680
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9480
      TabIndex        =   25
      Top             =   7080
      Width           =   4455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SHOW"
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
      Left            =   12120
      TabIndex        =   23
      Top             =   9240
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16080
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9480
      TabIndex        =   22
      Text            =   "Select Department"
      Top             =   4920
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "FIND"
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
      Left            =   10080
      TabIndex        =   21
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UPDATE"
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
      Left            =   12120
      TabIndex        =   20
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
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
      Left            =   8160
      TabIndex        =   19
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
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
      Left            =   10080
      TabIndex        =   18
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD NEW"
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
      Left            =   7680
      TabIndex        =   17
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD PHOTO"
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
      Left            =   15000
      TabIndex        =   16
      Top             =   4200
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2880
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   15120
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   15
      Top             =   1080
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "FEMALE"
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
      Left            =   12000
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "MALE"
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
      Left            =   9480
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Top             =   3240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   44105729
      CurrentDate     =   42525
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
      Left            =   9480
      TabIndex        =   11
      Top             =   6360
      Width           =   4455
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
      Left            =   9480
      TabIndex        =   10
      Top             =   5640
      Width           =   4455
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
      TabIndex        =   9
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7200
      TabIndex        =   26
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
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
      Left            =   7080
      TabIndex        =   24
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "PHONE NO. :"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "ADDRESS :"
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
      TabIndex        =   6
      Top             =   5640
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "GENDER :"
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
      Left            =   6840
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "DOB :"
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
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   5040
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "NAME :"
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
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
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
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Command1_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "*.jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Command2_Click()
clear
Me.Visible = True
Text1.SetFocus
End Sub

Private Sub Command3_Click()
If Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
.Fields("email").Value = Text6.Text
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("email").Value = Text6.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
ElseIf Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
ElseIf Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
ElseIf Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_student_info", con, adOpenDynamic, adLockPessimistic
rs.AddNew
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Added Successfully !!", vbInformation, "Add/Save Successfull"
rs.Update
rs.Close
End If
End Sub

Private Sub Command4_Click()
confirm = MsgBox("Do You Wish To Delete The Record ?", vbYesNo + vbCritical, "Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record Deleted successfully !!", vbInformation, "Message "
rs.Update
End If
End Sub

Private Sub Command5_Click()
If Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
ElseIf Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
ElseIf Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
ElseIf Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_student_info", con, adOpenDynamic, adLockPessimistic
With rs
.Fields("register_number").Value = Text1.Text
.Fields("name").Value = Text2.Text
.Fields("dob").Value = DTPicker1.Value
If Option1.Enabled = True Then
.Fields("gender").Value = "male"
ElseIf Option2.Enabled = True Then
.Fields("gender").Value = "female"
End If
.Fields("department").Value = Combo1.Text
.Fields("address").Value = Text3.Text
.Fields("phone_no").Value = Text4.Text
.Fields("photo").Value = CommonDialog1.FileName
.Fields("batch").Value = Text5.Text
.Fields("email").Value = Text6.Text
End With
MsgBox "Data Has Been Updated Successfully !!", vbInformation, "Updated"
rs.Update
rs.Close
End If
End Sub

Private Sub Command6_Click()
Label3.Caption = "REGISTER NUMBER* :"
Label6.Caption = "DEPARTMENT* :"
If Combo1.Text = "COMPUTER SCIENCE" Then
rs.Open "select * from cse_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
Text2.Text = rs.Fields("name").Value
Text3.Text = rs.Fields("address").Value
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRICAL" Then
rs.Open "select * from eee_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
rs.Update
rs.Close
ElseIf Combo1.Text = "ELECTRONICS" Then
rs.Open "select * from ece_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
rs.Update
rs.Close
ElseIf Combo1.Text = "CIVIL" Then
rs.Open "select * from civil_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
rs.Update
rs.Close
ElseIf Combo1.Text = "MECHANICAL" Then
rs.Open "select * from mech_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
rs.Update
rs.Close
ElseIf Combo1.Text = "AUTOMOBILE" Then
rs.Open "select * from auto_student_info where register_number='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
rs.Update
rs.Close
End If
End Sub

Private Sub Command7_Click()
Form3.Show
End Sub

Private Sub Form_Load()
Text4.MaxLength = 10
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
Combo1.AddItem ("COMPUTER SCIENCE")
Combo1.AddItem ("ELECTRICAL")
Combo1.AddItem ("ELECTRONICS")
Combo1.AddItem ("CIVIL")
Combo1.AddItem ("MECHANICAL")
Combo1.AddItem ("AUTOMOBILE")
Me.Visible = True
Text1.SetFocus
End Sub

Private Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = "select department"
End Sub

