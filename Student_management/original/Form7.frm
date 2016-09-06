VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BackColor       =   &H80000007&
   Caption         =   "Form7"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form7"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD "
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
      Left            =   8520
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE ATTENDANCE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD NEW BATCH"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "ENTER THE YEAR OF JOINING :"
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
      Height          =   975
      Left            =   6720
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "ENTER THE ENDING REGISTER NUMBER :"
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
      Height          =   855
      Left            =   6240
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "ENTER THE STARTING REGISTER NUMBER :"
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
      Height          =   735
      Left            =   6120
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   3735
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
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim rs As New ADODB.Recordset
Dim exap As Excel.Application
Dim con As New ADODB.Connection

Private Sub Command1_Click()
Label2.Visible = True
Label3.Visible = True
Text1.Visible = True
Text2.Visible = True
Command3.Visible = True
Text3.Visible = True
Label4.Visible = True
End Sub

Private Sub Command2_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\Projects\ERP\original\db\attendance.xlsx")
j = 2
exwb.Worksheets.Add().Name = Text3.Text
Set exws = exwb.Worksheets(Text3.Text)
If Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 104 Then
rs.Open "select * from cse_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
ElseIf Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 106 Then
rs.Open "select * from ece_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
ElseIf Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 103 Then
rs.Open "select * from civil_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
ElseIf Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 102 Then
rs.Open "select * from auto_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
ElseIf Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 105 Then
rs.Open "select * from eee_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
ElseIf Mid(Text1.Text, 7, 3) = Mid(Text2.Text, 7, 3) And Mid(Text1.Text, 7, 3) = 114 Then
rs.Open "select * from mech_attendance", con, adOpenDynamic, adLockPessimistic
For i = Text1.Text To Text2.Text
exws.Cells(j, 1).Value = Val(i)
rs.AddNew
rs.Fields("register_number").Value = Val(i)
rs.Fields("january") = 0
rs.Fields("february") = 0
rs.Fields("march") = 0
rs.Fields("april") = 0
rs.Fields("may") = 0
rs.Fields("june") = 0
rs.Fields("july") = 0
rs.Fields("august") = 0
rs.Fields("september") = 0
rs.Fields("october") = 0
rs.Fields("november") = 0
rs.Fields("december") = 0
rs.Update
j = j + 1
Next i
exws.Cells(1, 2).Value = "january"
exws.Cells(1, 3).Value = "february"
exws.Cells(1, 4).Value = "march"
exws.Cells(1, 5).Value = "april"
exws.Cells(1, 6).Value = "may"
exws.Cells(1, 7).Value = "june"
exws.Cells(1, 8).Value = "july"
exws.Cells(1, 9).Value = "august"
exws.Cells(1, 10).Value = "september"
exws.Cells(1, 11).Value = "october"
exws.Cells(1, 12).Value = "november"
exws.Cells(1, 13).Value = "december"
exwb.Close
exap.Quit
rs.Close
End If
con.Close
End Sub


