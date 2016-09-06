VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H80000007&
   Caption         =   "Form8"
   ClientHeight    =   10560
   ClientLeft      =   315
   ClientTop       =   645
   ClientWidth     =   20025
   LinkTopic       =   "Form8"
   ScaleHeight     =   10560
   ScaleWidth      =   20025
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1080
      Top             =   5160
      Visible         =   0   'False
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
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      TabIndex        =   7
      Text            =   "30"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7920
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9360
      TabIndex        =   4
      Top             =   2760
      Width           =   6495
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "NUMBER OF WORKING DAYS THIS MONTH :"
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
      Left            =   5520
      TabIndex        =   6
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "ABSENTEES :             ( REGISTER NUMBER )"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "SELECT THE STARTING YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   4935
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
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exap As Excel.Application
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub Command1_Click()
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\ERP\original\db\admin.mdb;Persist Security Info=False")
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\Projects\ERP\original\db\attendance.xlsx")
Set exws = exwb.Worksheets(List1.Text)


''insert a to absentees
rowend = exws.UsedRange.Rows.Count
tsize = ((Len(Text1.Text) + 1) / 13)
tr = 1
For i = 1 To tsize
td = Mid(Text1.Text, tr, 12)
For j = 2 To rowend
exws.Cells(j, (Month(Now) + 2)).Value = ""
If exws.Cells(j, 1).Value = Val(td) Then
exws.Cells(j, (Month(Now) + 1)).Value = exws.Cells(j, (Month(Now) + 1)).Value + "a"
rs.Open "select * from cse_attendance where register_number='" + td + "'", con, adOpenDynamic, adLockPessimistic
If (Month(Now)) = 1 Then
If rs.EOF Then
Else
rs.Fields("january").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 2 Then
If rs.EOF Then
Else
rs.Fields("february").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 3 Then
If rs.EOF Then
Else
rs.Fields("march").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 4 Then
If rs.EOF Then
Else
rs.Fields("april").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 5 Then
If rs.EOF Then
Else
rs.Fields("may").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 6 Then
If rs.EOF Then
Else
rs.Fields("june").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 7 Then
If rs.EOF Then
Else
rs.Fields("july").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 8 Then
If rs.EOF Then
Else
rs.Fields("august").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 9 Then
If rs.EOF Then
Else
rs.Fields("september").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 10 Then
If rs.EOF Then
Else
rs.Fields("october").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 11 Then
If rs.EOF Then
Else
rs.Fields("november").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
ElseIf (Month(Now)) = 12 Then
If rs.EOF Then
Else
rs.Fields("december").Value = Len(exws.Cells(j, (Month(Now) + 1)).Value)
rs.Update
rs.Close
End If
End If
End If
Next j
tr = tr + 13
Next i



'' add total number of days
rs.Open "select * from cse_attendance ", con, adOpenDynamic, adLockPessimistic
For i = 1 To rowend
If (Month(Now)) = 1 Then
If rs.EOF Then
Else
rs.Fields("jan_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 2 Then
If rs.EOF Then
Else
rs.Fields("feb_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 3 Then
If rs.EOF Then
Else
rs.Fields("mar_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 4 Then
If rs.EOF Then
Else
rs.Fields("apr_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 5 Then
If rs.EOF Then
Else
rs.Fields("may_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 6 Then
If rs.EOF Then
Else
rs.Fields("jun_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 7 Then
If rs.EOF Then
Else
rs.Fields("jul_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 8 Then
If rs.EOF Then
Else
rs.Fields("aug_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 9 Then
If rs.EOF Then
Else
rs.Fields("sep_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 10 Then
If rs.EOF Then
Else
rs.Fields("oct_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 11 Then
If rs.EOF Then
Else
rs.Fields("nov_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now)) = 12 Then
If rs.EOF Then
Else
rs.Fields("dec_total").Value = Text4.Text
rs.Update
rs.MoveNext
End If
End If
Next i
rs.Close


''make forward zero
If Day(Now) = 1 Or 2 Or 3 Then
rs.Open "select * from cse_attendance", con, adOpenDynamic, adLockPessimistic
For i = 1 To rowend
If (Month(Now) + 1) = 13 Then
If rs.EOF Then
Else
rs.Fields("january").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 2 Then
If rs.EOF Then
Else
rs.Fields("february").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 3 Then
If rs.EOF Then
Else
rs.Fields("march").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 4 Then
If rs.EOF Then
Else
rs.Fields("april").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 5 Then
If rs.EOF Then
Else
rs.Fields("may").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 6 Then
If rs.EOF Then
Else
rs.Fields("june").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 7 Then
If rs.EOF Then
Else
rs.Fields("july").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 8 Then
If rs.EOF Then
Else
rs.Fields("august").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 9 Then
If rs.EOF Then
Else
rs.Fields("september").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 10 Then
If rs.EOF Then
Else
rs.Fields("october").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 11 Then
If rs.EOF Then
Else
rs.Fields("november").Value = 0
rs.Update
rs.MoveNext
End If
ElseIf (Month(Now) + 1) = 12 Then
If rs.EOF Then
Else
rs.Fields("december").Value = 0
rs.Update
rs.MoveNext
End If
End If
Next i
rs.Close
End If
con.Close
exwb.Close
exap.Quit
End Sub

Private Sub Form_Load()
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\Projects\ERP\original\db\attendance.xlsx")
For i = 1 To exwb.Worksheets.Count
List1.AddItem (exwb.Worksheets(i).Name)
Next i
exwb.Close
exap.Quit
End Sub

