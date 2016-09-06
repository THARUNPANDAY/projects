VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "FINAL YEAR UPDATION"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20175
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
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
      Left            =   16080
      TabIndex        =   12
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
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
      Left            =   16080
      TabIndex        =   11
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
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
      Left            =   16080
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHOW SHEET"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16080
      TabIndex        =   9
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   8280
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6120
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   4575
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   9840
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000008&
      Caption         =   $"Form4.frx":0000
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   14880
      TabIndex        =   13
      Top             =   8880
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "MEDICAL LEAVE REG. NUMBER:               ( LAST 4 DIGITS )"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3960
      TabIndex        =   8
      Top             =   8160
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "ON-DUTY REG. NUMBER :               ( LAST 4 DIGITS )"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4920
      TabIndex        =   7
      Top             =   6120
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "ABSENTEES REG. NUMBER :            ( LAST 4 DIGITS)"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4560
      TabIndex        =   6
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   7920
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ATTENDANCE - GENERATOR                      fourth year"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim exap As Excel.Application



Private Sub Command1_Click()
Dat = Text1.Text
rowend = exws.UsedRange.Rows.Count
col = exws.UsedRange.Columns.Count
col = col + 1
exws.Cells(1, col).Value = Dat
tsize = ((Len(Text2.Text) + 1) / 5)
tr = 1
For i = 1 To tsize
num = Mid(Text2.Text, tr, 4)
For r = 1 To rowend
 If exws.Cells(r, 1).Value = Val(num) Then
 exws.Cells(r, col).Value = " a "
 End If
  Next r
 tr = tr + 5
 Next i
 t3size = ((Len(Text3.Text) + 1) / 5)
 t3 = 1
 For x = 1 To t3size
 od = Mid(Text3.Text, t3, 4)
 For y = 1 To rowend
 If exws.Cells(y, 1).Value = Val(od) Then
 exws.Cells(y, col).Value = "OD"
 End If
 Next y
 t3 = t3 + 5
  Next x
  t4size = ((Len(Text4.Text) + 1) / 5)
 t4 = 1
 For p = 1 To t4size
 od = Mid(Text4.Text, t4, 4)
 For q = 1 To rowend
 If exws.Cells(q, 1).Value = Val(od) Then
 exws.Cells(q, col).Value = "ML"
 End If
 Next q
 t4 = t4 + 5
  Next p
 MsgBox ("OPERATION 101.05: CHANGES UPDATED SUCCESFULLY")
 'Dim sapi, tell
'tell = "absentees ,    onduty ,     medical  leave      students     updated"
'Set sapi = CreateObject("sapi.spvoice")
'sapi.Speak tell
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Command3_Click()
MsgBox ("WARNING 402.56 : DO NOT CLOSE THE EXCEL SHEET,MINIMIZE IT & SAVE IT")
'Dim sapi, tell
'tell = " WARNING : DO  NOT    CLOSE    THE    EXCEL    SHEET   ,   MINIMIZE    &   SAVE   It   "
'Set sapi = CreateObject("sapi.spvoice")
'sapi.Speak tell
exap.Visible = True

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\vb\New folder\finalyear.xlsx")
Set exws = exwb.Worksheets(1)
exap.Visible = False
End Sub



