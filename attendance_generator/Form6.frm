VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000008&
   Caption         =   "SETUP--( ENTER KEY )"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20175
   BeginProperty Font 
      Name            =   "Modern No. 20"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   10950
   ScaleWidth      =   20175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT KEY"
      Height          =   615
      Left            =   9240
      TabIndex        =   3
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   2
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "ENTER KEY :"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6480
      TabIndex        =   1
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ATTENDANCE - GENERATOR        ADMINISTRATION  LOGIN"
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
      Height          =   1935
      Left            =   5520
      TabIndex        =   0
      Top             =   480
      Width           =   10215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim exap As Excel.Application
Private Sub Command1_Click()
Dim a, b As String
c = Text1.Text
If c = "viralbrothers" Then
MsgBox "THE KEY ENTERED IS ACCEPTED "
exws.Cells(2, 1).Value = c
a = DateValue(Now)
exws.Cells(4, 2).Value = a
b = DateAdd("m", 3, a)
exws.Cells(4, 1).Value = Mid(b, 1, 9)
Form1.Show
exap.Visible = True
exwb.SaveAs ("D:\vb\New folder\DB\dont use\dontwork.xls")
exap.Workbooks.Close
exap.Application.Quit
Else
MsgBox "PLEASE TYPE A VALID KEY "
Form6.Show
End If
End Sub

Private Sub Form_Load()
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\vb\New folder\DB\dont use\dontwork.xls", , False)
Set exws = exwb.Worksheets(1)
exap.Visible = False
End Sub
