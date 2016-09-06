VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000008&
   Caption         =   "SETUP FOR ATTENDANCE GENERATOR"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20175
   LinkTopic       =   "Form5"
   ScaleHeight     =   10950
   ScaleWidth      =   20175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton end 
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   3
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ENTER KEY"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   2
      Top             =   4560
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "USE TRAIL VERSION (10-USAGES)"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   1
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ATTENDANCE - GENERATOR "
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
      Height          =   1575
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   9855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim exap As Excel.Application
Private Sub Command1_Click()
Dim i, j As Integer
i = 1
j = 1
If exws.Cells(1, 1).Value = "0" Then
exws.Cells(1, 1).Value = i
z = 10
exws.Cells(1, 2).Value = z
exap.Visible = False
MsgBox "PLEASE PRESS YES IN THE UPCOMING DIALOGUE BOX, TO USE TRIAL VERSION CORRECTLY"
exwb.SaveAs ("D:\vb\New folder\DB\dont use\dontwork.xls")
exap.Workbooks.Close
exap.Application.Quit
Form1.Show
Else
 If exws.Cells(1, 2).Value = exws.Cells(3, 1).Value Then
  MsgBox "YOUR TRIAL VERSION HAS BEEN EXPIRED. ENTER KEY FOR FURTHER USAGE "
  exap.Workbooks.Close
  exap.Application.Quit
  Form6.Show
  Else
  i = exws.Cells(1, 1).Value
  exws.Cells(3, 1).Value = i + 1
  exws.Cells(1, 1).Value = i + 1
  exap.Visible = False
  MsgBox "PLEASE PRESS YES IN THE UPCOMING DIALOGUE BOX, TO USE TRIAL VERSION CORRECTLY"
  exwb.SaveAs ("D:\vb\New folder\DB\dont use\dontwork.xls")
  exap.Workbooks.Close
  exap.Application.Quit
  Form1.Show
  End If
  End If
End Sub


Private Sub Command2_Click()
exap.Workbooks.Close
exap.Application.Quit
Form6.Show
End Sub

Private Sub end_Click()
End
End Sub

Private Sub Form_Load()
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\vb\New folder\DB\dont use\dontwork.xls", , False)
Set exws = exwb.Worksheets(1)
exap.Visible = False
End Sub
