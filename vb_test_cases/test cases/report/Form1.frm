VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20175
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   6720
      TabIndex        =   2
      Top             =   3960
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6840
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exwb As Excel.Workbook
Dim exws As Excel.Worksheet
Dim exap As Excel.Application
Private Sub Command1_Click()
dat = Val(Text1.Text)
rowend = exws.UsedRange.Rows.Count
col = exws.UsedRange.Columns.Count
Dim key As String
key = "a"
For r = 1 To rowend
 If exws.cells(r, 1).Value = dat Then
 For c = 2 To col
 hi = exws.cells(r, col).Value
 If hi = key Then
 Text2.Text = exws.cells(r, col).Value & ","
 End If
 Next c
 End If
  Next r
End Sub

Private Sub Form_Load()
Set exap = CreateObject("Excel.application")
Set exwb = exap.Workbooks.Open("D:\vb\New folder\thirdyear.xlsx")
Set exws = exwb.Worksheets(1)
exap.Visible = False
End Sub


