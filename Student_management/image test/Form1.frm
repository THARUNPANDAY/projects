VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   15480
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   4680
      ScaleHeight     =   2595
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub Form_Load()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "*.jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub
