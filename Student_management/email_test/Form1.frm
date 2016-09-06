VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   4560
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your Gmail email address
    oSmtp.FromAddr = "engineertharun@gmail.com"
    
    ' Add recipient email address
    oSmtp.AddRecipientEx "engineertharun@gmail.com", 0
    
    ' Set email subject
    oSmtp.Subject = "test email from gmail account"
    
    ' Set email body
    oSmtp.BodyText = "this is a test email sent from VB 6.0 project with gmail"
    
    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.gmail.com"
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = 465
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.UserName = "engineertharun@gmail.com"
    oSmtp.Password = "hibuddyitsme"
    
    MsgBox "start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If
    
End Sub
