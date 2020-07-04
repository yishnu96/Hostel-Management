VERSION 5.00
Begin VB.Form sendmail 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SEND MAIL"
      Height          =   1095
      Left            =   4440
      TabIndex        =   0
      Top             =   4320
      Width           =   2295
   End
End
Attribute VB_Name = "sendmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
    Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"

    ' Set your Gmail email address
    oSmtp.FromAddr = "yishnu14pramanik@gmail.com"

    ' Add recipient email address
    oSmtp.AddRecipientEx "rohanmdk1998@gmail.com", 0

    ' Set email subject
    oSmtp.Subject = "test email from gmail account"

    ' Set email body
    oSmtp.BodyText = "this is a test email sent from VB 6.0 project with gmail"

    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.gmail.com"

    ' If you want to use direct SSL 465 port,
    ' Please add this line, otherwise TLS will be used.
    ' oSmtp.ServerPort = 465

    ' set 25 or 587 port
    oSmtp.ServerPort = 587

    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.UserName = "yishnu14pramanik@gmail.com"
    oSmtp.Password = "yourpassword"

    MsgBox "start to send email ..."

    If oSmtp.sendmail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If
End Sub

