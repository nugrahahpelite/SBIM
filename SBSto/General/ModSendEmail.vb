Imports System.Data.SqlClient
Imports System.Net.Mail
Imports GlobalUtil

Module ModSendEmail
    Public Function fbuSendEMail(vriTextStream As Scripting.TextStream, vriLogFileName As String, vriLogFileNameError As String, vriMailTo01 As String, vriMailTo02 As String, vriSubject As String, vriMessage As String, vriAttachment1 As String, vriAttachment2 As String) As Boolean
        fbuSendEMail = False
        vriTextStream.WriteLine("-----------------------------Sending Email Start---------------------------------------")
        vriTextStream.WriteLine("")

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("S.1")

        Dim vnSending As Boolean

        Dim vnMailFrom As String = ""
        Dim vnMailClientHost As String = ""
        Dim vnMailFromPwd As String = ""

        vnMailClientHost = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("MailClientHost"), "MyEncryptPassword")
        vnMailFrom = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("MailFrom"), "MyEncryptPassword")
        vnMailFromPwd = EncryptDecrypt.Decrypt(ConfigurationManager.AppSettings("MailFromPwd"), "MyEncryptPassword")

        vriTextStream.WriteLine("")
        vriTextStream.WriteLine("S.2")

        Dim vnLogPos As String = "0"
        Try
            vnSending = False
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("Sending Email Start " & Date.Now)
            Dim vnMailMessage As New MailMessage
            Dim vnSmtpClient As New SmtpClient

            vnLogPos = "vnLogPos 1"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnMailMessage.To.Add(New MailAddress(vriMailTo01))
            If vriMailTo02 <> "" Then
                vnMailMessage.To.Add(New MailAddress(vriMailTo02))
            End If

            vnLogPos = "vnLogPos 2"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnMailMessage.IsBodyHtml = True

            vnLogPos = "vnLogPos 3"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnMailMessage.From = (New MailAddress(vnMailFrom))
            vnMailMessage.Subject = vriSubject
            vnMailMessage.Body = vriMessage

            vnLogPos = "vnLogPos 4"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnSmtpClient.Host = vnMailClientHost

            vnLogPos = "vnLogPos 5"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            'vnSmtpClient.Port = 465
            vnSmtpClient.EnableSsl = True

            vnLogPos = "vnLogPos 6"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnSmtpClient.UseDefaultCredentials = False

            vnLogPos = "vnLogPos 7"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnSmtpClient.Credentials = New System.Net.NetworkCredential(vnMailFrom, vnMailFromPwd)

            If vriAttachment1 <> "" Then
                vnMailMessage.Attachments.Add(New Attachment(vriAttachment1))
            End If
            If vriAttachment2 <> "" Then
                vnMailMessage.Attachments.Add(New Attachment(vriAttachment2))
            End If

            vnLogPos = "vnLogPos 8"
            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("vnLogPos " & vnLogPos)

            vnSending = True
            vnSmtpClient.Send(vnMailMessage)

            vriTextStream.WriteLine("Sending Email End " & Date.Now)
            vriTextStream.WriteLine("")

            vriTextStream.WriteLine("")
            vriTextStream.WriteLine("-----------------------------End Of Sending Email---------------------------------------")

            fbuSendEMail = True
        Catch ex As Exception
            pbMsgError = ex.Message
            If vnSending Then
                FileCopy(vriLogFileName, vriLogFileNameError)
                vriTextStream.WriteLine("ERROR WHILE SENDING EMAIL ** vnSmtpClient.Send(vnMailMessage) **")
                vriTextStream.WriteLine(pbMsgError)
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("-----------------------------End Of Sending Email---------------------------------------")
            Else
                vriTextStream.WriteLine("")
                vriTextStream.WriteLine("ERROR While Sending Email " & Date.Now)
                vriTextStream.WriteLine(pbMsgError)
                vriTextStream.WriteLine("-----------------------------End Of Sending Email---------------------------------------")
                FileCopy(vriLogFileName, vriLogFileNameError)
            End If
        End Try
    End Function
End Module
