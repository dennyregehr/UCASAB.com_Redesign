Imports System.Net.Mail

Public Class EmailUtil

    Public Sub Send(ByVal MailSender As String, ByVal MailSenderName As String, ByVal MailRecipient As List(Of String), ByVal MailSubject As String, ByVal MailBody As String)
        Dim myMail As New MailMessage()
        myMail.IsBodyHtml = False
        myMail.From = New MailAddress(MailSender, MailSenderName)
        For Each eml As String In MailRecipient
            myMail.To.Add(eml)
        Next
        If ConfigurationManager.AppSettings("WebsiteAdmin_Email") IsNot Nothing Then
            myMail.Bcc.Add(ConfigurationManager.AppSettings("WebsiteAdmin_Email"))
        End If
        myMail.Subject = MailSubject	' & AddIdentityInfo()
        myMail.Body = MailBody & AddIdentityInfo(myMail.IsBodyHtml)
        Dim smtp As New SmtpClient()
        smtp.Send(myMail)
    End Sub

    Private Function AddIdentityInfo(IsHTML as Boolean) As String
        Dim txt As New StringBuilder(String.Empty)
		Dim msg As String
		msg = "This email is part of the regular communications coming from the University of Central Arkansas' Student Activities Board."
		If IsHTML Then
			txt.AppendFormat("<div style=""clear:both;"">{0}</div>", msg)
		Else
			txt.Append(vbCrLf).Append(msg)
		End If
        Return txt.ToString
    End Function

End Class
