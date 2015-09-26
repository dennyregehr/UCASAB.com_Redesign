
Partial Class Contact
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            LoadRecipients()
        End If
    End Sub

    Private Sub LoadRecipients()
        Dim execMbrUtil As New ExecMemberUtil
        With lstRecipients
            .DataTextField = "ExecMember"
            .DataValueField = "Position"
            .DataSource = execMbrUtil.GetExecMembers
            .DataBind()
        End With
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            ccJoin.ValidateCaptcha(txtCaptchaResponse.Text)
            If ccJoin.UserValidated Then
                SendEmail()
                pnlMessage.Visible = True
                pnlForm.Visible = False
            Else
                lblCaptchaError.Visible = True
            End If
        Catch ex As Exception
            pnlErrorMessage.Visible = True
            lblErrorMessage.Text = ex.Message
            pnlForm.Visible = False
            pnlMessage.Visible = False
        End Try
    End Sub

    Private Sub SendEmail()
        Dim senderMail As String = ConfigurationManager.AppSettings("WebsiteEmail")
        Dim recipientMail As New List(Of String)
        recipientMail.Add(GetEmailAddress())
        Dim messageBody As New StringBuilder
        messageBody.AppendFormat("Message from: {0}", txtName.Text).Append(vbCrLf)
        messageBody.AppendFormat("Email: {0}", txtEmail.Text).Append(vbCrLf)
        messageBody.Append("Message:").Append(vbCrLf)
        messageBody.Append(txtBody.Text)
        Dim myEmail As New EmailUtil
        myEmail.Send(senderMail, txtName.Text, recipientMail, ConfigurationManager.AppSettings("WebsiteEmail_Subject"), messageBody.ToString)
    End Sub

    Private Function GetEmailAddress() As String
        Dim execMbrUtil As New ExecMemberUtil
        Return execMbrUtil.GetExecMembers.Find(Function(m) m.Position = lstRecipients.SelectedValue).Email
    End Function

    Protected Sub lstRecipients_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstRecipients.DataBound

        Dim recipId As String
        recipId = Request.QueryString("ebm")
        If recipId IsNot Nothing Then
            lstRecipients.SelectedValue = recipId
            If lstRecipients.SelectedValue <> recipId Then
                lstRecipients.Items.Clear()
                lstRecipients.Items.Add(New ListItem("Webmaster", "10"))
            End If
        End If

    End Sub

End Class
