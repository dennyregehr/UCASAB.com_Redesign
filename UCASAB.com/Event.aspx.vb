
Partial Class SABEvent
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not String.IsNullOrEmpty(Request("id")) Then
            Dim eventId = Request("id")
            If IsNumeric(eventId) Then
                Dim evUtl As New EventUtil()
                Dim ev As EventDetails = evUtl.GetEvent(eventId)
                lblEventName.Text = Server.UrlDecode(ev.EventTypeName) & " Event"
                EventDetailControl1.EventDetailClass = ev
                SocialMetaTags1.EventDetailClass = ev
            Else
                Throw New HttpRequestValidationException()
            End If
        End If
    End Sub
End Class
