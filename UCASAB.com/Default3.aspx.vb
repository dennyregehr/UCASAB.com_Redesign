
Partial Class Default3
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim evUtl As New EventUtil()
        EventDetailControl1.EventDetailClass = evUtl.GetEvent(402)
    End Sub
End Class
