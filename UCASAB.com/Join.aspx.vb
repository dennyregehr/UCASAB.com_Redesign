
Partial Class Join
    Inherits System.Web.UI.Page

    Protected Sub btnSabApplication_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSabApplication.Click
        FileManager.DownloadFile("2010_SAB_APP2.doc", "UCA_SAB_2010_Application.doc")
    End Sub
End Class
