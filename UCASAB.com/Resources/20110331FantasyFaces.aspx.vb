Imports System.IO

Partial Class Resources_20110331FantasyFaces
    Inherits System.Web.UI.Page

    Private SourceFolder As String = "~/Resources/20110331FantasyFaces/thumbs/"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fs = FileManager.GetFiles(Server.MapPath(SourceFolder))
        For Each f In fs
            Dim link As New HyperLink()
            With link
                .ImageUrl = String.Concat(SourceFolder, Path.GetFileName(f))
                .NavigateUrl = .ImageUrl.Replace("/thumbs", "").Replace(".png", ".jpg")
                .Attributes.Add("style", "padding:10px")
            End With
            PlaceHolder1.Controls.Add(link)
        Next
    End Sub
End Class
