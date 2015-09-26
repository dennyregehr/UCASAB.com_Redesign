Imports System.Xml

Partial Class rss
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CreateRSS()
    End Sub

    Private Sub CreateRSS()
        Response.Clear()
        Response.ContentType = "text/xml"
        Dim tempUri = New UriBuilder(Request.Url)
        Dim tempPort As String = IIf(tempUri.Port = 80, "", ":" & tempUri.Port)
        Dim objX As New XmlTextWriter(Response.OutputStream, Encoding.UTF8)
        With objX
            .WriteStartDocument()
            .WriteStartElement("rss")
            .WriteAttributeString("xmlns:Media", "http://search.yahoo.com/mrss/")
            .WriteAttributeString("version", "2.0")
            .WriteStartElement("channel")
            .WriteElementString("title", "UCA's Upcoming Student Activities")
            .WriteElementString("link", "http://www.ucasab.com")
            .WriteElementString("description", "Don't miss the upcoming Student Activities events at the University of Central Arkansas.")
            .WriteElementString("copyright", "")
            .WriteElementString("ttl", "120 ")
            .WriteStartElement("image")
            .WriteElementString("title", "UCA Student Activities Board")
            .WriteElementString("link", "http://www.ucasab.com")
            'build url
            Dim imageUrl As String
            imageUrl = String.Format("{0}://{1}{2}{3}", tempUri.Scheme, Request.ServerVariables("server_name"), tempPort, ResolveUrl("~/images/sabtemplogo1.png"))
            .WriteElementString("url", imageUrl)
            .WriteElementString("width", "175")
            .WriteElementString("height", "63")
            .WriteElementString("description", "Student Activities events at the University of Central Arkansas")
            .WriteEndElement()

            Dim numberOfItems As Int16 = ConfigurationManager.AppSettings("RSSNumberOfItemsToDisplay")
            If numberOfItems = 0 Then numberOfItems = 10

            Dim myEvts As New EventUtil
            For Each evt As EventDetails In myEvts.GetEventsForRssFeed(numberOfItems)
                .WriteStartElement("item")
                .WriteElementString("title", String.Format("{0} - {1}", CDate(evt.StartDate).ToShortDateString, evt.EventName))
                Dim desc As String
                If evt.EventDescription.Trim = String.Empty Then
                    desc = " - Join us for the fun!"
                Else
                    desc = evt.EventDescription
                End If
                .WriteElementString("description", String.Format("{0} @ {1} - {2}", evt.StartTime, evt.Location, desc))

                'build link
                Dim newUrl As String
                newUrl = String.Format("{0}://{1}{2}{3}", tempUri.Scheme, Request.ServerVariables("server_name"), tempPort, ResolveUrl(evt.DetailURL))
                .WriteElementString("link", newUrl)

                .WriteElementString("pubDate", CDate(evt.StartDate).AddHours(6).ToString("r"))
                .WriteEndElement()
            Next

            .WriteEndElement()
            .WriteEndElement()
            .WriteEndDocument()
            .Flush()
            .Close()
        End With

    End Sub

End Class
