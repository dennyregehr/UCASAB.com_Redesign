
Partial Class UserControls_SocialMetaTags
    Inherits System.Web.UI.UserControl

    Private metaTagFb As String = "<meta property=""{0}"" content=""{1}"" />"
    'Private metaTagTwitter As String = "<meta name=""{0}"" content=""{1}"" />"

    Private _EvtDtl As EventDetails
    Public Property EventDetailClass() As EventDetails
        Get
            Return _EvtDtl
        End Get
        Set(ByVal value As EventDetails)
            _EvtDtl = value
            LoadEventDetailsFb()
        End Set
    End Property

    Private Sub LoadEventDetailsFb()
        With PlaceHolder1.Controls
            .Add(CreateCtrl(metaTagFb, "og:title", String.Concat("UCA SAB - ", _EvtDtl.EventName)))
            Dim shortDesc As String
            If _EvtDtl.EventDescription.Length > 500 Then
                shortDesc = _EvtDtl.EventDescription.Substring(0, 500)
            Else
                shortDesc = _EvtDtl.EventDescription
            End If
            If shortDesc.Length = 0 Then
                shortDesc = _EvtDtl.EventName
            End If
            .Add(CreateCtrl(metaTagFb, "og:description", String.Concat("UCA SAB ", _EvtDtl.EventTypeName, " - ", Server.HtmlEncode(shortDesc))))
            .Add(CreateCtrl(metaTagFb, "og:type", "university"))
            .Add(CreateCtrl(metaTagFb, "og:image_src", ResolveImageUrl(_EvtDtl.ImageURL)))
            .Add(CreateCtrl(metaTagFb, "og:image", ResolveImageUrl(_EvtDtl.ImageURL)))
            .Add(CreateCtrl(metaTagFb, "og:url", String.Concat(Request.Url.GetLeftPart(UriPartial.Authority), Page.ResolveUrl(_EvtDtl.DetailURL))))
            .Add(CreateCtrl(metaTagFb, "og:site_name", "University of Central Arkansas Student Activities Board"))
        End With
    End Sub

    Private Function ResolveImageUrl(ImageUrl As String) As String
        Select Case ImageUrl.Substring(0, 1)
            Case "/"
                Return Request.Url.GetLeftPart(UriPartial.Authority) & ImageUrl
            Case "~"
                Return Request.Url.GetLeftPart(UriPartial.Authority) & Page.ResolveUrl(ImageUrl)
            Case Else
                Return ImageUrl
        End Select
    End Function

    Private Function CreateCtrl(MetaTag As String, Descriptor As String, Description As String) As LiteralControl
        Return New LiteralControl(String.Format(MetaTag, Descriptor, Description) & vbCrLf)
    End Function
    '<meta property="og:title" content="UCA SAB - "/>
    '<meta property="" content=""/>
    '<meta property="" content=""/>
    '<meta property="" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    '<meta property="" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    '<meta property="" content="http://www.ucasab.com/novelty.aspx"/>
    '<meta property="" content=""/>

End Class
