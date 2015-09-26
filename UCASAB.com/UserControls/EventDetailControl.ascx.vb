
Partial Class UserControls_EventDetailControl
    Inherits System.Web.UI.UserControl

    'Private _CommitteePage As String = "~/Committee.aspx?cmt={0}#evt{1}"
    Private _EventDescriptionMaxLength As Integer = 350
    Private _EvtDtl As EventDetails
    'Private _MovieImageWidth As String = ConfigurationManager.AppSettings("MovieImageWidthOnHomePage")
    Public Property EventDetailClass() As EventDetails
        Get
            Return _EvtDtl
        End Get
        Set(ByVal value As EventDetails)
            _EvtDtl = value
            LoadEventDetails()
        End Set
    End Property

    'Private _EventHeader As String
    'Public Property EventHeader() As String
    '    Get
    '        Return _EventHeader
    '    End Get
    '    Set(ByVal value As String)
    '        _EventHeader = value
    '    End Set
    'End Property

    Private _IsHomePage As Boolean
    'Public WriteOnly Property IsHomePage() As Boolean
    '    Set(ByVal value As Boolean)
    '        _IsHomePage = value
    '    End Set
    'End Property


    ''' <summary>
    ''' Load controls with event detail from the EventDetail class.  Any elements that need to be hidden
    ''' will be handled in the prerender event.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadEventDetails()
        'lblHeaderDescription.Text = _EventHeader
        'lnkEventImage.NavigateUrl = _EvtDtl.DetailURL
        imgEventImage.ImageUrl = _EvtDtl.ImageURL
        'If _IsHomePage _
        'AndAlso _EvtDtl.EventTypeId = EventTypes.Films Then
        '    imgEventImage.Width = New Unit(_MovieImageWidth)
        'End If
        lblEventName.Text = _EvtDtl.EventName
        'lnkEventDetail.Text = _EvtDtl.EventName
        'lnkEventDetail.Attributes.Add("name", String.Format("evt{0}", _EvtDtl.EventId))
        'lnkEventDetail.NavigateUrl = _EvtDtl.DetailURL
        lblStartDate.Text = IIf(_EvtDtl.StartDate = Date.MinValue, "", CType(_EvtDtl.StartDate, Date).ToShortDateString)
        lblEndDate.Text = IIf(_EvtDtl.EndDate = Date.MinValue, "", String.Format(" - {0}", CType(_EvtDtl.EndDate, Date).ToShortDateString))
        lblStartTime.Text = _EvtDtl.StartTime
        lblEndTime.Text = IIf(_EvtDtl.EndTime = String.Empty, String.Empty, String.Format(" - {0}", _EvtDtl.EndTime))
        lblLocation.Text = _EvtDtl.Location
        'If _IsHomePage AndAlso _EvtDtl.EventDescription.Length > _EventDescriptionMaxLength Then
        '    lblEventDescription.Text = _EvtDtl.EventDescription.Substring(0, _EvtDtl.EventDescription.IndexOf(" ", _EventDescriptionMaxLength)) & " ..."
        '    lnkEventDescriptionMore.Visible = True
        '    lnkEventDescriptionMore.NavigateUrl = _EvtDtl.DetailURL
        'Else
        lblEventDescription.Text = _EvtDtl.EventDescription
        lnkEventDescriptionMore.Visible = False
        'End If
        lnkWebsite.NavigateUrl = _EvtDtl.Website
        lnkVideo.NavigateUrl = _EvtDtl.VideoURL
        lnkAudio1.NavigateUrl = _EvtDtl.AudioURL1
        lnkAudio2.NavigateUrl = _EvtDtl.AudioURL2
        'If _IsHomePage Then
        '    tblEventDetail.Attributes.Add("class", "EventDetailHomePage")
        'Else
        '    tblEventDetail.Attributes.Add("class", "EventDetail")
        lnkVideo.ImageUrl = "~/images/movie.gif"
        lnkAudio1.ImageUrl = "~/images/audio.jpg"
        lnkAudio2.ImageUrl = "~/images/audio.jpg"
        'End If
    End Sub

    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        lnkAudio1.Visible = Not lnkAudio1.NavigateUrl.Equals("")
        lnkAudio2.Visible = Not lnkAudio2.NavigateUrl.Equals("")
        lnkVideo.Visible = Not lnkVideo.NavigateUrl.Equals("")
        lnkWebsite.Visible = Not lnkWebsite.NavigateUrl.Equals("")
        pnlEventImage.Visible = Not imgEventImage.ImageUrl.Equals("")
    End Sub

End Class
