Imports System.Data

<Serializable()> _
Public Class EventDetails

    Private _EventId As Integer
    Public ReadOnly Property EventId() As Integer
        Get
            Return _EventId
        End Get
    End Property
    Private _EventName As String
    Public ReadOnly Property EventName() As String
        Get
            Return _EventName
        End Get
    End Property
    Private _EventTypeId As Integer
    Public ReadOnly Property EventTypeId() As Integer
        Get
            Return _EventTypeId
        End Get
    End Property
    Public ReadOnly Property EventTypeName() As String
        Get
            Return GetEventTypeName(_EventTypeId)
        End Get
    End Property

    Private _Location As String
    Public ReadOnly Property Location() As String
        Get
            Return _Location
        End Get
    End Property
    Private _StartDate As Nullable(Of Date) = Date.MinValue
    Public ReadOnly Property StartDate() As Nullable(Of Date)
        Get
            Return _StartDate
        End Get
    End Property
    Private _EndDate As Nullable(Of Date) = Date.MinValue
    Public ReadOnly Property EndDate() As Nullable(Of Date)
        Get
            Return _EndDate
        End Get
    End Property
    Private _StartTime As String
    Public Property StartTime() As String
        Get
            Return _StartTime
        End Get
        Set(ByVal value As String)
            _StartTime = value.Replace(":00", "")
        End Set
    End Property
    Private _EndTime As String
    Public Property EndTime() As String
        Get
            Return _EndTime
        End Get
        Set(ByVal value As String)
            If value <> "" Then
                _EndTime = value.Replace(":00", "").Insert(0, " - ")
            Else
                _EndTime = ""
            End If
        End Set
    End Property
    Private _EventDescription As String
    Public ReadOnly Property EventDescription() As String
        Get
            Return _EventDescription
        End Get
    End Property
    Private _Notes As String
    Public ReadOnly Property Notes() As String
        Get
            Return _Notes
        End Get
    End Property
    Private _ImageURL As String
    Public ReadOnly Property ImageURL() As String
        Get
            Return FixImageURL(_ImageURL)
        End Get
    End Property
    Private _Website As String
    Public ReadOnly Property Website() As String
        Get
            Return _Website
        End Get
    End Property
    Private _VideoURL As String
    Public ReadOnly Property VideoURL() As String
        Get
            Return _VideoURL
        End Get
    End Property
    Private _AudioURL1 As String
    Public ReadOnly Property AudioURL1() As String
        Get
            Return _AudioURL1
        End Get
    End Property
    Private _AudioURL2 As String
    Public ReadOnly Property AudioURL2() As String
        Get
            Return _AudioURL2
        End Get
    End Property
    Private _DetailURL As String
    Public ReadOnly Property DetailURL() As String
        Get
            Return _DetailURL
        End Get
    End Property

    Public Sub New(ByVal EventId As Integer, ByVal EventName As String, ByVal EventTypeId As Integer, ByVal Location As String, ByVal StartDate As Nullable(Of Date), _
                   ByVal EndDate As Nullable(Of Date), ByVal StartTime As String, ByVal EndTime As String, ByVal EventDescription As String, ByVal Notes As String, _
                   ByVal ImageUrl As String, ByVal Website As String, ByVal VideoUrl As String, ByVal AudioUrl1 As String, ByVal AudioUrl2 As String)
        With Me
            ._EventId = EventId
            ._EventName = EventName
            ._EventTypeId = EventTypeId
            ._Location = Location
            ._StartDate = StartDate
            ._EndDate = EndDate
            ._StartTime = StartTime
            ._EndTime = EndTime
            ._EventDescription = EventDescription
            ._Notes = Notes
            ._ImageURL = ImageUrl
            ._Website = Website
            ._VideoURL = VideoUrl
            ._AudioURL1 = AudioUrl1
            ._AudioURL2 = AudioUrl2
            '._DetailURL = GetLinkURL(Me._EventTypeId, Me._EventId)
            ._DetailURL = GetLinkURL(_EventId)
        End With
    End Sub
    'Public Sub New(ByVal calendarEvent As Calendar)
    '    With calendarEvent
    '        _EventId = .EventID
    '        _EventName = .EventName
    '        _EventTypeId = .EventTypeID
    '        _Location = .Location
    '        _StartDate = .StartDate
    '        _EndDate = .EndDate
    '        _StartTime = .StartTime.ToString
    '        _EndTime = .EndTime.ToString
    '        _EventDescription = .EventDescription
    '        _Notes = .Notes
    '        _ImageURL = .imageURL
    '        _Website = .website
    '        _VideoURL = .videoURL
    '        _AudioURL1 = .audioURL1
    '        _AudioURL2 = .audioURL2
    '        _DetailURL = GetLinkURL(Me._EventTypeId, Me._EventId)
    '    End With
    'End Sub

    Public Sub New(ByVal EventDataRow As DataRow)
        With Me
            ._EventId = SafeData(EventDataRow("EventId"))
            ._EventName = SafeData(EventDataRow("EventName"))
            ._EventTypeId = SafeData(EventDataRow("EventTypeId"))
            ._Location = SafeData(EventDataRow("Location"))
            Dim stDt As String = SafeData(EventDataRow("StartDate"))
            If stDt IsNot String.Empty Then
                ._StartDate = Convert.ToDateTime(stDt)
            End If
            Dim enDt As String = SafeData(EventDataRow("EndDate"))
            If enDt IsNot String.Empty Then
                ._EndDate = Convert.ToDateTime(enDt)
            End If
            .StartTime = SafeData(EventDataRow("StartTime"))
            .EndTime = SafeData(EventDataRow("EndTime"))
            ._EventDescription = SafeData(EventDataRow("EventDescription"))
            ._Notes = SafeData(EventDataRow("Notes"))
            ._ImageURL = SafeData(EventDataRow("ImageURL"))
            If ._ImageURL = "" Then
                ._ImageURL = "/images/sab_event_generic.png"
            End If
            ._Website = SafeData(EventDataRow("Website"))
            ._VideoURL = SafeData(EventDataRow("videoURL"))
            ._AudioURL1 = SafeData(EventDataRow("audioURL1"))
            ._AudioURL2 = SafeData(EventDataRow("audioURL2"))
            '._DetailURL = GetLinkURL(Me._EventTypeId, Me._EventId)
            ._DetailURL = GetLinkURL(_EventId)
        End With
    End Sub

    Private Function SafeData(ByVal DataItem As Object) As String
        Return IIf(IsDBNull(DataItem), String.Empty, DataItem)
    End Function

    Private Function FixImageURL(ByVal URL As String) As String
        Dim myNewURL As String = URL
        If URL.IndexOf("images") = 0 Then
            'accomodate older coded urls
            myNewURL = String.Format("~/{0}", URL)
        End If
        Return myNewURL
    End Function

    Private Function GetLinkURL(ByVal EventId As String) As String
        Return String.Format("/Event.aspx?id={0}&desc={1}", EventId, _EventName.Replace(" ", "-"))
    End Function
    Private Function GetLinkURL(ByVal EventTypeId As String, ByVal EventId As String) As String
        Dim retVal As String
        Select Case EventTypeId
            Case EventTypes.Music
                retVal = "~/committee.aspx?cmt=music#evt{0}"
            Case EventTypes.Comedy
                retVal = "~/committee.aspx?cmt=comedy#evt{0}"
            Case EventTypes.PopCulture
                retVal = "~/committee.aspx?cmt=pop%20culture#evt{0}"
            Case EventTypes.Films
                'retVal = "~/weekend.aspx#evt{0}"
                retVal = "~/committee.aspx?cmt=films#evt{0}"
            Case EventTypes.Novelty
                retVal = "~/committee.aspx?cmt=novelty#evt{0}"
            Case Else
                retVal = "~/calendar.aspx"
        End Select
        retVal = String.Format(retVal, EventId)
        Return retVal
    End Function

    Private Function GetEventTypeName(ByVal EventTypeId As Integer) As String
        Select Case EventTypeId
            Case EventTypes.Comedy
                Return "comedy"
            Case EventTypes.Music
                Return "music"
            Case EventTypes.Films
                Return "movies"
            Case EventTypes.Novelty
                Return "novelty"
            Case Else   'EventTypes.PopCulture
                Return "pop%20culture"
        End Select
    End Function
End Class
