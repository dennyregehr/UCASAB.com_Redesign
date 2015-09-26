Imports System.Data
Imports System.Web.Caching

Public Class EventUtil

    Private _eventCollectionName As String = "EventCollection"
    Private _eventCollectionNameCommitteeOnly As String = "EventCollectionCommitteeOnly"

    Public Sub New()
    End Sub
    Public Sub New(ByVal ClearCache As Boolean)
        DisposeEventCache()
    End Sub

    Private Sub DisposeEventCache()
        HttpContext.Current.Cache.Remove(_eventCollectionName)
        HttpContext.Current.Cache.Remove(_eventCollectionNameCommitteeOnly)
    End Sub

    Public Function GetEvents() As List(Of EventDetails)
        Dim evtDtls As New List(Of EventDetails)
        If HttpContext.Current.Cache(_eventCollectionName) Is Nothing Then
            Dim mySqlUtil As New SqlUtil()
            Dim dsEvents As DataSet = mySqlUtil.GetEventsByDate(Semester.StartDate, Semester.EndDate)
            If dsEvents.Tables(0) IsNot Nothing Then
                For Each dr As DataRow In dsEvents.Tables(0).Rows
                    evtDtls.Add(New EventDetails(dr))
                Next
            End If
            HttpContext.Current.Cache.Insert(_eventCollectionName, evtDtls, Nothing, Semester.EndDate, Cache.NoSlidingExpiration, CacheItemPriority.Default, Nothing)
        Else
            evtDtls = CType(HttpContext.Current.Cache(_eventCollectionName), List(Of EventDetails))
        End If
        Return evtDtls
    End Function

    Public Function GetEvents(ByVal CommitteeEventsOnly As Boolean) As List(Of EventDetails)
        Dim evtDtls As New List(Of EventDetails)
        If CommitteeEventsOnly Then
            If HttpContext.Current.Cache(_eventCollectionNameCommitteeOnly) Is Nothing Then
                evtDtls = (From e In Me.GetEvents _
                                 Where e.EventTypeId <= 5 _
                                 Select e).ToList()
                HttpContext.Current.Cache.Insert(_eventCollectionNameCommitteeOnly, evtDtls, Nothing, Semester.EndDate, Cache.NoSlidingExpiration, CacheItemPriority.Default, Nothing)
            Else
                evtDtls = CType(HttpContext.Current.Cache(_eventCollectionNameCommitteeOnly), List(Of EventDetails))
            End If
        Else
            evtDtls = GetEvents()
        End If
        Return evtDtls
    End Function

    Public Function GetEvents2(ByVal NumberOfEvents As Integer) As List(Of EventDetails)
        Dim evtDtls As New List(Of EventDetails)
        'If HttpContext.Current.Cache(_eventCollectionName) Is Nothing Then
        Dim mySqlUtil As New SqlUtil()
        Dim dsEvents As DataSet = mySqlUtil.GetEventsByDate(Semester.StartDate, Semester.EndDate)
        If dsEvents.Tables(0) IsNot Nothing Then
            For Each dr As DataRow In dsEvents.Tables(0).Rows
                If evtDtls.Count <= NumberOfEvents Then
                    evtDtls.Add(New EventDetails(dr))
                Else
                    Exit For
                End If
            Next
        End If
        '    HttpContext.Current.Cache.Insert(_eventCollectionName, evtDtls, Nothing, Semester.EndDate, Cache.NoSlidingExpiration, CacheItemPriority.Default, Nothing)
        'Else
        '    evtDtls = CType(HttpContext.Current.Cache(_eventCollectionName), List(Of EventDetails))
        'End If
        Return evtDtls
    End Function

    'Public Function GetEvents3(ByVal NumberOfEvents As Integer) As List(Of EventDetails)
    '    Dim eventsDB = New sab_boardDataContext
    '    Dim eventDtls = eventsDB.Calendars.Where(Function(e) e.StartDate >= Semester.StartDate) _
    '                    .Where(Function(e) e.StartDate <= Semester.EndDate) _
    '                    .Where(Function(e) e.StartDate >= Today) _
    '                    .Select(Function(ev) New EventDetails(ev)) _
    '                    .Take(NumberOfEvents) _
    '                    .ToList()
    '    Return eventDtls
    'End Function

    Public Function GetRemainingEvents(ByVal AddPreviousEventsAtEnd As Boolean, ByVal CommitteeEventsOnly As Boolean) As List(Of EventDetails)
        Dim ev = (From e In GetEvents(CommitteeEventsOnly) _
                  Where e.StartDate >= Today _
                  Select e).Concat(From e2 In GetEvents(CommitteeEventsOnly) _
                                   Where e2.StartDate >= Semester.StartDate _
                                   And e2.StartDate < Today _
                                   Select e2).ToList()
        Return ev
    End Function

    Public Function GetEventsByEventType(ByVal EventTypeId As Integer) As List(Of EventDetails)
        Dim ev = From e In Me.GetEvents _
                 Where e.EventTypeId = EventTypeId _
                 Select e
        Return CType(ev.ToList(), List(Of EventDetails))
    End Function

    Public Function GetEventsForRssFeed(ByVal NumberOfEvents As Int16) As List(Of EventDetails)
        Dim ev = (From e In Me.GetEvents _
                 Where e.StartDate > Today _
                 Select e).Take(NumberOfEvents)
        Return CType(ev.ToList(), List(Of EventDetails))
    End Function

    Public Function GetEvent(ByVal EventId As Integer) As EventDetails
        Try
            Return CType((From e In Me.GetEvents _
                     Where e.EventId = EventId
                     Select e).Single, EventDetails)
        Catch ex As Exception
            Return CType(GetRemainingEvents(True, False).Take(1), EventDetails)
        End Try
    End Function

End Class
