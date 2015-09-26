
Partial Class Calendar
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim eventUtil As New EventUtil()
        'Dim events = eventUtil.GetEvents _
        '             .Where(Function(ev) ev.EventTypeId <= 5)
        Dim events = eventUtil.GetEvents

        Dim sb As New StringBuilder(String.Empty)
        Dim currentMonthNumber As Integer
        For Each evt As EventDetails In events
            If IsNewMonth(currentMonthNumber, evt.StartDate) Then
                currentMonthNumber = Month(evt.StartDate)
                sb.Length = 0
                sb.AppendFormat("<h2 style='clear:both;'>{0}</h2>", MonthName(currentMonthNumber))
                ph1.Controls.Add(New LiteralControl(sb.ToString()))
            End If
            With sb
                .Length = 0
                .AppendFormat("<div class=""CalendarItem"" onclick=""window.location='Event.aspx?id={0}'"">", evt.EventId)
                .AppendFormat("<img src='{0}' /><br />", evt.ImageURL)
                .AppendFormat("{0}<br/>", evt.EventName)
                .AppendFormat("{0}, {1} @ {2}<br/>", CDate(evt.StartDate).DayOfWeek, CDate(evt.StartDate).ToShortDateString(), evt.StartTime)
                '.AppendFormat("{0}", evt.StartTime)
                .AppendFormat("{0}", evt.Location)
                .Append("</div>")
                ph1.Controls.Add(New LiteralControl(.ToString()))
            End With
        Next


    End Sub

    Private Function IsNewMonth(ByVal currentMonth As Integer, ByVal newDate As Date?) As Boolean
        Return currentMonth <> Month(newDate)
    End Function

End Class
