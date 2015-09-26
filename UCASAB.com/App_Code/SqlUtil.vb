Imports System.Data.OleDb
Imports System.Data

Public Class SqlUtil

    Dim _con As OleDbConnection
    Dim _cmd As OleDbCommand

    Public Sub New()
    End Sub

    Private Sub GetConnection()
        Try
            If _con Is Nothing Then
                Dim myConStr As String = ConfigurationManager.AppSettings("CalendarConnectionString")
                Dim myDbPath As String = ConfigurationManager.AppSettings("CalendarDatasourcePath")
                _con = New OleDbConnection(String.Format(myConStr, HttpContext.Current.Server.MapPath(myDbPath)))
            End If

        Catch ex As OleDbException
            Throw New Exception("Trouble creating a database connection.", ex)
        End Try
        Try
            If _con.State <> Data.ConnectionState.Open Then
                _con.Open()
            End If

        Catch ex As OleDbException
            Throw New Exception("Trouble opening the database connection.", ex)
        End Try
    End Sub

    Public Function GetHomePageEventDetails() As DataSet

        Dim da As New OleDbDataAdapter
        Dim sql As String = ""
        Dim ds As New DataSet
        
        Try
            GetConnection()

            sql = "Select Top 2 * " & _
                " From calendar " & _
                " Where eventtypeid in (1,2,3,5) " & _
                " And startdate >= #" & Today.ToShortDateString & "# " & _
                " Order By startdate, starttime"
            da.SelectCommand = New OleDbCommand(sql, _con)
            da.Fill(ds, "Table1")

            sql = "Select Top 2 * " & _
                " From calendar " & _
                " Where eventtypeid in (4) " & _
                " And startdate >= #" & Today.ToShortDateString & "# " & _
                " Order By startdate, starttime"
            da.SelectCommand.CommandText = sql
            da.Fill(ds, "Table2")

            ds.Tables(0).Merge(ds.Tables(1))

            Return ds

        Catch ex As OleDbException
            Throw New Exception("Trouble getting event data from the database.", ex)
        End Try

        Return Nothing
    End Function

    Public Function GetEventsForScroller(ByVal NumberOfDaysAhead As Integer) As DataSet
        
        Dim da As New OleDbDataAdapter
        Dim sql As String = ""
        Dim ds As New DataSet

        Try
            GetConnection()

            sql = "Select eventid, startdate, eventname, starttime, eventtypeid " & _
                "From calendar " & _
                "Where startdate Between #" & Today.ToShortDateString & "# And #" & DateAdd("d", NumberOfDaysAhead, Today) & "#" & _
                " Order By startdate, starttime"

            da.SelectCommand = New OleDbCommand(sql, _con)
            da.Fill(ds, "Table1")

            Return ds
        Catch ex As OleDbException
            Throw New Exception("Trouble getting events for scroller.", ex)
        End Try

        Return Nothing
    End Function

    Public Function GetEventsByDate(ByVal StartDate As Date, ByVal EndDate As Date) As DataSet

        Dim ds As New DataSet

        Try
            GetConnection()
            Using _con
                Dim sql As String = ""
                sql = "Select * " & _
                    " From calendar " & _
                    " Where startdate >= #" & StartDate & "# " & _
                    " AND startdate <= #" & EndDate & "# " & _
                    " Order By startdate, starttime"
                Using cmd As New OleDbCommand(sql, _con)
                    Using da As New OleDbDataAdapter(cmd)
                        da.Fill(ds, "Table1")
                    End Using
                End Using
            End Using

            Return ds

        Catch ex As OleDbException
            Throw New Exception("Trouble getting event data from the database.", ex)
        End Try

        Return Nothing
    End Function

    Public Function GetExecutiveMembers() As DataSet

        Dim ds As New DataSet

        Try
            GetConnection()
            Using _con
                Dim sql As String = "spExec_GetPositionList2"
                Using cmd As New OleDbCommand(sql, _con)
                    cmd.CommandType = CommandType.StoredProcedure
                    Using da As New OleDbDataAdapter(cmd)
                        da.Fill(ds, "Table1")
                    End Using
                End Using
            End Using

            Return ds

        Catch ex As OleDbException
            Throw New Exception("Trouble getting executive members data from the database.", ex)
        End Try

        Return Nothing

    End Function

End Class
