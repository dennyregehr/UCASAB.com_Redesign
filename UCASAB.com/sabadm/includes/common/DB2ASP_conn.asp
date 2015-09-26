<%
' Establish Connection To The Database
Set oConn = Server.CreateObject("ADODB.Connection")
'oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WebDev\SAB\_private\sab_calendar.mdb;Persist Security Info=False;"
'oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/_private/sab_calendar.mdb") & ";Persist Security Info=False;"
oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/App_Data/sab_calendar.mdb") & ";Persist Security Info=False;"
oConn.ConnectionTimeout = 7  ' Time in seconds, adjust if needed
oConn.CommandTimeout = 25  ' Time in seconds, adjust if needed
oConn.Open

If Err Then
	sErr = "Error Establishing Database Connection."
	If Err.Number = -2147467259 Then
		sErr = sErr & " (Can't Connect To Database)<BR><BR>"
		sErr = sErr & "Details:<BR>"
		sErr = sErr & Err.Description & "<BR><BR>"
		sErr = sErr & "If your database is in the right directory and it still can't connect, then the table may be locked by another program or the web user doesnt have permission to use the database. Make sure the Windows [IUSR_...] user accounts have read and write permissions for the database directory and file. Then refresh this page.<BR><BR>"
	ElseIf Err.Number = -2147217843 Then
		sErr = sErr & "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This table may be locked by another program or the web user doesnt have permission to use the database. Make sure the IUSR_... users have read and write permissions for the database directory and file. Then refresh this page.<BR><BR>"
	Else
		sErr = sErr & "<BR> "
	End If
	Err.Clear
End If
%>

