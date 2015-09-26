<%
If Session("Group") > 0 Then

Else
	If Request.Cookies("DB2ASP_ADMIN")("g") <> "" Then
		Session("Group") = Request.Cookies("DB2ASP_ADMIN")("g")
	Else
		'Response.Write "Off To Login Page (" & Session.TimeOut & ")<HR>"
		Response.Redirect "DB2ASP_login.asp"
	End If
End If
%>

