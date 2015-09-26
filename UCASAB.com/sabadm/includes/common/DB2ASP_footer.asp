<%
' JScript - Stop Timer and Display Message
On Error Resume Next
iTimeEnd = GetTimeStamp() ' Stop the stopwatch
Response.Write "<SPAN class=DB2ASPvalue>Page Took " & FormatNumber(iTimeEnd - iTimeStart, 0) & " Milliseconds.</SPAN>"
%>
</TD></TR></TABLE>
</BODY>
</HTML>
<%

On Error Resume Next
oRS.Close: Set oRS = Nothing: oConn.Close: Set oConn = Nothing

Response.End
%>

