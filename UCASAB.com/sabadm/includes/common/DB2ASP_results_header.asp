<%
If Err Then
	If Err = "" Then
		sErr = "Error On Web Page.<BR>"
		sErr = sErr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SPAN class=DB2ASPerror>[" & Err.Number & "] - " & Err.Description & "</SPAN><BR><BR>"
		Err.Clear
	End If

End If

' Initialize main recordset and setup paging parameters
' Response.Write sSQL & "<HR>": Response.Flush
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.CursorLocation = adUseClient
oRS.ActiveConnection = oConn
oRS.Open sSQL, oConn, adOpenForwardOnly, adLockOptimistic, &H1

If Err Then
	If Err = "" Then
		sErr = "Error Running Query.<BR>"
		sErr = sErr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SPAN class=DB2ASPerror>[" & Err.Number & "] - " & Err.Description & "</SPAN><BR><BR>"
		Err.Clear
	End If

End If

oRS.PageSize = iPageSize
oRS.CacheSize = iPageSize
iPageCount = oRS.PageCount
iRecordCount = oRS.RecordCount
iTotalRecords = oRS.RecordCount

If iPageCount > CInt(iMaxRecords) / CInt(iPageSize) Then
	iPageCount = CInt(iMaxRecords) / CInt(iPageSize)
End If

If Request("view") = "newsort" Then
	iPage = 1
Else
	iPage = GetLastPage(sObjectName)
End If

If CInt(iPage) > CInt(iPageCount) Then iPage = iPageCount
If CInt(iPage) <= 0 Then iPage = 1

' Determine the record range the page is showing
If iRecordCount > 0 Then
	oRS.AbsolutePage = iPage
	iStart = iPageSize * (iPage - 1) + 1
	If iPage = iPageCount Then
		iFinish = iRecordCount
	Else
		iFinish = iStart + iPageSize - 1
	End If
End If

If Err Then
	If Err = "" Then
		sErr = "Error On Web Page.<BR>"
		sErr = sErr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SPAN class=DB2ASPerror>[" & Err.Number & "] - " & Err.Description & "</SPAN><BR><BR>"
		Err.Clear
	End If
End If
%>

