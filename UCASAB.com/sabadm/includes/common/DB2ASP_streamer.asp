<%
Dim sPK
Dim oleHeaderSize, binImage
Dim lngImageSize, binTemp, pvoImageType
Response.Expires = 0
Response.Buffer = True
Response.Clear
pvoImageType = Request("type")

' Establish Connection To The Database
Set oConnPic = Server.CreateObject("ADODB.Connection")
oConnPic.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WebDev\SAB\_private\sab_calendar.mdb;Persist Security Info=False;"
oConnPic.ConnectionTimeout = 7  ' Time in seconds, adjust if needed
oConnPic.CommandTimeout = 10  ' Time in seconds, adjust if needed
oConnPic.Open

For Each Item In Request.QueryString
	If Left(Item, 3) = "PK_" Then
		If Mid(Item, 4, 1) = "s" Then
			sPK = sPK & "[" & Replace(CStr(Mid(Item, 5)), "__", " ") & "]='" & Request(Item) & "' AND "
		ElseIf Mid(Item, 4, 1) = "d" Then
			sPK = sPK & "[" & Replace(CStr(Mid(Item, 5)), "__", " ") & "]=#" & Request(Item) & "# AND "
		Else
			sPK = sPK & "[" & Replace(CStr(Mid(Item, 5)), "__", " ") & "]=" & Request(Item) & " AND "
		End If
	End If
Next

If Right(sPK, 4) = "AND " Then sPK = Left(sPK, Len(sPK) - 5)

' table and where clause needs to be dynamic
sSQL = "SELECT [" & Request("picfield") & "] FROM [" & Request("pictable") & "] " & _
	"WHERE " & sPK
'Response.Write sSQL & "<HR WIDTH=300>": Response.Flush
Set oRSPic = oConnPic.Execute(sSQL)

Select Case LCase(pvoImageType)
Case "gif", "jpg", "bmp", "jpeg"
	pvoImageType = "image/" & pvoImageType
	binImage = oRSPic(0).Value
Case "ole"
	oleHeaderSize = 78
	pvoImageType = "image/bmp"
	lngImageSize = oRSPic(0).ActualSize
	binTemp = oRSPic(0).GetChunk(oleHeaderSize)
	binImage = oRSPic(0).GetChunk(lngImageSize - oleHeaderSize)
End Select

Response.ContentType = pvoImageType
Response.BinaryWrite binImage
oRSPic.Close: Set oRSPic = Nothing: oConnPic.Close: Set oConnPic = Nothing
Response.End
%>

