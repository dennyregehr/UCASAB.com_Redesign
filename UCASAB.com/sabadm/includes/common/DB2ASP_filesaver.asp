<%
' ********************************
' DB2ASP File Saver
' Copyright 2002 - 303 Media, LLC
' All rights reserved
' ********************************

Option Explicit
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

Dim oConnPic
Dim oRS
Dim sSQL
Dim oDB2ASPFileUpload
Dim varSplit
Dim sFileName
Dim sFilePath
Dim sDestinationPath
Dim strFile
Dim sSQLUpdate
Dim sSQLDelete
Dim yByteStream
Dim iPictureID
Dim sPicField
Dim sQS
Dim sPK
Dim Item
Dim sAction
Dim sPicTableName
Dim sPicFieldName
Dim sPageName

sAction = Request.QueryString("a")
sPicTableName = Request.QueryString("pictable")
sPicFieldName = Request.QueryString("picfield")
sPageName = Request.QueryString("pagename")
sFilePath = "images/DB2ASP/upload/"
sDestinationPath = Server.MapPath("..\..\images\DB2ASP\upload\")

sQS = Request.QueryString

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

'Response.Write Request.QueryString & "<HR>"
'Response.Write Request("method") & "<HR>"

sSQLUpdate = "SELECT [" & sPicFieldName & "] FROM [" & sPicTableName & "] WHERE " & sPK
sSQLDelete = "UPDATE [" & sPicTableName & "] SET [" & sPicFieldName & "] = Null WHERE " & sPK
'Response.Write sSQLUpdate & "<HR>"
'Response.Write sSQLDelete & "<HR>": Response.Flush
Select Case Request("method")
Case "db2asp"

sFileName = Request("File1")

'Response.Write sFileName & "<HR>": Response.Flush

Set oDB2ASPFileUpload = Server.CreateObject("DB2ASPFileUpload.CUpload")
oDB2ASPFileUpload.Path = sDestinationPath
oDB2ASPFileUpload.FileName = sFileName

If oDB2ASPFileUpload.Save Then
	sFileName = oDB2ASPFileUpload.Path & oDB2ASPFileUpload.FileName
End If

If Len(sFileName) < 1 Then
	If DeleteImage() Then
		Response.Redirect "../../" & sPageName & "?" & Request.QueryString()
	Else
		Response.Redirect "../../" & sPageName & "?msg=bad&" & Request.QueryString()
	End If
Else
	If UpdateImage() Then
		Response.Redirect "../../" & sPageName & "?" & Request.QueryString()
	Else
		Response.Redirect "../../" & sPageName & "?msg=bad&" & Request.QueryString()
	End If
End If

Response.End

Case "asp"
	Dim iFileUploaded
	Dim iBytes
	Dim vData
	Dim iLenData
	Dim sBoundary
	Dim iBoundaryPos
	Dim iRawFileStartPos
	Dim iRawFileEndPos
	Dim sRawFileData
	Dim iFileNameStartPos
	Dim iFileNameEndPos
	Dim sRawFileName
	'Dim sFileName
	Dim iCatch
	Dim iPos1
	Dim iPos2
	Dim iContentTypePos
	Dim oFSO
	Dim oFile
	Dim iDataFileStartPos
	Dim iDataFileEndPos
	Dim sFileData
	Dim iFileSize
	Dim iFilesUploaded
	Dim sData

Server.ScriptTimeout = 600 ' 10 Minutes
iFilesUploaded = 0
'Response.Write "<B>Saving File/s...</B><BR>"': Response.Flush
'Response.Write "<B>Saving File/s...</B><BR>" & Server.MapPath(".\images\upload"): Response.Flush

iBytes = Request.TotalBytes
vData = Request.BinaryRead(iBytes)
iLenData = LenB(vData)
'convery the binary data to a string
If iLenData > 0 Then
	Set oRS = CreateObject("ADODB.Recordset")
	oRS.Fields.Append "myBinary", 201, iLenData
	oRS.Open
	oRS.AddNew
	oRS("myBinary").AppendChunk vData
	oRS.Update
	sData = oRS("myBinary")
End If

sBoundary = Request.ServerVariables("HTTP_CONTENT_TYPE")
iBoundaryPos = InStr(1, sBoundary, "boundary=") + 8
sBoundary = "--" & Right(sBoundary, Len(sBoundary) - iBoundaryPos)
iRawFileStartPos = InStr(1, sData, sBoundary)
iRawFileEndPos = InStr(iRawFileStartPos + 1, sData, sBoundary) - 1
iCatch = 0

Do While iRawFileEndPos > 0
	sRawFileData = Mid(sData, iRawFileStartPos, (iRawFileEndPos - iRawFileStartPos) + 1)
	iFileNameStartPos = InStr(1, sRawFileData, "filename=") + 10
	iFileNameEndPos = InStr(iFileNameStartPos, sRawFileData, Chr(34))
	
	If iFileNameStartPos <> iFileNameEndPos And iFileNameStartPos - 10 <> 0 Then
		sRawFileName = Mid(sRawFileData, iFileNameStartPos, iFileNameEndPos - iFileNameStartPos)

		iPos1 = InStr(1, sRawFileName, "\")
		Do While iPos1 > 0
			iPos2 = iPos1
			iPos1 = InStr(iPos2 + 1, sRawFileName, "\")
		Loop

		sFileName = Right(sRawFileName, Len(sRawFileName) - iPos2)
		'Response.Write sDestinationPath & sFileName & "<BR>": Response.Flush
		iContentTypePos = InStr(1, sRawFileData, "Content-Type:")

		If iContentTypePos > 0 Then
			iDataFileStartPos = InStr(iContentTypePos, sRawFileData, Chr(13) & Chr(10)) + 4
		Else
			iDataFileStartPos = iFileNameEndPos
		End If

		iDataFileEndPos = Len(sRawFileData)

		iFileSize = (iDataFileEndPos - iDataFileStartPos) - 1
		sFileData = Mid(sRawFileData, iDataFileStartPos, iFileSize)

		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oFile = oFSO.OpenTextFile(sDestinationPath & sFileName, 2, True)
		oFile.Write sFileData
		Set oFile = Nothing: Set oFSO = Nothing

		iFilesUploaded = iFilesUploaded + 1

	End If

	iRawFileStartPos = iRawFileEndPos
	iRawFileEndPos = InStr(iRawFileStartPos + 9, sData, sBoundary) - 1
	iCatch = iCatch + 1

	If iCatch = 20 Then
		Response.Write "looped 100 times terminating script!"
		Response.End
	End If
Loop

If Err Then
	Response.Redirect sPageName & "?msg=error1&" & sQS
Else
	sFileName = sDestinationPath & sFileName

	If Len(sFileName) < 1 Then
		If DeleteImage() Then
			Response.Redirect "../../" & sPageName & "?" & Request.QueryString()
		Else
			Response.Redirect "../../" & sPageName & "?msg=bad&" & Request.QueryString()
		End If
	Else
		If UpdateImage() Then
			Response.Redirect "../../" & sPageName & "?" & Request.QueryString()
		Else
			Response.Redirect "../../" & sPageName & "?msg=bad&" & Request.QueryString()
		End If
	End If

End If

End Select

Response.End

Function DeleteImage()
	On Error Resume Next
	' Establish Connection To The Database
	Set oConnPic = Server.CreateObject("ADODB.Connection")
	oConnPic.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WebDev\SAB\_private\sab_calendar.mdb;Persist Security Info=False;"
	oConnPic.Open
	oConnPic.Execute (sSQLDelete)
	If Err Then
		DeleteImage = False
	Else
		DeleteImage = True
	End If
End Function

Function UpdateImage()
	Dim varSplit
	Dim yByteStream

	On Error Resume Next

	' Establish Connection To The Database
	Set oConnPic = Server.CreateObject("ADODB.Connection")
	oConnPic.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WebDev\SAB\_private\sab_calendar.mdb;Persist Security Info=False;"
	oConnPic.Open

	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.Open sSQLUpdate, oConnPic, adOpenKeyset, adLockOptimistic
	If sAction = "1" Then
		Set yByteStream = Server.CreateObject("ADODB.Stream")
		yByteStream.Type = 1 '1=Binary Data, 2=Text Data
		yByteStream.Open
		yByteStream.LoadFromFile sFileName
		oRS.Fields(sPicFieldName).Value = yByteStream.Read
	ElseIf sAction = "2" Then
		oRS.Fields(sPicFieldName).Value = sFileName
	End If
	oRS.Update: oRS.Close: Set oRS = Nothing: Set yByteStream = Nothing: Set oConnPic = Nothing
	If Err Then
		UpdateImage = False
	Else
		UpdateImage = True
	End If
	Set oDB2ASPFileUpload = Nothing
End Function

%>
