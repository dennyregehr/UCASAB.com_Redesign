<%
' Add A New Record To The Database
Function DetailAdd()
	Dim bReturn
	Dim sDateFormat1, sDateFormat2
	Dim varEventTypeID
	Dim varStartDate
	Dim varEndDate
	Dim varStartTime
	Dim varEndTime

	' Check For empty/null values submited from web page form.
	If Request("EventTypeID") <> "" Then
		varEventTypeID = Replace(Request("EventTypeID"), "'", "''")
	Else
		varEventTypeID = "Null"
	End If

	If Request("StartDate") <> "" Then
		varStartDate ="#" & Replace(Request("StartDate"), "'", "''") & "#"
	Else
		varStartDate = "Null"
	End If

	If Request("EndDate") <> "" Then
		varEndDate ="#" & Replace(Request("EndDate"), "'", "''") & "#"
	Else
		varEndDate = "Null"
	End If

	If Request("StartTime") <> "" Then
		varStartTime ="#" & Replace(Request("StartTime"), "'", "''") & "#"
	Else
		varStartTime = "Null"
	End If

	If Request("EndTime") <> "" Then
		varEndTime ="#" & Replace(Request("EndTime"), "'", "''") & "#"
	Else
		varEndTime = "Null"
	End If


	sSQL = _
		"INSERT INTO [Calendar] (" & _
			"[EventName], " & _
			"[EventTypeID], " & _
			"[Location], " & _
			"[StartDate], " & _
			"[EndDate], " & _
			"[StartTime], " & _
			"[EndTime], " & _
			"[EventDescription], " & _
			"[Notes], " & _
			"[imageURL], " & _
			"[website], " & _
			"[videoURL], " & _
			"[audioURL1], " & _
			"[audioURL2] " & _
		") VALUES (" & _
			"'" & Replace(Request("EventName"),"'","''") & "', " & _
			"" & varEventTypeID & ", " & _
			"'" & Replace(Request("Location"),"'","''") & "', " & _
			"" & varStartDate & ", " & _
			"" & varEndDate & ", " & _
			"" & varStartTime & ", " & _
			"" & varEndTime & ", " & _
			"'" & Replace(Request("EventDescription"),"'","''") & "', " & _
			"'" & Replace(Request("Notes"),"'","''") & "', " & _
			"'" & Replace(Request("imageURL"),"'","''") & "', " & _
			"'" & Replace(Request("website"),"'","''") & "', " & _
			"'" & Replace(Request("videoURL"),"'","''") & "', " & _
			"'" & Replace(Request("audioURL1"),"'","''") & "', " & _
			"'" & Replace(Request("audioURL2"),"'","''") & "')"
	'Response.Write sSQL & "<HR width=300>": Response.Flush

	'oConn.BeginTrans
	oConn.Execute(sSQL)

	If Err Then
		'oConn.RollbackTrans
		bReturn = False
		Err.Clear
	Else
		Application("EventTypeeventTypeIDLookup") = ""
		Application("ExecPositionexecPositionIDLookup") = ""
		'oConn.CommitTrans
		bReturn = True
	End If

	DetailAdd = bReturn

End Function

' Update Database Record
Function DetailUpdate()
	Dim bReturn
	Dim sDateFormat1, sDateFormat2
	Dim varEventTypeID
	Dim varStartDate
	Dim varEndDate
	Dim varStartTime
	Dim varEndTime

	' Check For empty/null values submited from web page form.
	If Request("EventTypeID") <> "" Then
		varEventTypeID = Replace(Request("EventTypeID"), "'", "''")
	Else
		varEventTypeID = "Null"
	End If

	If Request("StartDate") <> "" Then
		varStartDate = "#" & Replace(Request("StartDate"), "'", "''") & "#"
	Else
		varStartDate = "Null"
	End If

	If Request("EndDate") <> "" Then
		varEndDate = "#" & Replace(Request("EndDate"), "'", "''") & "#"
	Else
		varEndDate = "Null"
	End If

	If Request("StartTime") <> "" Then
		varStartTime = "#" & Replace(Request("StartTime"), "'", "''") & "#"
	Else
		varStartTime = "Null"
	End If

	If Request("EndTime") <> "" Then
		varEndTime = "#" & Replace(Request("EndTime"), "'", "''") & "#"
	Else
		varEndTime = "Null"
	End If


	sSQL = _
	"UPDATE [Calendar] " & _
	"SET " & _
	"[EventName]='" & Replace(Request("EventName"),"'","''") & "', " & _
	"[EventTypeID]=" & varEventTypeID & ", " & _
	"[Location]='" & Replace(Request("Location"),"'","''") & "', " & _
	"[StartDate]=" & varStartDate & ", " & _
	"[EndDate]=" & varEndDate & ", " & _
	"[StartTime]=" & varStartTime & ", " & _
	"[EndTime]=" & varEndTime & ", " & _
	"[EventDescription]='" & Replace(Request("EventDescription"),"'","''") & "', " & _
	"[Notes]='" & Replace(Request("Notes"),"'","''") & "', " & _
	"[imageURL]='" & Replace(Request("imageURL"),"'","''") & "', " & _
	"[website]='" & Replace(Request("website"),"'","''") & "', " & _
	"[videoURL]='" & Replace(Request("videoURL"),"'","''") & "', " & _
	"[audioURL1]='" & Replace(Request("audioURL1"),"'","''") & "', " & _
	"[audioURL2]='" & Replace(Request("audioURL2"),"'","''") & "' "
	sSQL = sSQL & "WHERE [EventID]=" & Request("PK_iEventID") & ""
	'Response.Write sSQL & "<HR width=300>": Response.Flush

	'oConn.BeginTrans
	oConn.Execute(sSQL)

	If Err Then
		'oConn.RollbackTrans
		bReturn = False
	Else
		Application("EventTypeeventTypeIDLookup") = ""
		Application("ExecPositionexecPositionIDLookup") = ""
		'oConn.CommitTrans
		bReturn = True
	End If

	DetailUpdate = bReturn
End Function

' Delete A Record From The Database
Function DetailDelete()
	Dim bReturn
	sSQL = _
		"DELETE FROM [Calendar] " & _
		"WHERE [EventID]=" & Request("PK_iEventID") & ""
	'Response.Write sSQL & "<HR width=300>": Response.Flush

	'oConn.BeginTrans
	oConn.Execute(sSQL)

	If Err Then
		'oConn.RollbackTrans
		bReturn = False
	Else
		'oConn.CommitTrans
		bReturn = True
	End If

	Set oConn = Nothing
	DetailDelete = bReturn
End Function

' Delete Record/s From The Database
Function ResultsDelete()
	Dim bReturn, Item

	' Transactional database processing - each delete is "buffered". 
	' If all requested deletes are good then they are executed, otherwise they are canceled or rolled back.
	' Using a transaction is optional, but its safer with one
	oConn.BeginTrans

	iCounter = 0 
	On Error Resume Next
	For Each Item In Request.Form
		If Len(Request(Item)) > 0 Then
			If Left(Item, 4) = "chkd" Then
				iCounter = iCounter + 1
				sSQL = "DELETE FROM [Calendar] " & _
					"WHERE " & Replace(Request(Item), "__", " ")
				'Response.Write sSQL & "<HR width=300>": Response.Flush
	
				oConn.Execute(sSQL)
			End If
		End If
	Next

	If Err Then
		sErr = Err.Description & " (" & Err.Number & ")"
		oConn.RollbackTrans
		bReturn = -1
	Else
		Application("EventTypeeventTypeIDLookup") = ""
		Application("ExecPositionexecPositionIDLookup") = ""
		oConn.CommitTrans
		bReturn = iCounter
	End If

	On Error Goto 0
	Set oConn = Nothing
	ResultsDelete = bReturn
End Function

' Add A New Record To The Database From The Results Page
Function ResultsAdd()
	Dim bReturn
	Dim sDateFormat1, sDateFormat2
	Dim varEventTypeID
	Dim varStartDate
	Dim varEndDate
	Dim varStartTime

	' Check For empty/null values submited from web page form.
	If Request("AddEventTypeID") <> "" Then
		varEventTypeID = Replace(Request("AddEventTypeID"), "'", "''")
	Else
		varEventTypeID = "Null"
	End If

	If Request("AddStartDate") <> "" Then
		varStartDate ="#" &Replace(Request("AddStartDate"), "'", "''") & "#"
	Else
		varStartDate = "Null"
	End If

	If Request("AddEndDate") <> "" Then
		varEndDate ="#" &Replace(Request("AddEndDate"), "'", "''") & "#"
	Else
		varEndDate = "Null"
	End If

	If Request("AddStartTime") <> "" Then
		varStartTime ="#" &Replace(Request("AddStartTime"), "'", "''") & "#"
	Else
		varStartTime = "Null"
	End If


	sSQL = _
		"INSERT INTO [Calendar] (" & _
			"[EventName], " & _
			"[EventTypeID], " & _
			"[Location], " & _
			"[StartDate], " & _
			"[EndDate], " & _
			"[StartTime] " & _
		") VALUES (" & _
			"'" & Replace(Request("AddEventName"),"'","''") & "', " & _
			"" & varEventTypeID & ", " & _
			"'" & Replace(Request("AddLocation"),"'","''") & "', " & _
			"" & varStartDate & ", " & _
			"" & varEndDate & ", " & _
			"" & varStartTime & ")"
	'Response.Write sSQL & "<HR width=300>": Response.Flush

	'oConn.BeginTrans
	oConn.Execute(sSQL)

	If Err Then
	If Err.Number = -2147467259 Then
		' Cant add duplicate key value in index
		sErr = "Duplicate values in the index, primary key, or relationship"
	End If

		'oConn.RollbackTrans
		bReturn = False
		Err.Clear
	Else
		Application("EventTypeeventTypeIDLookup") = ""
		Application("ExecPositionexecPositionIDLookup") = ""
		'oConn.CommitTrans
		bReturn = True
	End If

	Set oConn = Nothing
	ResultsAdd = bReturn

End Function

%>
