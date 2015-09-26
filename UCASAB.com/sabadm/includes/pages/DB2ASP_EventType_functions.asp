<%
' Add A New Record To The Database
Function DetailAdd()
	Dim bReturn
	Dim sDateFormat1, sDateFormat2

	' Check For empty/null values submited from web page form.

	sSQL = _
		"INSERT INTO [EventType] (" & _
			"[Description] " & _
		") VALUES (" & _
			"'" & Replace(Request("Description"),"'","''") & "')"
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

	' Check For empty/null values submited from web page form.

	sSQL = _
	"UPDATE [EventType] " & _
	"SET " & _
	"[Description]='" & Replace(Request("Description"),"'","''") & "' "
	sSQL = sSQL & "WHERE [eventTypeID]=" & Request("PK_ieventTypeID") & ""
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
		"DELETE FROM [EventType] " & _
		"WHERE [eventTypeID]=" & Request("PK_ieventTypeID") & ""
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

' Update Record/s From The Database
Function ResultsUpdate()
	Dim bReturn, Item

	' Transactional database processing - each update is "buffered". 
	' If all requested updates are good then they are executed, otherwise they are canceled or rolled back.
	' Using a transaction is optional, but its safer with one
	oConn.BeginTrans

	iCounter = 0 
	For Each Item In Request.Form
		If Len(Request(Item)) > 0 Then
			If Left(Item, 3) = "row" Then
				iCounter = iCounter + 1
				' Check For empty/null values submited from web page form.

				sSQL=_
					"UPDATE [EventType] " & _
					"SET " & _
						"[Description]='" & Replace(Request("Description" & Mid(Item, 4)),"'","''") & "' "
				sSQL = sSQL & "WHERE " & _
					Replace(Request("row" & Mid(Item, 4)), "__", " ")
				'Response.Write sSQL & "<HR width=300>": Response.Flush

				oConn.Execute(sSQL)
			End If
		End If
	Next

	If Err Then
		oConn.RollbackTrans
		bReturn = iCounter
	Else
		Application("EventTypeeventTypeIDLookup") = ""
		Application("ExecPositionexecPositionIDLookup") = ""
		oConn.CommitTrans
		bReturn = iCounter
	End If

	Set oConn = Nothing
	ResultsUpdate = bReturn
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
				sSQL = "DELETE FROM [EventType] " & _
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

	' Check For empty/null values submited from web page form.

	sSQL = _
		"INSERT INTO [EventType] (" & _
			"[Description] " & _
		") VALUES (" & _
			"'" & Replace(Request("AddDescription"),"'","''") & "')"
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
