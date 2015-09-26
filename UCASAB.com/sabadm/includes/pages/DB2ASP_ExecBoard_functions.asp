<%
' Add A New Record To The Database
Function DetailAdd()
	Dim bReturn
	Dim sDateFormat1, sDateFormat2
	Dim varposition
	Dim varserviceDates
	Dim varactive

	' Check For empty/null values submited from web page form.
	If Request("position") <> "" Then
		varposition = Replace(Request("position"), "'", "''")
	Else
		varposition = "Null"
	End If

	If Request("serviceDates") <> "" Then
		varserviceDates ="#" & Replace(Request("serviceDates"), "'", "''") & "#"
	Else
		varserviceDates = "Null"
	End If

	If UCase(Request("active")) = "FALSE" Or UCase(Request("active")) = "NO" Then
		varactive = 0
	ElseIf UCase(Request("active")) = "TRUE" Or UCase(Request("active")) = "YES" Then
		varactive = 1
	ElseIf Request("active") = "" Then
		varactive = "Null"
	Else
		varactive = Request("active")
	End If

	sSQL = _
		"INSERT INTO [ExecBoard] (" & _
			"[fname], " & _
			"[lname], " & _
			"[email], " & _
			"[phone1], " & _
			"[phone2], " & _
			"[position], " & _
			"[serviceDates], " & _
			"[active] " & _
		") VALUES (" & _
			"'" & Replace(Request("fname"),"'","''") & "', " & _
			"'" & Replace(Request("lname"),"'","''") & "', " & _
			"'" & Replace(Request("email"),"'","''") & "', " & _
			"'" & Replace(Request("phone1"),"'","''") & "', " & _
			"'" & Replace(Request("phone2"),"'","''") & "', " & _
			"" & varposition & ", " & _
			"" & varserviceDates & ", " & _
			"" & varactive & ")"
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
	Dim varposition
	Dim varserviceDates
	Dim varactive

	' Check For empty/null values submited from web page form.
	If Request("position") <> "" Then
		varposition = Replace(Request("position"), "'", "''")
	Else
		varposition = "Null"
	End If

	If Request("serviceDates") <> "" Then
		varserviceDates = "#" & Replace(Request("serviceDates"), "'", "''") & "#"
	Else
		varserviceDates = "Null"
	End If

	If UCase(Request("active")) = "FALSE" Or UCase(Request("active")) = "NO" Then
		varactive = 0
	ElseIf UCase(Request("active")) = "TRUE" Or UCase(Request("active")) = "YES" Then
		varactive = 1
	ElseIf Request("active") = "" Then
		varactive = "Null"
	Else
		varactive = Request("active")
	End If

	sSQL = _
	"UPDATE [ExecBoard] " & _
	"SET " & _
	"[fname]='" & Replace(Request("fname"),"'","''") & "', " & _
	"[lname]='" & Replace(Request("lname"),"'","''") & "', " & _
	"[email]='" & Replace(Request("email"),"'","''") & "', " & _
	"[phone1]='" & Replace(Request("phone1"),"'","''") & "', " & _
	"[phone2]='" & Replace(Request("phone2"),"'","''") & "', " & _
	"[position]=" & varposition & ", " & _
	"[serviceDates]=" & varserviceDates & ", " & _
	"[active]=" & varactive & " "
	sSQL = sSQL & "WHERE [execID]=" & Request("PK_iexecID") & ""
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
		"DELETE FROM [ExecBoard] " & _
		"WHERE [execID]=" & Request("PK_iexecID") & ""
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
				sSQL = "DELETE FROM [ExecBoard] " & _
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
	Dim varposition
	Dim varserviceDates
	Dim varactive

	' Check For empty/null values submited from web page form.
	If Request("Addposition") <> "" Then
		varposition = Replace(Request("Addposition"), "'", "''")
	Else
		varposition = "Null"
	End If

	If Request("AddserviceDates") <> "" Then
		varserviceDates ="#" &Replace(Request("AddserviceDates"), "'", "''") & "#"
	Else
		varserviceDates = "Null"
	End If

	If UCase(Request("active")) = "FALSE" Or UCase(Request("active")) = "NO" Then
		varactive = 0
	ElseIf UCase(Request("Addactive")) = "TRUE" Or UCase(Request("Addactive")) = "YES" Then
		varactive = 1
	ElseIf Request("Addactive") = "" Then
		varactive = "Null"
	Else
		varactive = Request("Addactive")
	End If

	sSQL = _
		"INSERT INTO [ExecBoard] (" & _
			"[fname], " & _
			"[lname], " & _
			"[email], " & _
			"[position], " & _
			"[serviceDates], " & _
			"[active] " & _
		") VALUES (" & _
			"'" & Replace(Request("Addfname"),"'","''") & "', " & _
			"'" & Replace(Request("Addlname"),"'","''") & "', " & _
			"'" & Replace(Request("Addemail"),"'","''") & "', " & _
			"" & varposition & ", " & _
			"" & varserviceDates & ", " & _
			"" & varactive & ")"
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
