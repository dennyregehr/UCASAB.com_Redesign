<%
Function FillEventTypeeventTypeIDLookup()
	Dim oRSLookup
	Dim aReturn()
	Dim i

	sSQL = _
		"SELECT [eventTypeID], [Description] " & _
		"FROM [EventType] " & _
		"ORDER BY [eventTypeID]"

	Set oRSLookup = oConn.Execute(sSQL)

	If oRSLookup.EOF Then

	Else
		i = -1
		Do Until oRSLookup.EOF
			i = i + 1
			ReDim Preserve aReturn(2, i)
			aReturn(0, i) = oRSLookup("eventTypeID")
			aReturn(1, i) = oRSLookup("Description")
			aReturn(2, i) = ""
			oRSLookup.MoveNext
		Loop
	End If
	Set oRSLookup = Nothing
	Application("EventTypeeventTypeIDLookup") = aReturn
End Function

Function LookupEventTypeeventTypeID(pviSelect, pvsType, pviEdit, pviCntr, pvsName)
	Dim sReturn
	Dim aEventType
	Dim i
	Dim bFound
	Dim sName

	If Not IsArray(Application("EventTypeeventTypeIDLookup")) Then
		FillEventTypeeventTypeIDLookup
	End If

	aEventType = Application("EventTypeeventTypeIDLookup")
	If pvsType = "Value" Then
		bFound = False
		For i = 0 To UBound(aEventType, 2)
			If Not IsEmpty(pviSelect) And Not IsNull(pviSelect) Then
				If CStr(pviSelect) = CStr(aEventType(0, i)) Then ' ????? compare data type
					sReturn = aEventType(1, i) & " " & aEventType(2, i)
					bFound = True
					Exit For
				Else
					sReturn = ""
				End If
			End If
		Next
		If pviSelect <> "" And Not bFound Then
			sReturn = pviSelect
		End If
	Else
		If pviEdit = 2 Then
			sName = "AddEventTypeID"
		Else
			If pviCntr > 0 Then
				sName = "EventTypeID" & pviCntr
			Else
				sName = pvsName
			End If

		End If

		sReturn = "<SELECT name=""" & sName & """>"
		If pviEdit = 1 Then
			sReturn = sReturn & "<OPTION value="""">"
		Else
			sReturn = sReturn & "<OPTION value="""">Select"
		End If

		bFound = False
		For i = 0 To UBound(aEventType, 2)
			If IsEmpty(pviSelect) Or IsNull(pviSelect) Then
				sReturn = sReturn & "<OPTION value=""" & aEventType(0, i) & """>" & aEventType(1, i) & " " & aEventType(2, i)
			Else
				If CStr(pviSelect) = CStr(aEventType(0, i)) Then ' ????? compare data type
					sReturn = sReturn & "<OPTION value=""" & aEventType(0, i) & """ selected>" & aEventType(1, i) & " " & aEventType(2, i)
					bFound = True
				Else
					sReturn = sReturn & "<OPTION value=""" & aEventType(0, i) & """>" & aEventType(1, i) & " " & aEventType(2, i)
				End If
			End If
		Next

		sReturn = sReturn & "</SELECT>"
		If pviSelect <> "" And Not bFound Then
			sReturn = pviSelect & "<BR>" & sReturn
		End If

	End If
	aEventType = ""
	LookupEventTypeeventTypeID = sReturn
End Function

Function FillExecPositionexecPositionIDLookup()
	Dim oRSLookup
	Dim aReturn()
	Dim i

	sSQL = _
		"SELECT [execPositionID], [Position] " & _
		"FROM [ExecPosition] " & _
		"ORDER BY [execPositionID]"

	Set oRSLookup = oConn.Execute(sSQL)

	If oRSLookup.EOF Then

	Else
		i = -1
		Do Until oRSLookup.EOF
			i = i + 1
			ReDim Preserve aReturn(2, i)
			aReturn(0, i) = oRSLookup("execPositionID")
			aReturn(1, i) = oRSLookup("Position")
			aReturn(2, i) = ""
			oRSLookup.MoveNext
		Loop
	End If
	Set oRSLookup = Nothing
	Application("ExecPositionexecPositionIDLookup") = aReturn
End Function

Function LookupExecPositionexecPositionID(pviSelect, pvsType, pviEdit, pviCntr, pvsName)
	Dim sReturn
	Dim aExecPosition
	Dim i
	Dim bFound
	Dim sName

	If Not IsArray(Application("ExecPositionexecPositionIDLookup")) Then
		FillExecPositionexecPositionIDLookup
	End If

	aExecPosition = Application("ExecPositionexecPositionIDLookup")
	If pvsType = "Value" Then
		bFound = False
		For i = 0 To UBound(aExecPosition, 2)
			If Not IsEmpty(pviSelect) And Not IsNull(pviSelect) Then
				If CStr(pviSelect) = CStr(aExecPosition(0, i)) Then ' ????? compare data type
					sReturn = aExecPosition(1, i) & " " & aExecPosition(2, i)
					bFound = True
					Exit For
				Else
					sReturn = ""
				End If
			End If
		Next
		If pviSelect <> "" And Not bFound Then
			sReturn = pviSelect
		End If
	Else
		If pviEdit = 2 Then
			sName = "Addposition"
		Else
			If pviCntr > 0 Then
				sName = "position" & pviCntr
			Else
				sName = pvsName
			End If

		End If

		sReturn = "<SELECT name=""" & sName & """>"
		If pviEdit = 1 Then
			sReturn = sReturn & "<OPTION value="""">"
		Else
			sReturn = sReturn & "<OPTION value="""">Select"
		End If

		bFound = False
		For i = 0 To UBound(aExecPosition, 2)
			If IsEmpty(pviSelect) Or IsNull(pviSelect) Then
				sReturn = sReturn & "<OPTION value=""" & aExecPosition(0, i) & """>" & aExecPosition(1, i) & " " & aExecPosition(2, i)
			Else
				If CStr(pviSelect) = CStr(aExecPosition(0, i)) Then ' ????? compare data type
					sReturn = sReturn & "<OPTION value=""" & aExecPosition(0, i) & """ selected>" & aExecPosition(1, i) & " " & aExecPosition(2, i)
					bFound = True
				Else
					sReturn = sReturn & "<OPTION value=""" & aExecPosition(0, i) & """>" & aExecPosition(1, i) & " " & aExecPosition(2, i)
				End If
			End If
		Next

		sReturn = sReturn & "</SELECT>"
		If pviSelect <> "" And Not bFound Then
			sReturn = pviSelect & "<BR>" & sReturn
		End If

	End If
	aExecPosition = ""
	LookupExecPositionexecPositionID = sReturn
End Function

%>

