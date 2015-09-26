<%
' This function creates the correct syntax for href
Function Sort(pvstrSort)
	If Request("sort") = pvstrSort & " ASC" Then
		Sort = pvstrSort & "%20DESC"
	Else
		Sort = pvstrSort & "%20ASC"
	End If
End Function

Function DropdownSelection(pvvarDBValue, pvvarOptionValue)
	If Len(pvvarDBValue) > 0 And Len(pvvarOptionValue) > 0 Then
		If CStr(pvvarDBValue) = CStr(pvvarOptionValue) Then
			DropdownSelection = "selected"
		End If
	End If
End Function

Function RadioSelection(pvvarDBValue, pvvarRadioValue)
	If Len(pvvarDBValue) > 0 And Len(pvvarRadioValue) > 0 Then
		If CBool(pvvarDBValue) = CBool(pvvarRadioValue) Then
			RadioSelection = "checked"
		End If
	End If
End Function

Function ConditionRadio(pvInt, pvsLast)
	Dim sReturn, sState1, sState2

	If Trim(pvsLast) = " And " Then
		sState1 = "checked"
		sState2 = ""
	Else
		sState1 = ""
		sState2 = "checked"
	End If

	sReturn = "And <INPUT type=radio name=""cndxx" & pvInt & """ value=""AND"" " & sState1 & ">&nbsp;&nbsp;"
	sReturn = sReturn & "Or <INPUT type=radio name=""cndxx" & pvInt & """ value=""OR"" " & sState2 & ">"

	ConditionRadio = sReturn
End Function

Function ConditionOption(pvInt)
	Dim sReturn
	sReturn = "<SELECT name=""cndxx" & pvInt & """>"
	sReturn = sReturn & "<OPTION value=""AND"">AND"
	sReturn = sReturn & "<OPTION value=""OR"" " & IsOptionSelected(pvInt) & ">OR"
	sReturn = sReturn & "</SELECT>"
	Condition = sReturn
End Function

Function IsOptionSelected(pvInt2)
	If Request("cndxx" & pvInt2) = "OR" Then
		IsOptionSelected = "selected"
	Else
		IsOptionSelected = ""
	End If
End Function

Function GetPageSizes(pviCurrentValue)
	Dim sReturn
	sReturn = "<SELECT name=PageSize onChange=""javascript: document.location.href='" & gsResultsPageName & GetQueryString("ResultsPage", Empty, Empty, 0, Empty) & "'+document.results.PageSize.options[document.results.PageSize.selectedIndex].value;"">"
	sReturn = sReturn & "<OPTION value=5"
	If pviCurrentValue = 5 Then sReturn = sReturn & " selected"
	sReturn = sReturn & ">5"
	sReturn = sReturn & "<OPTION value=10"
	If pviCurrentValue = 10 Then sReturn = sReturn & " selected"
	sReturn = sReturn & ">10"
	sReturn = sReturn & "<OPTION value=25"
	If pviCurrentValue = 25 Then sReturn = sReturn & " selected"
	sReturn = sReturn & ">25"
	sReturn = sReturn & "<OPTION value=50"
	If pviCurrentValue = 50 Then sReturn = sReturn & " selected"
	sReturn = sReturn & ">50"
	sReturn = sReturn & "<OPTION value=100"
	If pviCurrentValue = 100 Then sReturn = sReturn & " selected"
	sReturn = sReturn & ">100"
	sReturn = sReturn & "</SELECT>"
	GetPageSizes = sReturn
End Function

Function GetLastSearch(pvsObjectName)
	Dim sReturn

	If Len(Request("where")) > 0 Then
		sReturn = Request("where")
		Session(pvsObjectName & "LSearch") = sReturn
		Response.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSearch") = sReturn
		Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
		sReturn = "WHERE " & sReturn
	Else
		If Len(Session(pvsObjectName & "LSearch")) > 0 Then
			sReturn = "WHERE " & Session(pvsObjectName & "LSearch")
		Else
			If Len(Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSearch")) > 0 Then
				Session(pvsObjectName & "LSearch") = sReturn
				sReturn = Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSearch")
				sReturn = "WHERE " & sReturn
			Else
				sReturn = ""
			End If
		End If
	End If

	GetLastSearch = sReturn
End Function

Function GetLastSort(pvsObjectName)
	Dim sReturn

	If Len(Request.QueryString("sort")) > 0 Then
		' User Is Specifying A New Sort
		sReturn = Request.QueryString("sort")
		Session(pvsObjectName & "LSort") = sReturn
		Response.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSort") = sReturn
		Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
		sReturn = " ORDER BY " & sReturn
	Else
		If Len(Session(pvsObjectName & "LSort")) > 0 Then
		' Get Last Used Sort
			sReturn = " ORDER BY " & Session(pvsObjectName & "LSort")
		Else
			If Len(Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSort")) > 0 Then
				sReturn = Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LSort")
				Session(pvsObjectName & "LSort") = sReturn
				sReturn = " ORDER BY " & sReturn
			Else
				sReturn = ""
			End If
		End If
	End If
	GetLastSort = sReturn
End Function

' If anything has been writen to client/browser, the write cookie will fail
Function GetLastPage(pvsObjectName) ' Called From DB2ASP_results_header.asp
	Dim sReturn

	' ##### make more universal "Customers" as var
	If Len(Request.QueryString("rpage")) > 0 Then
		sReturn = Request.QueryString("rpage")
		Session(pvsObjectName & "LPage") = sReturn
		Response.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPage") = sReturn
		Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
	Else
		If Len(Session(pvsObjectName & "LPage")) > 0 Then
			sReturn = Session(pvsObjectName & "LPage")
		Else
			If Len(Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPage")) > 0 Then
				sReturn = Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPage")
				Session(pvsObjectName & "LPage") = sReturn
			Else
				sReturn = 1
			End If
		End If
	End If
	GetLastPage = sReturn
End Function

Function GetLastPageSize(pvsObjectName)  ' Called From DB2ASP_environ.asp
	Dim sReturn

	' ##### make more universal "Customers" as var
	If Len(Request.QueryString("pagesize")) > 0 Then
		sReturn = Request.QueryString("pagesize")
		Session(pvsObjectName & "LPageSize") = sReturn
		Response.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPageSize") = sReturn
		Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
	Else
		If Len(Session(pvsObjectName & "LPageSize")) > 0 Then
			sReturn = Session(pvsObjectName & "LPageSize")
		Else
			If Len(Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPageSize")) > 0 Then
				sReturn = Request.Cookies("DB2ASP_ADMIN")(pvsObjectName & "LPageSize")
				Session(pvsObjectName & "LPageSize") = sReturn
			Else
				sReturn = 10 ' Default Page Size
			End If
		End If
	End If
	GetLastPageSize = sReturn
End Function

Function VerifySQL(pvsSQL)
	' This is a security precaution
	If InStr(UCase(pvsSQL), "UPDATE ") > 0 Or InStr(UCase(pvsSQL), "INSERT ") > 0 Or InStr(UCase(pvsSQL), "DELETE ") > 0 Then
		sErr = "Can't Retrive Data (Bad Query)"
		VerifySQL = ""
	Else
		VerifySQL = pvsSQL
	End If
End Function

Function GetQueryString(pvsPage, pvsSortColumn, pviPage, pviPageSize, pvsKey)
	Dim sReturn

	If pvsPage = "Nav" or pvsPage = "NewSearch" Or pvsPage = "ListAll" Then
		sReturn = "&"
	Else
		sReturn = "?"
	End If

	If pvsPage = "Results" Then sReturn = sReturn & "ab=" & oRS.AbsolutePosition & "&"

	If pvsKey <> "" Then
		sReturn = sReturn & pvsKey & "&"
	End If

	If pvsPage = "DetailAdd" Then sReturn = sReturn & "fnc=add"&"&"

	If pvsSortColumn <> "" Then
		sReturn = sReturn & "sort=" & Sort(pvsSortColumn) & "&"
	Else
		If Len(Session(sObjectName & "LSort")) < 1 And sSort <> "" Then
			sSort = Replace(sSort, "ORDER BY ", "")
			sReturn = sReturn & "sort=" & Server.URLEncode(sSort) & "&"  'Server.URLEncode(sSort)
		End If
	End If

	If Len(Session(sObjectName & "LSearch")) < 1 And sWhere <> "" Then
		sWhere = Replace(sWhere, "WHERE ", "")
		sReturn = sReturn & "where=" & Trim(Server.URLEncode(sWhere)) & "&" 'Server.URLEncode(sWhere)
	End If

	If pviPage > 0 Then
		sReturn = sReturn & "rpage=" & pviPage & "&"
	Else
		If Len(Session(sObjectName & "LPage")) < 1 And iPage > 0 Then sReturn = sReturn & "rpage=" & iPage & "&"
	End If

	If pviPageSize = 0 Then
		sReturn = sReturn & "pagesize="
	Else
		If Len(Session(sObjectName & "LPageSize")) < 1 Then sReturn = sReturn & "pagesize=" & iPageSize
	End If

	If Len(sReturn) = 1 Then
		sReturn = ""
	End If

	If Right(sReturn, 1) = "&" Then
		sReturn = Left(sReturn, Len(sReturn) - 1)
	End If

	GetQueryString = sReturn
End Function
%>

