<%
' ***********************************
' Coded By: DB2ASP v4.3.5 (12) - A 303 Media Company
' For More Information:
' Please Visit http://www.DB2ASP.com
' Or Email support@303media.com
' ***********************************
' Page Name: Calendar_results.asp
' Date: 3/4/2005 12:28:24 AM
' Purpose: Enumerate Records From Table "Calendar"
' Database: Access
' Table: Calendar
' ***********************************

%>
<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<%
Response.CacheControl = "no-cache"
Response.Expires = -10
Response.AddHeader "cache-control", "no-cache"
Response.AddHeader "pragma", "no-cache"



sObjectName = "Calendar"
gsResultsPageName = "Calendar_results.asp"
sSubTitle = "List All Records"

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Request("view") = "listall" Then
	Session(sObjectName & "LSearch") = ""
	Session(sObjectName & "LSort") = ""
	Session(sObjectName & "LPage") = ""
	Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LSearch") = ""
	Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LSort") = ""
	Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LPage") = ""
	Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
Else
	If Request("view") = "newsearch" Then
		Session(sObjectName & "LSearch") = ""
		Session(sObjectName & "LPage") = ""
		Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LSearch") = ""
		Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LPage") = ""
		Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"

			If Request("EventID") <> "" Then
				If Request("EventID2") <> "" Then
					sWhere = sWhere & "[EventID] BETWEEN " & Request("EventID") & " AND " & Request("EventID2") & " " & Request("cndxx1") & " "
				Else
					sWhere = sWhere & "[EventID]=" & Request("EventID") & " " & Request("cndxx1") & " "
				End If
			End If

			If Request("EventName") <> "" Then
				sWhere = sWhere & "[EventName] LIKE '%" & Request("EventName") & "%' " & Request("cndxx2") & " "
			End If

			If Request("EventTypeID") <> "" Then
				If Request("EventTypeID2") <> "" Then
					sWhere = sWhere & "[EventTypeID] BETWEEN " & Request("EventTypeID") & " AND " & Request("EventTypeID2") & " " & Request("cndxx3") & " "
				Else
					sWhere = sWhere & "[EventTypeID]=" & Request("EventTypeID") & " " & Request("cndxx3") & " "
				End If
			End If

			If Request("Location") <> "" Then
				sWhere = sWhere & "[Location] LIKE '%" & Request("Location") & "%' " & Request("cndxx4") & " "
			End If

			If Request("StartDate") <> "" Then
				sWhere = sWhere & "[StartDate] BETWEEN #" & Request("StartDate") & "# AND #" & Request("StartDate2") & "# " & Request("cndxx5") & " "
			End If

			If Request("EndDate") <> "" Then
				sWhere = sWhere & "[EndDate] BETWEEN #" & Request("EndDate") & "# AND #" & Request("EndDate2") & "# " & Request("cndxx6") & " "
			End If

			If Request("StartTime") <> "" Then
				sWhere = sWhere & "[StartTime] BETWEEN #" & Request("StartTime") & "# AND #" & Request("StartTime2") & "# " & Request("cndxx7") & " "
			End If

			If sWhere <> "" Then sSubTitle = "Search Results"
			If Right(sWhere, 3) = "ND " Or Right(sWhere, 3) = "OR " Then sWhere = Left(sWhere, Len(sWhere) - 4)
			Session(sObjectName & "LSearch") = sWhere
			Response.Cookies("DB2ASP_ADMIN")(sObjectName & "LSearch") = sWhere
			Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"

			If sWhere <> "" Then
				sWhere = "WHERE " & sWhere & " "
			End If
			Call RememberSearchValues
	Else
		If Request("view") = "related" Then
			sSubTitle = "Related Records"

		Else
			sSubTitle = "Search Results"
			sWhere = GetLastSearch(sObjectName)
		End If
	End If
End If

sSort = GetLastSort(sObjectName)

If sSort = "" Then
	'sSQL = sSQL & " ORDER BY "  ' Unrem this and add field name for a default sort
End If

iMaxRecords = 5000

iPageSize = GetLastPageSize(sObjectName)
If iPageSize < 5 Then
	iPageSize = 50
End If

sSQL = _
	"SELECT TOP " & iMaxRecords & " * " & _
	"FROM [Calendar] " & " " & _
	sWhere & _
	sSort

If sSQL <> "" Then
	' This is a security precaution
	If VerifySQL(sSQL) = "" Then
		Response.Redirect "Calendar_results.asp?view=listall"
	End If
End If

If Request("err") <> "" Then sErr = Request("err")

%>
<!-- #INCLUDE file="includes/common/DB2ASP_results_header.asp" -->
<HTML>
<HEAD>
	<META name="GENERATOR" content="303 Media's DB2ASP v4.3.5"/>
	<TITLE>Calendar - Results</TITLE>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=results action="Calendar_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>" method="POST">
<TABLE cellspacing=1 class=DB2ASP>
	<TR>
		<TD colspan=9 class=DB2ASPlight>
			<TABLE cellspacing=1 class=menu width=100%><TR>
			<TD class=menu>
				<SPAN class=DB2ASPtitle>Calendar - <% If sWhere <> "" Then Response.Write "Search Results" Else Response.Write "List All Records" End If %></SPAN>
			</TD>
			<TD align=right class=menu width=120>
				<A href="Calendar_search.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>" title="Search For Specific Records" class=DB2ASPvalue>Search</A><br>
				<A href="Calendar_results.asp?view=listall<%=GetQueryString("ListAll", Empty, Empty, Empty, Empty)%>" title="List All Records" class=DB2ASPvalue>List All</A><br>
			</TD></TR></TABLE>
		</TD>
	</TR>
	<TR>
		<TH colspan=9 class=DB2ASPdark align=left>
			<%
			If iTotalRecords = 0 Then
				Response.Write "No records found matching your request."
			ElseIf iTotalRecords < CInt(iPageSize) + 1 Then
				Response.Write "Below are all results found (" & iTotalRecords & ")"
			Else
				If iTotalRecords >= iMaxRecords Then
					Response.Write "Below are results " & iStart & " through " & iFinish & " of the first " & iMaxRecords & " records."
				Else
					Response.Write "Below are results " & iStart & " through " & iFinish & " of " & iTotalRecords & " found. "
				End If
			End If
			%>
		 &nbsp;&nbsp;&nbsp;&nbsp; Records Per Page: <%=GetPageSizes(iPageSize)%>
		</TH>
	</TR>
	<%=GetCriteria()%>

<% If iTotalRecords > 0 Then %>
	<% If Request("msg") <> "" Then Response.Write "<TR><TD colspan=9 class=DB2ASPlight><SPAN class=DB2ASPmessage>" & Request("msg") & "</SPAN></TD></TR>" End If%>
	<% If Request("err") <> "" Then Response.Write "<TR><TD colspan=9 class=DB2ASPlight><SPAN class=DB2ASPerror>" & Request("err") & "</SPAN></TD></TR>" End If%>
	<TR>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[EventID]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By EventID" class=DB2ASPdarkA>EventID</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[EventName]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By EventName" class=DB2ASPdarkA>EventName</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[EventTypeID]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By EventTypeID" class=DB2ASPdarkA>EventTypeID</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[Location]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Location" class=DB2ASPdarkA>Location</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[StartDate]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By StartDate" class=DB2ASPdarkA>StartDate</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[EndDate]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By EndDate" class=DB2ASPdarkA>EndDate</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="Calendar_results.asp<%=GetQueryString("Results", "[StartTime]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By StartTime" class=DB2ASPdarkA>StartTime</A>
		</TH>
		<TH class=DB2ASPdark><INPUT name="chkall" type="checkbox" value="Check All" onClick="CheckAll();" title="Mark All Records On This Page To Be Deleted"> <SPAN class=DB2ASPsmall>Check All</SPAN></TH>
	</TR>
<%
	For iCounter = 1 To oRS.PageSize
		If sRowColor = " class=DB2ASPlight" Then
			sRowColor = " class=DB2ASPdark"
		Else
			sRowColor = " class=DB2ASPlight"
		End If
		%>
		<TR>
		<TD<%=sRowColor%>>
			<A href="Calendar_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, "PK_iEventID=" & oRS("EventID") & "")%>" title="Edit Record" <%=sRowColor & "A"%>><%=oRS("EventID")%></A>
			<INPUT type=hidden name="row<%=iCounter%>" value="EventID=<%=oRS("EventID")%>">
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("EventName")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
			<%=LookupEventTypeeventTypeID(oRS("EventTypeID"), "Value", 0, Empty, "EventTypeID")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("Location")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<% If Not IsNull(oRS("StartDate")) Then Response.Write FormatDateTime(oRS("StartDate"), 2)%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<% If Not IsNull(oRS("EndDate")) Then Response.Write FormatDateTime(oRS("EndDate"), 2)%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<% If Not IsNull(oRS("StartTime")) Then Response.Write FormatDateTime(oRS("StartTime"), 3)%>&nbsp;
		</TD>
		<TD align=center<%=sRowColor%>>
			<INPUT type=checkbox name=chkd<%=iCounter%> value="EventID=<%=oRS("EventID")%>" title="Mark Record To Be Deleted">
			<A href="Calendar_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, "PK_iEventID=" & oRS("EventID") & "")%>" title="Edit Record" <%=sRowColor & "A"%>>detail</A>
		</TD>
		</TR>
		<%
		oRS.MoveNext
		If oRS.EOF Then Exit For
	Next

	If iPageCount > 1 Then
		Response.Write "<TR><TH colspan=9 class=DB2ASPdark>"

		Response.Write "Go To Page :&nbsp;&nbsp;&nbsp;"
	End If

	If CInt(iPage) > 1 Then
		Response.Write "<A href=""Calendar_results.asp" & GetQueryString("Results", Empty, iPage-1, Empty, Empty) & """ title=""Show Previous Page Of Records"" " & sRowColor & "A" & "><B>Previous</B></A>&nbsp;&nbsp;"
	End If

	If iPageCount > 1 Then
		For iCounter = 1 To iPageCount
			If CInt(iCounter) = CInt(iPage) Then
				Response.Write "[" & iCounter & "] "
			Else
				Response.Write "<A href=""Calendar_results.asp" & GetQueryString("Results", Empty, iCounter, Empty, Empty) & """ " & sRowColor & "A" & ">" & iCounter & "</A> "
			End If
			If iCounter > 9 Then
				Response.Write "..."
				Exit For
			End If
		Next
		If iPage > 10 Then
			Response.Write " [" & iPage & "] "
		End If
	End If
	If CInt(iPage) < CInt(iPageCount) Then
		Response.Write "&nbsp;&nbsp;<A href=""Calendar_results.asp" & GetQueryString("Results", Empty, iPage+1, Empty, Empty) & """ title=""Show Next Page Of Records"" " & sRowColor & "A" & "><B>Next</B></A>"
	End If 
	Response.Write "</TH></TR>"
End If 
%>
<TR><TD colspan=9 class=DB2ASPlight>To sort records, click the column name, click again to reverse.</TD></TR><BR>
</TABLE>
<BR>
<!--<A href='Calendar_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>'>Add New Record</A>&nbsp;&nbsp;-->
<INPUT type=Submit name=btnAdd value="Add New Record" onClick="javascript: document.location.href='Calendar_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>&fnc=add'; return false;" class=DB2ASPactionbtn>
<INPUT type=Submit name=btnDelete value="Delete Checked Records" onClick="javascript: if(confirm('Are You Sure You Want To Delete Checked Record/s?')==false) return false;" class=DB2ASPactionbtn> 
</FORM>
<!--END--42821-->
<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_lookups.asp" -->
<%
Sub RememberSearchValues()
	' This procedure can be omitted, but user functionality will be reduced
	Session("vEventName") = Request("EventName")
	Session("vEventTypeID") = Request("EventTypeID")
	Session("vEventTypeID2") = Request("EventTypeID2")
	Session("vLocation") = Request("Location")
	Session("vStartDate") = Request("StartDate")
	Session("vStartDate2") = Request("StartDate2")
	Session("vEndDate") = Request("EndDate")
	Session("vEndDate2") = Request("EndDate2")
	Session("vEventDescription") = Request("EventDescription")
	Session("vCnd1") = Request("vCnd1")
	Session("vCnd2") = Request("vCnd2")
	Session("vCnd3") = Request("vCnd3")
	Session("vCnd4") = Request("vCnd4")
	Session("vCnd5") = Request("vCnd5")
	Session("vCnd6") = Request("vCnd6")
End Sub

Function GetCriteria()
	Dim sReturn
	If sWhere <> "" Then
		sReturn = sReturn & "<SPAN class=DB2ASPhighlight>Search Criteria:</SPAN> "
		If InStr(sWhere, "EventID") > 0 Then sReturn = sReturn & "EventID, "
		If InStr(sWhere, "EventName") > 0 Then sReturn = sReturn & "EventName, "
		If InStr(sWhere, "EventTypeID") > 0 Then sReturn = sReturn & "EventTypeID, "
		If InStr(sWhere, "Location") > 0 Then sReturn = sReturn & "Location, "
		If InStr(sWhere, "StartDate") > 0 Then sReturn = sReturn & "StartDate, "
		If InStr(sWhere, "EndDate") > 0 Then sReturn = sReturn & "EndDate, "
		If InStr(sWhere, "StartTime") > 0 Then sReturn = sReturn & "StartTime, "
	End If
	If sSort <> "" Then
		sReturn = sReturn & "<SPAN class=DB2ASPhighlight>Sorted By:</SPAN> " '& Replace(sSort, " ORDER BY ", "") & ", "
		If InStr(sSort, "EventID") > 0 Then sReturn = sReturn & "EventID "
		If InStr(sSort, "EventName") > 0 Then sReturn = sReturn & "EventName "
		If InStr(sSort, "EventTypeID") > 0 Then sReturn = sReturn & "EventTypeID "
		If InStr(sSort, "Location") > 0 Then sReturn = sReturn & "Location "
		If InStr(sSort, "StartDate") > 0 Then sReturn = sReturn & "StartDate "
		If InStr(sSort, "EndDate") > 0 Then sReturn = sReturn & "EndDate "
		If InStr(sSort, "StartTime") > 0 Then sReturn = sReturn & "StartTime "
		If InStr(sSort, "ASC") > 0 Then sReturn = sReturn & "<IMG src=""images/DB2ASP/darrow.gif"" border=0 width=9 height=9>, "
		If InStr(sSort, "DESC") > 0 Then sReturn = sReturn & "<IMG src=""images/DB2ASP/uarrow.gif"" border=0 width=9 height=9>, "
	End If
	If iPage > 1 Then
		sReturn = sReturn & "<SPAN class=DB2ASPhighlight>Page:</SPAN> " & iPage & ", "
	End If

	If sReturn <> "" Then
		If Right(sReturn, 2) = ", " Then sReturn = Left(sReturn, Len(sReturn) - 2)
		sReturn = "<TR><TH colspan=9 class=DB2ASPdark align=left>" & sReturn & "</TH></TR>"
		GetCriteria = sReturn
	Else
		GetCriteria = ""
	End If
End Function
%>
