<%
' ***********************************
' Coded By: DB2ASP v4.3.5 (12) - A 303 Media Company
' For More Information:
' Please Visit http://www.DB2ASP.com
' Or Email support@303media.com
' ***********************************
' Page Name: ExecBoard_results.asp
' Date: 3/4/2005 12:28:30 AM
' Purpose: Enumerate Records From Table "ExecBoard"
' Database: Access
' Table: ExecBoard
' ***********************************

%>
<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<%
Response.CacheControl = "no-cache"
Response.Expires = -10
Response.AddHeader "cache-control", "no-cache"
Response.AddHeader "pragma", "no-cache"



sObjectName = "ExecBoard"
gsResultsPageName = "ExecBoard_results.asp"
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

			If Request("execID") <> "" Then
				If Request("execID2") <> "" Then
					sWhere = sWhere & "[execID] BETWEEN " & Request("execID") & " AND " & Request("execID2") & " " & Request("cndxx1") & " "
				Else
					sWhere = sWhere & "[execID]=" & Request("execID") & " " & Request("cndxx1") & " "
				End If
			End If

			If Request("fname") <> "" Then
				sWhere = sWhere & "[fname] LIKE '%" & Request("fname") & "%' " & Request("cndxx2") & " "
			End If

			If Request("lname") <> "" Then
				sWhere = sWhere & "[lname] LIKE '%" & Request("lname") & "%' " & Request("cndxx3") & " "
			End If

			If Request("email") <> "" Then
				sWhere = sWhere & "[email] LIKE '%" & Request("email") & "%' " & Request("cndxx4") & " "
			End If

			If Request("position") <> "" Then
				If Request("position2") <> "" Then
					sWhere = sWhere & "[position] BETWEEN " & Request("position") & " AND " & Request("position2") & " " & Request("cndxx5") & " "
				Else
					sWhere = sWhere & "[position]=" & Request("position") & " " & Request("cndxx5") & " "
				End If
			End If

			If Request("serviceDates") <> "" Then
				sWhere = sWhere & "[serviceDates] BETWEEN #" & Request("serviceDates") & "# AND #" & Request("serviceDates2") & "# " & Request("cndxx6") & " "
			End If

			If Request("active") = "1" Then
				sWhere = sWhere & "[active] = True " & Request("cndxx7") & " "
			End If

			If Request("active") = "0" Then
				sWhere = sWhere & "[active] = False " & Request("cndxx7") & " "
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
	"FROM [ExecBoard] " & " " & _
	sWhere & _
	sSort

If sSQL <> "" Then
	' This is a security precaution
	If VerifySQL(sSQL) = "" Then
		Response.Redirect "ExecBoard_results.asp?view=listall"
	End If
End If

If Request("err") <> "" Then sErr = Request("err")

%>
<!-- #INCLUDE file="includes/common/DB2ASP_results_header.asp" -->
<HTML>
<HEAD>
	<META name="GENERATOR" content="303 Media's DB2ASP v4.3.5"/>
	<TITLE>Exec Board - Results</TITLE>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=results action="ExecBoard_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>" method="POST">
<TABLE cellspacing=1 class=DB2ASP>
	<TR>
		<TD colspan=9 class=DB2ASPlight>
			<TABLE cellspacing=1 class=menu width=100%><TR>
			<TD class=menu>
				<SPAN class=DB2ASPtitle>Exec Board - <% If sWhere <> "" Then Response.Write "Search Results" Else Response.Write "List All Records" End If %></SPAN>
			</TD>
			<TD align=right class=menu width=120>
				<A href="ExecBoard_search.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>" title="Search For Specific Records" class=DB2ASPvalue>Search</A><br>
				<A href="ExecBoard_results.asp?view=listall<%=GetQueryString("ListAll", Empty, Empty, Empty, Empty)%>" title="List All Records" class=DB2ASPvalue>List All</A><br>
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
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[execID]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Exec ID" class=DB2ASPdarkA>Exec ID</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[fname]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By First Name" class=DB2ASPdarkA>First Name</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[lname]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Last Name" class=DB2ASPdarkA>Last Name</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[email]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Email" class=DB2ASPdarkA>Email</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[position]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Position" class=DB2ASPdarkA>Position</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[serviceDates]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Service Dates" class=DB2ASPdarkA>Service Dates</A>
		</TH>
		<TH align=left class=DB2ASPdark>
			<A href="ExecBoard_results.asp<%=GetQueryString("Results", "[active]", Empty, Empty, Empty)%>&view=newsort" title="Sort Records By Active" class=DB2ASPdarkA>Active</A>
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
			<A href="ExecBoard_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, "PK_iexecID=" & oRS("execID") & "")%>" title="Edit Record" <%=sRowColor & "A"%>><%=oRS("execID")%></A>
			<INPUT type=hidden name="row<%=iCounter%>" value="execID=<%=oRS("execID")%>">
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("fname")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("lname")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("email")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
			<%=LookupExecPositionexecPositionID(oRS("position"), "Value", 0, Empty, "position")%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<% If Not IsNull(oRS("serviceDates")) Then Response.Write FormatDateTime(oRS("serviceDates"), 2)%>&nbsp;
		</TD>
		<TD<%=sRowColor%>>
				<%=oRS("active")%>&nbsp;
		</TD>
		<TD align=center<%=sRowColor%>>
			<INPUT type=checkbox name=chkd<%=iCounter%> value="execID=<%=oRS("execID")%>" title="Mark Record To Be Deleted">
			<A href="ExecBoard_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, "PK_iexecID=" & oRS("execID") & "")%>" title="Edit Record" <%=sRowColor & "A"%>>detail</A>
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
		Response.Write "<A href=""ExecBoard_results.asp" & GetQueryString("Results", Empty, iPage-1, Empty, Empty) & """ title=""Show Previous Page Of Records"" " & sRowColor & "A" & "><B>Previous</B></A>&nbsp;&nbsp;"
	End If

	If iPageCount > 1 Then
		For iCounter = 1 To iPageCount
			If CInt(iCounter) = CInt(iPage) Then
				Response.Write "[" & iCounter & "] "
			Else
				Response.Write "<A href=""ExecBoard_results.asp" & GetQueryString("Results", Empty, iCounter, Empty, Empty) & """ " & sRowColor & "A" & ">" & iCounter & "</A> "
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
		Response.Write "&nbsp;&nbsp;<A href=""ExecBoard_results.asp" & GetQueryString("Results", Empty, iPage+1, Empty, Empty) & """ title=""Show Next Page Of Records"" " & sRowColor & "A" & "><B>Next</B></A>"
	End If 
	Response.Write "</TH></TR>"
End If 
%>
<TR><TD colspan=9 class=DB2ASPlight>To sort records, click the column name, click again to reverse.</TD></TR><BR>
</TABLE>
<BR>
<!--<A href='ExecBoard_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>'>Add New Record</A>&nbsp;&nbsp;-->
<INPUT type=Submit name=btnAdd value="Add New Record" onClick="javascript: document.location.href='ExecBoard_detail.asp<%=GetQueryString("Results", Empty, Empty, Empty, Empty)%>&fnc=add'; return false;" class=DB2ASPactionbtn>
<INPUT type=Submit name=btnDelete value="Delete Checked Records" onClick="javascript: if(confirm('Are You Sure You Want To Delete Checked Record/s?')==false) return false;" class=DB2ASPactionbtn> 
</FORM>
<!--END--42821-->
<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_lookups.asp" -->
<%
Sub RememberSearchValues()
	' This procedure can be omitted, but user functionality will be reduced
	Session("vfname") = Request("fname")
	Session("vlname") = Request("lname")
	Session("vemail") = Request("email")
	Session("vposition") = Request("position")
	Session("vposition2") = Request("position2")
	Session("vCnd1") = Request("vCnd1")
	Session("vCnd2") = Request("vCnd2")
	Session("vCnd3") = Request("vCnd3")
	Session("vCnd4") = Request("vCnd4")
End Sub

Function GetCriteria()
	Dim sReturn
	If sWhere <> "" Then
		sReturn = sReturn & "<SPAN class=DB2ASPhighlight>Search Criteria:</SPAN> "
		If InStr(sWhere, "execID") > 0 Then sReturn = sReturn & "Exec ID, "
		If InStr(sWhere, "fname") > 0 Then sReturn = sReturn & "First Name, "
		If InStr(sWhere, "lname") > 0 Then sReturn = sReturn & "Last Name, "
		If InStr(sWhere, "email") > 0 Then sReturn = sReturn & "Email, "
		If InStr(sWhere, "position") > 0 Then sReturn = sReturn & "Position, "
		If InStr(sWhere, "serviceDates") > 0 Then sReturn = sReturn & "Service Dates, "
		If InStr(sWhere, "active") > 0 Then sReturn = sReturn & "Active, "
	End If
	If sSort <> "" Then
		sReturn = sReturn & "<SPAN class=DB2ASPhighlight>Sorted By:</SPAN> " '& Replace(sSort, " ORDER BY ", "") & ", "
		If InStr(sSort, "execID") > 0 Then sReturn = sReturn & "Exec ID "
		If InStr(sSort, "fname") > 0 Then sReturn = sReturn & "First Name "
		If InStr(sSort, "lname") > 0 Then sReturn = sReturn & "Last Name "
		If InStr(sSort, "email") > 0 Then sReturn = sReturn & "Email "
		If InStr(sSort, "position") > 0 Then sReturn = sReturn & "Position "
		If InStr(sSort, "serviceDates") > 0 Then sReturn = sReturn & "Service Dates "
		If InStr(sSort, "active") > 0 Then sReturn = sReturn & "Active "
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
