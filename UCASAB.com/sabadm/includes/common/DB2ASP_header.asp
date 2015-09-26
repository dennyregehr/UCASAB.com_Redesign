	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function CheckAll() {
		for (var i=0;i<document.results.elements.length;i++) {
			var ele = document.results.elements[i];
			if ((ele.name != 'chkall') && (ele.type=='checkbox')) {
				ele.checked = document.results.chkall.checked;
			}
		}
	}
	//-->
	</SCRIPT>

	<META name="GENERATOR" content="303 Media - DB2ASP v4.3.5"/>
	<LINK rel="stylesheet" href="styles/DB2ASP_stylesheet.css">

</HEAD>
<BODY>

<TABLE border=0 cellspacing=0 cellpadding=0 width="100%" height="100%">
<TR>
	<TD colspan=2 width=150 height=57 class=DB2ASPTDSideMenu><A HREF="http://www.DB2ASP.com"><IMG SRC="images/DB2ASP/303.jpg" height=57 width=150 border=0 align=abstop></A></TD>
	<TD colspan=3 height=57>
		<TABLE border=0 cellspacing=0 cellpadding=0 width="100%" height=57>
		<TR class=DB2ASPTopBar1 height=30>
			<TD class=DB2ASPTopBar1 height=30>&nbsp;DB2ASP - Database Administration</TD>
		</TR>
		<TR class=DB2ASPTopBar2 height=27>
			<TD class=DB2ASPTopBar2 height=27>&nbsp;
				<A HREF="DB2ASP_login.asp?fnc=logout" class=DB2ASPATop>Log Out</A>
			</TD>
		</TR>
		</TABLE>
	</TD>
</TR>
<TR height=1>
	<TD colspan=5 height=1><IMG SRC="images/DB2ASP/b.gif" width="100%" height=1 border=0></TD>
</TR>
<TR>
	<TD width=4 valign=top class=DB2ASPTDSideMenu><IMG SRC="images/DB2ASP/s.gif" width=4 height=1 border=0></TD>
	<TD align=left valign=top width=146 class=DB2ASPTDSideMenu>
		<% If Session("Group") > 0 Then %>
		<!-- Begin Left Navigation Menu -->
		<BR><BR>
		<B>Calendar</B><BR>
		&nbsp;&nbsp;- <A href="Calendar_search.asp?view=lastsearch<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Search</A><BR>
		&nbsp;&nbsp;- <A href="Calendar_results.asp?view=lastresults<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Results</A>&nbsp;
		<A href="Calendar_results.asp?view=listall<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>List All</a><BR>
		<BR>
		<B>Event Type</B><BR>
		&nbsp;&nbsp;- <A href="EventType_search.asp?view=lastsearch<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Search</A><BR>
		&nbsp;&nbsp;- <A href="EventType_results.asp?view=lastresults<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Results</A>&nbsp;
		<A href="EventType_results.asp?view=listall<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>List All</a><BR>
		<BR>
		<B>Exec Board</B><BR>
		&nbsp;&nbsp;- <A href="ExecBoard_search.asp?view=lastsearch<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Search</A><BR>
		&nbsp;&nbsp;- <A href="ExecBoard_results.asp?view=lastresults<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Results</A>&nbsp;
		<A href="ExecBoard_results.asp?view=listall<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>List All</a><BR>
		<BR>
		<B>Exec Position</B><BR>
		&nbsp;&nbsp;- <A href="ExecPosition_search.asp?view=lastsearch<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Search</A><BR>
		&nbsp;&nbsp;- <A href="ExecPosition_results.asp?view=lastresults<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>Results</A>&nbsp;
		<A href="ExecPosition_results.asp?view=listall<%=GetQueryString("Nav", Empty, Empty, Empty, Empty)%>" class=DB2ASPAMenu>List All</a><BR>
		<BR>
		<BR><BR>
        <p>
            The website makes use of caching technology which makes it load faster and gives the user a better
            experience.  Caching means that the website doesn't update automatically when you save data in
            this administration tool.
            <br />
            <a href="/ClearAllCache.aspx">Click here to clear the cache and update the website immediately.</a>
        </p>
		<!-- End Left Navigation Menu -->
		<% Else %>

			<BR><BR><A href="DB2ASP_login.asp" class=DB2ASPAMenu>Please Login</a><BR>
		<% End If %>
	</TD>
	<TD valign=top width=1 height="100%" bgcolor=black><IMG SRC="images/DB2ASP/s.gif" width=1 height=1 border=0></TD>
	<TD valign=top width=10 class=DB2ASPmainbody>&nbsp;</TD>
	<TD valign=top align=left width=""100%" class=DB2ASPmainbody>&nbsp;
		<%
		' Check For Errors and Display Them to User
		If sErr <> "" Then 
			Response.Write "<BR><BR><B>" & sErr & "</B>"
			Response.End
		End If
		%>

