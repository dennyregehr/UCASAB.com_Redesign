<%
' ***********************************
' Coded By: DB2ASP v4.3.5 (12) - A 303 Media Company
' For More Information:
' Please Visit http://www.DB2ASP.com
' Or Email support@303media.com
' ***********************************
' Page Name: ExecBoard_search.asp
' Date: 3/4/2005 12:28:29 AM
' Purpose: Search Records In Stored Procedure Results "ExecBoard"
' Database: Access
' Table: ExecBoard
' ***********************************

%>
<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<%
sObjectName = "ExecBoard"
Dim vfname
Dim vlname
Dim vemail
Dim vposition, vposition2
Dim vCnd1, vCnd2, vCnd3, vCnd4

If Request("view") <> "newsearch" Then
	GetLastSearchValues()
End If
%>

<HTML>
<HEAD>
	<META name="GENERATOR" content="303 Media's DB2ASP v4.3.5"/>
	<TITLE>Exec Board - Search</title>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=frmSearch action="ExecBoard_results.asp?view=newsearch<%=GetQueryString("NewSearch", Empty, Empty, Empty, Empty)%>" method=POST>
<TABLE cellspacing=1 class=DB2ASP>
	<TR>
		<TD colspan=3 class=DB2ASPlight>
			<TABLE cellspacing=1 class=menu width=100%><TR>
			<TD class=menu>
				<SPAN class=DB2ASPtitle>Exec Board - Search</SPAN>
			</TD>
			<TD align=right class=menu width=120>
				<A href="ExecBoard_results.asp<%=GetQueryString("Search", Empty, Empty, Empty, Empty)%>" title="Last Results View of Records" class=DB2ASPvalue>Last Results View</A><br>
				<A href="ExecBoard_results.asp?view=listall<%=GetQueryString("ListAll", Empty, Empty, Empty, Empty)%>" title="List All Records" class=DB2ASPvalue>List All</A><br>
			</TD></TR></TABLE>
		</TD>
	</TR>
<% If Request("msg") <> "" Then %>
<TR><TD colspan=55 class=dark><SPAN class=message><%=Request("msg")%></SPAN></TD></TR>
<% End If %>
<TR>
	<TH class=DB2ASPdark align=right>First Name</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="fname" SIZE=25 MAXLENGTH=40 value="<%=vfname%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(1, vCnd1)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>Last Name</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="lname" SIZE=25 MAXLENGTH=40 value="<%=vlname%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(2, vCnd2)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>Email</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="email" SIZE=25 MAXLENGTH=40 value="<%=vemail%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(3, vCnd3)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>Position</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<%=LookupExecPositionexecPositionID(vposition, "", 0, Empty, "position")%>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(4, vCnd4)%></TD>
</TR>
</TR></TABLE>
<BR>
<INPUT type=Submit class=DB2ASPactionbtn name=params value="Submit Search">
<INPUT type=reset class=DB2ASPotherbtn  name=btnReset>
</FORM>
<!--END--42821-->
<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_lookups.asp" -->
<%
Sub GetLastSearchValues()
	' This procedure can be omitted, but user functionality will be reduced
	If Session("vfname") <> "" Then vfname = Session("vfname")
	If Session("vlname") <> "" Then vlname = Session("vlname")
	If Session("vemail") <> "" Then vemail = Session("vemail")
	If Session("vposition") <> "" Then vposition = Session("vposition")
	If Session("vposition2") <> "" Then vposition2 = Session("vposition2")
	If Session("vCnd1") <> "" Then vCnd1 = Session("vCnd1")
	If Session("vCnd2") <> "" Then vCnd2 = Session("vCnd2")
	If Session("vCnd3") <> "" Then vCnd3 = Session("vCnd3")
	If Session("vCnd4") <> "" Then vCnd4 = Session("vCnd4")
End Sub
%>
