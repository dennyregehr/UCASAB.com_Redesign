<%
' ***********************************
' Coded By: DB2ASP v4.3.5 (12) - A 303 Media Company
' For More Information:
' Please Visit http://www.DB2ASP.com
' Or Email support@303media.com
' ***********************************
' Page Name: Calendar_search.asp
' Date: 3/4/2005 12:28:23 AM
' Purpose: Search Records In Stored Procedure Results "Calendar"
' Database: Access
' Table: Calendar
' ***********************************

%>
<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<%
sObjectName = "Calendar"
Dim vEventName
Dim vEventTypeID, vEventTypeID2
Dim vLocation
Dim vStartDate, vStartDate2
Dim vEndDate, vEndDate2
Dim vEventDescription
Dim vCnd1, vCnd2, vCnd3, vCnd4, vCnd5, vCnd6

If Request("view") <> "newsearch" Then
	GetLastSearchValues()
End If
%>

<HTML>
<HEAD>
	<META name="GENERATOR" content="303 Media's DB2ASP v4.3.5"/>
	<TITLE>Calendar - Search</title>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=frmSearch action="Calendar_results.asp?view=newsearch<%=GetQueryString("NewSearch", Empty, Empty, Empty, Empty)%>" method=POST>
<TABLE cellspacing=1 class=DB2ASP>
	<TR>
		<TD colspan=3 class=DB2ASPlight>
			<TABLE cellspacing=1 class=menu width=100%><TR>
			<TD class=menu>
				<SPAN class=DB2ASPtitle>Calendar - Search</SPAN>
			</TD>
			<TD align=right class=menu width=120>
				<A href="Calendar_results.asp<%=GetQueryString("Search", Empty, Empty, Empty, Empty)%>" title="Last Results View of Records" class=DB2ASPvalue>Last Results View</A><br>
				<A href="Calendar_results.asp?view=listall<%=GetQueryString("ListAll", Empty, Empty, Empty, Empty)%>" title="List All Records" class=DB2ASPvalue>List All</A><br>
			</TD></TR></TABLE>
		</TD>
	</TR>
<% If Request("msg") <> "" Then %>
<TR><TD colspan=55 class=dark><SPAN class=message><%=Request("msg")%></SPAN></TD></TR>
<% End If %>
<TR>
	<TH class=DB2ASPdark align=right>EventName</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="EventName" SIZE=25 MAXLENGTH=40 value="<%=vEventName%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(1, vCnd1)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>EventTypeID</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<%=LookupEventTypeeventTypeID(vEventTypeID, "", 0, Empty, "EventTypeID")%>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(2, vCnd2)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>Location</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="Location" SIZE=25 MAXLENGTH=40 value="<%=vLocation%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(3, vCnd3)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>StartDate</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		Between<INPUT type=TEXT name="StartDate" SIZE=7 MAXLENGTH=10 value="<%=vStartDate%>" class=DB2ASPdetailinput>&nbsp;&nbsp;And <INPUT type=TEXT name="StartDate2" SIZE=7 MAXLENGTH=10 value="<%=vStartDate2%>" class=DB2ASPdetailinput>	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(4, vCnd4)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>EndDate</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		Between<INPUT type=TEXT name="EndDate" SIZE=7 MAXLENGTH=10 value="<%=vEndDate%>" class=DB2ASPdetailinput>&nbsp;&nbsp;And <INPUT type=TEXT name="EndDate2" SIZE=7 MAXLENGTH=10 value="<%=vEndDate2%>" class=DB2ASPdetailinput>	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(5, vCnd5)%></TD>
</TR>
<TR>
	<TH class=DB2ASPdark align=right>EventDescription</TH>
	<TD class=DB2ASPlight class=DB2ASPdetailcellbg>
		<INPUT type=text name="EventDescription" SIZE=25 MAXLENGTH=40 value="<%=vEventDescription%>" class=DB2ASPdetailinput>
	</TD>
	<TD class=DB2ASPlight><%=ConditionRadio(6, vCnd6)%></TD>
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
	If Session("vEventName") <> "" Then vEventName = Session("vEventName")
	If Session("vEventTypeID") <> "" Then vEventTypeID = Session("vEventTypeID")
	If Session("vEventTypeID2") <> "" Then vEventTypeID2 = Session("vEventTypeID2")
	If Session("vLocation") <> "" Then vLocation = Session("vLocation")
	If Session("vStartDate") <> "" Then vStartDate = Session("vStartDate")
	If Session("vStartDate2") <> "" Then vStartDate2 = Session("vStartDate2")
	If Session("vEndDate") <> "" Then vEndDate = Session("vEndDate")
	If Session("vEndDate2") <> "" Then vEndDate2 = Session("vEndDate2")
	If Session("vEventDescription") <> "" Then vEventDescription = Session("vEventDescription")
	If Session("vCnd1") <> "" Then vCnd1 = Session("vCnd1")
	If Session("vCnd2") <> "" Then vCnd2 = Session("vCnd2")
	If Session("vCnd3") <> "" Then vCnd3 = Session("vCnd3")
	If Session("vCnd4") <> "" Then vCnd4 = Session("vCnd4")
	If Session("vCnd5") <> "" Then vCnd5 = Session("vCnd5")
	If Session("vCnd6") <> "" Then vCnd6 = Session("vCnd6")
End Sub
%>
