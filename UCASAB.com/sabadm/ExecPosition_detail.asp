<%
' ***********************************
' Coded By: DB2ASP v4.3.5 (12) - A 303 Media Company
' For More Information:
' Please Visit http://www.DB2ASP.com
' Or Email support@303media.com
' ***********************************
' Page Name: ExecPosition_detail.asp
' Date: 3/4/2005 12:28:33 AM
' Purpose: Show detail information for one record. View, edit, add, and/or delete record.
' Database: Access
' Table: ExecPosition
' ***********************************

%>
<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<%
Dim vexecPositionID
Dim vPosition
Dim vRelevanceOrder
Dim vOrderForExecBoardPage
Dim vActive

sObjectName = "ExecPosition"
iPosition = Request("ab")

' If user submits a request choose the appropriate action
' All actions if sucessful, will redirect user back to results page
If (Request("btnSubmit") = "Add" Or Request("DB2ASPAction") = "Add") And sErr = "" Then
	If Not DetailAdd Then
		sErr = "Error Adding Record.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
	Else
		sMsg = "Record Added.<BR>"
	End If
ElseIf Request("btnSubmit") = "Delete" Or Request("fnc") = "delete" And sErr = "" Then
	If Not DetailDelete Then
		sErr = "Error Deleting Record.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
	Else
		sMsg = "Record Deleted.<BR>"
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg) & "&err=" & Server.URLEncode(sErr)
	End If
ElseIf  Request("btnDelete") = "Delete Checked Records" Then
	iRecordsDeleted = ResultsDelete
	If iRecordsDeleted = 0 Then
		sMsg = "No Records Deleted. " & sErr
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg)
	ElseIf iRecordsDeleted = -1 Then
		sMsg = "No Records Deleted, " & sErr
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg)
	Else
		sMsg = iRecordsDeleted & " Record/s Deleted.<BR>"
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg)
	End If
ElseIf (Request("btnSubmit") = "Update" Or (Request("DB2ASPAction") = "Update") And Request("btnMove") = "") And sErr = "" Then
	If Not DetailUpdate Then
		sErr = "Error Updating Record.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
		If Err.Number = -2147467259 Then
			sErr = Err.Description & " <BR>(You can fix this by changing your field properties in your database table design window.<BR>This can be done for all text/memo fields that can be blank/empty.)"
		End If
	Else
		sMsg = "Record Updated.<BR>"
	End If
ElseIf (Request("btnAdd") = "Add" And sErr = "") Then
	If Not ResultsAdd Then
		sErr = "Error Adding Record.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
	Else
		sMsg = "Record Added.<BR>"
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg)
	End If
ElseIf (Request("btnUpdate") = "Update Records" Or Request("DB2ASPAction") = "Update Records") And sErr = "" Then
	iRecordsUpdated = ResultsUpdate
	If iRecordsUpdated = 0 Then
		sErr = "Error Updating Records.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
		If Err.Number = -2147467259 Then
			sErr = Err.Description & " <BR>(You can fix this by changing your field properties in your database table design window.<BR>This can be done for all text/memo fields that can be blank/empty.)"
		End If
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg) & "&err=" & Server.URLEncode(sErr)
	Else
		sMsg = iRecordsUpdated & " Records Updated.<BR>"
		Response.Clear: Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&msg=" & Server.URLEncode(sMsg)
	End If
End If
' Main Detail Page Query - Retrieves records from table.
If Request("fnc") <> "add" And Request("DB2ASPAction") <> "Add" Then
	iMaxRecords = 5000

	iPageSize = GetLastPageSize(sObjectName)
	If iPageSize < 5 Then
		iPageSize = 25
	End If
	
	iPosition = Request("ab")
	sSort = GetLastSort(sObjectName)
	sWhere = GetLastSearch(sObjectName)

	If iPosition = "-1" Or iPosition = "" Then
		sWhere = " WHERE " & _
		"[execPositionID]=" & Replace(Request("PK_iexecPositionID"), "'", "''") & " "
	End If

		sSQL = _
		"SELECT TOP " & iMaxRecords & " * " & _
		"FROM [ExecPosition] " & " " & _
		sWhere & _
		sSort

	If sSQL <> "" Then
		' This is a security precaution
		If VerifySQL(sSQL) = "" Then
			Response.Redirect "ExecPosition_results.asp?view=listall"
		End If
	End If

	'Response.Write sSQL & "<BR>": Response.Flush
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.CursorLocation = adUseClient
	oRS.Open sSQL, oConn, adOpenForwardOnly, adLockOptimistic, &H1

	If iPosition = "-1" Or iPosition = "" Then
		iPosition = oRS.AbsolutePosition
	Else
		oRS.AbsolutePosition = CInt(iPosition)
		iPosition = oRS.AbsolutePosition
	End If

	If Len(Request("btnMove")) > 0 Then
		sAction = Request("btnMove")
		iPosition = CInt(iPosition)

		Select Case sAction
			Case "Previous"
				iPosition = CInt(iPosition) - 1
				oRS.MovePrevious
				If oRS.BOF Then
					oRS.MoveFirst
					sMsg = "Can't move beyond the first record."
					iPosition = CInt(iPosition) + 1
				End If
			Case "Next"
				iPosition = CInt(iPosition) + 1
				oRS.MoveNext
				If oRS.EOF Then
					oRS.MoveLast
					 sMsg = "Can't move beyond the last record."
					iPosition = CInt(iPosition) - 1
				End If
			Case "Last"
				oRS.MoveLast
				iPosition = oRS.AbsolutePosition
			Case "First"
				oRS.MoveFirst
				iPosition = oRS.AbsolutePosition
		End Select
	End If

		vexecPositionID = oRS("execPositionID")
		vPosition = oRS("Position")
		vRelevanceOrder = oRS("RelevanceOrder")
		vOrderForExecBoardPage = oRS("OrderForExecBoardPage")
		vActive = oRS("Active")

	If Err Then
		'Response.Redirect "ExecPosition_results.asp" & GetQueryString("Detail", Empty, Empty, Empty, Empty) & "&err=Error On Detail Page (" & Err.Description & " [" & Err.Number & "])"
		sErr = sErr & "Error Getting Records.<BR>"
		sErr = sErr & "[" & Err.Number & "] - " & Err.Description & ""
		Err.Clear
	End If
End If

%>
<HTML>
<HEAD>
	<META name="GENERATOR" content="303 Media's DB2ASP v4.3.5"/>
	<TITLE>Exec Position - Detail</TITLE>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=frmDetail method=post action="ExecPosition_detail.asp<%=GetQueryString("Detail", Empty, Empty, Empty, Empty)%>">
<TABLE cellspacing=1 class=DB2ASP>
	<TR>
		<TD colspan=5 class=DB2ASPlight>
			<TABLE cellspacing=1 class=menu width=100%><TR>
			<TD class=menu>
				<SPAN class=DB2ASPtitle>Exec Position</SPAN>
			</TD>
			<TD align=right class=menu width=120>
				<A href="ExecPosition_results.asp<%=GetQueryString("Detail", Empty, Empty, Empty, Empty)%>" title="Last Results View of Records" class=DB2ASPvalue>Last Results View</A><br>
				<A href="ExecPosition_results.asp?view=listall<%=GetQueryString("ListAll", Empty, Empty, Empty, Empty)%>" title="List All Records" class=DB2ASPvalue>List All</A><br>
			</TD></TR></TABLE>
		</TD>
	</TR>
<TR>
	<TD align=left colspan=2 class=DB2ASPlight>
		<%
		' Check For Errors or Messages and Display Them to User
		If sErr <> "" Then 
			Response.Write "<SPAN class=DB2ASPerror>" & sErr & "</SPAN><BR>"
		End If

		If sMsg <> "" Then 
			Response.Write "<SPAN class=DB2ASPmessage>" & sMsg & "</SPAN><BR>"
		End If
		%>
	</TD>
</TR>
<TR>
	<TH align=right class=DB2ASPdark>execPositionID</TH>
	<TD class=DB2ASPdetailcellbg>
		<%=vexecPositionID%>&nbsp;
	</TD>
</TR>
<TR>
	<TH align=right class=DB2ASPdark>Position</TH>
	<TD class=DB2ASPdetailcellbg>
		<INPUT type=text name="Position" SIZE=60 MAXLENGTH=100 value="<%=vPosition%>" class=DB2ASPdetailinput>
	</TD>
</TR>
<TR>
	<TH align=right class=DB2ASPdark>RelevanceOrder</TH>
	<TD class=DB2ASPdetailcellbg>
		<INPUT type=text name="RelevanceOrder" SIZE=4 MAXLENGTH=2 value="<%=vRelevanceOrder%>" class=DB2ASPdetailinput>
	</TD>
</TR>
<TR>
	<TH align=right class=DB2ASPdark>OrderForExecBoardPage</TH>
	<TD class=DB2ASPdetailcellbg>
		<INPUT type=text name="OrderForExecBoardPage" SIZE=4 MAXLENGTH=2 value="<%=vOrderForExecBoardPage%>" class=DB2ASPdetailinput>
	</TD>
</TR>
<TR>
	<TH align=right class=DB2ASPdark>Active</TH>
	<TD class=DB2ASPdetailcellbg>
		<INPUT type=radio name="Active" value=1 <%=RadioSelection(vActive, 1)%>>True&nbsp;&nbsp;
		<INPUT type=radio name="Active" value=0 <%=RadioSelection(vActive, 0)%>>False
	</TD>
</TR>
	<% If Request("fnc") <> "add" And iPosition <> "-1" Then %>
	<TR><TD align=right colspan=55 class=DB2ASPdark>
		<BR>
		<INPUT type=Submit name=btnMove value="First" class=DB2ASPactionbtn>&nbsp;&nbsp;<INPUT type=Submit name=btnMove value="Previous" class=DB2ASPactionbtn>&nbsp;&nbsp;<INPUT type=Submit name=btnMove value="Next" class=DB2ASPactionbtn>&nbsp;&nbsp;<INPUT type=Submit name=btnMove value="Last" class=DB2ASPactionbtn>
		<INPUT type=Hidden name=ab value="<%=iPosition%>">
	</TD>
	<% End If %>
</TABLE>
<BR><BR>
<INPUT type=hidden name="PK_iexecPositionID" value='<%=vexecPositionID%>'>
<% If Request("fnc")="add" Or Request("Db2ASPAction")="Add" Or (Request("btnSubmit") = "Add" And sErr <> "") Then %>
When the record is added, you will see a status msg (in green) and the<br>form will be automatically ready to add another record (blank).<br><br>
If this record is to have a database or file system image,<br>add record first then add image with the update link.<br><br>
<INPUT type=Submit name=btnSubmit value="Add" class=DB2ASPactionbtn>&nbsp;&nbsp;
<INPUT type=hidden name="DB2ASPAction" value="Add">
<INPUT type=hidden name="fnc" value="add">
<% Else %>
<INPUT type=Submit name=btnSubmit value="Update" class=DB2ASPactionbtn>&nbsp;&nbsp;
<INPUT type=hidden name="DB2ASPAction" value="Update">
<% End If %>

<% If Request("fnc") <> "add" Then %>
<INPUT type=Submit name=btnSubmit value="Delete" onClick="javascript: if(confirm('Are You Sure You Want To Delete This Record?')==false) return false; document.forms[0].btnSubmit.value = 'Delete'" class=DB2ASPactionbtn>&nbsp;&nbsp;
<% End If %>
<INPUT type=Submit name=btnBack value="Back To Results" onClick="javascript: document.location.href='ExecPosition_results.asp<%=GetQueryString("Detail", Empty, Empty, Empty, Empty)%>'; return false;" class=DB2ASPactionbtn>&nbsp;&nbsp;
<INPUT type=reset value="Reset" class=DB2ASPotherbtn>
</CENTER>
</FORM>
<!--END--42821-->
<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
<!-- #INCLUDE file="includes/pages/DB2ASP_ExecPosition_functions.asp" -->
