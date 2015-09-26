<!-- #INCLUDE file="includes/common/DB2ASP_environ.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->

<SCRIPT language=javascript>
	<!--
	function Validate() {
	var sMessage = '';
	var bFail = false;

	if (document.DB2ASPFileUpload.File1.value=='') {
		bFail = true;
		sMessage = sMessage + 'Please select an image to upload. \n';
		document.DB2ASPFileUpload.File1.focus()
	}
	if (bFail) {
		window.alert('The following errors were returned: \n\n' + sMessage);
		return false;
	}
	else {
		return true;
	}
	}
	//-->
</SCRIPT>

<!--<FORM ENCTYPE="multipart/form-data" ACTION="includes\common\DB2ASP_filesaver.asp" METHOD=POST></FORM>-->
<FORM NAME=DB2ASPFileUpload ENCTYPE="multipart/form-data" ACTION="includes\common\DB2ASP_filesaver.asp?<%=Request.QueryString()%>" METHOD=POST>
<TABLE CELLSPACING=1>
	<TR>
		<TD COLSPAN=2 CLASS=db2asplight><span class=db2asptitle>DB2ASP File Uploader</span></TD>
	</TR>
	<TR><TD><INPUT NAME=File1 SIZE=35 TYPE=file><BR></TD></TR>
	<!--<TR><TD><INPUT NAME=File2 SIZE=35 TYPE=file><BR></TD></TR>
	<TR><TD><INPUT NAME=File3 SIZE=35 TYPE=file><BR></TD></TR>-->
	<!-- <TR><TD CLASS=db2aspdark>Save File As:</TD><TD CLASS=db2asplight><INPUT TYPE=TEXT Name=SaveAs SIZE=25></TD></TR> -->
	<TR><TD>The upload time may vary depending on file size.<BR><BR></TD></TR>
</TABLE>
To delete image, leave "File Name" blank and click Upload.<BR>
<BR>
<INPUT TYPE=Submit NAME=btnUpload VALUE=Upload>&nbsp;&nbsp;<INPUT TYPE=Button NAME=btnCancel VALUE=Cancel onClick="javascript: location.href='<%=Request("pagename") & "?" & Request.QueryString%>'">
</FORM>
<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
