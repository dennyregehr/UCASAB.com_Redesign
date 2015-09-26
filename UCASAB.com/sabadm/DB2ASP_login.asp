<%
Option Explicit
Dim sErr
Dim sMsg
Dim sLogin
Dim sPassword

Const LOGIN = "kendradee"
Const PASSWORD = "jacobwade"

Session("Group") = 0

If Request.Cookies("DB2ASP_ADMIN")("r") = "1" Then
	sLogin = Request.Cookies("DB2ASP_ADMIN")("l")
	sPassword = Request.Cookies("DB2ASP_ADMIN")("p")
	Session("Group") = Request.Cookies("DB2ASP_ADMIN")("g")
End If

If sLogin = "" Or sPassword = "" Then
	sLogin = Request.Form("Login")
	sPassword = Request.Form("Password")
End If

If sLogin <> "" And sPassword <> "" Then
	If (sLogin = LOGIN And sPassword = PASSWORD) Then
		If Request.Form("Remember") = "1" Then
			Response.Cookies("DB2ASP_ADMIN")("l") = Request("Login")
			Response.Cookies("DB2ASP_ADMIN")("p") = Request("Password")
			Response.Cookies("DB2ASP_ADMIN")("r") = "1"
			Response.Cookies("DB2ASP_ADMIN")("g") = "2"
			Response.Cookies("DB2ASP_ADMIN").Expires = "January, 18 2031"
		End If
		Session("Group") = 2
		Response.Redirect "DB2ASP_menu.asp"
	End If
End If
%>
<HTML>
<HEAD>
	<TITLE>Login</TITLE>
<!-- #INCLUDE file="includes/common/DB2ASP_header.asp" -->
<FORM name=Login method=POST>
<%=sMsg%><BR><BR>
<CENTER>
<TABLE cellspacing=0>
<TR>
	<TH align=center class=dark>Please Login<BR><BR></TH>
</TR>
<TR>
	<td>
		<TABLE cellspacing=0 width=350>
			<TR>
				<TD align=right class=dark><B>Login:&nbsp;&nbsp;</TD>
				<TD class=light><input name=Login type=text value="<%=sLogin%>"></TD>
			</TR>
			<TR>
				<TD align=right class=dark><B>Password:&nbsp;&nbsp;</TD>
				<TD class=light><input name=Password type=password></TD>
			</TR>
			<TR>
				<TD align=right class=dark>&nbsp;</TD>
				<TD class=light><input name=Remember type=checkbox value=1> Remember my login and password on this computer</TD>
			</TR>
		</table>
		<center>
		<BR>
		<input name="btnLogin" type=Submit value="Login">
		<script language=javascript>
		<!--
		document.Login.Login.focus();
		// -->
		</script>
	</TD>
</TR>
</TABLE>
</FORM>

<!-- #INCLUDE file="includes/common/DB2ASP_footer.asp" -->
<!-- #INCLUDE file="includes/common/DB2ASP_functions.asp" -->
