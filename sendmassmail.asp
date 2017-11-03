<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LB-USERTYPE")) <> 1 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
%>
<html>
	<head>
		<title>LanguageBank - Mass Email - Sending...</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SendMails()
		{
			document.frmMemail.action = "massmail.asp?send=1"
			document.frmMemail.submit();
		}
		-->
		</script>
	</head>
	<body BACKGROUND='images/myMail.gif'>
		<form method='post' name='frmSemail'>
		<table width='100%' border='0'>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td align='center'>This might take a few minutes to complete. Please wait...</td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td align='center'>
					<img src='images/myMail.gif' border='0' alt='Sending Email' title='Sending Email'>
				</td>
			</tr>
		</table>
		</form>
	</body>
</html>
