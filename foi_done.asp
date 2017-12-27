<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%


fetchid = Request("fetchid")

Set rsUser = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT [userid], [empname], [addr], [cellno], [email], ir.[fname], ir.[mname], ir.[lname], [suffix], [chksig], [ip], [useragent], [last] " & _
		"FROM [InfoRelease] AS ir " & _
		"WHERE ir.[userid]='" & fetchid & "'"

rsUser.Open strSQL, g_strCONN, 3, 1
' Response.Write strSQL
If rsUser.EOF Then
	Response.Write "<h1>Consent Form Not Found</h1><p>That consent form is not in the database</p>"
	Response.End
Else
	blnDisp = True
	strEmpName = Z_FixNull(rsUser("empname"))
	strAddr = Z_FixNull(rsUser("addr"))
	strCellno = Z_FixNull(rsUser("cellno"))
	strEmail = Z_FixNull(rsUser("email"))
	lngUserID = Z_CLng(rsUser("userid"))
	fname	= Z_FixNull(rsUser("fname"))
	mname	= Z_FixNull(rsUser("mname"))
	lname	= Z_FixNull(rsUser("lname"))
	suffix	= Z_FixNull(rsUser("suffix"))
	strIP	= Z_FixNull(rsUser("ip"))
	strUA	= Z_FixNull(rsUser("useragent"))
	strLast = FormatDateTime(rsUser("last"), 0)
	If rsUser("chksig") = 1 Then chksig = "checked"
End If
rsUser.Close
Set rsUser = Nothing
%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<title>Authorization for Release of Information</title>
	<meta name="description" content="Authorization for Release of Information">
	<meta name="author" content="Argao.net">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.css"
			integrity="sha256-sxzrkOPuwljiVGWbxViPJ14ZewXLJHFJDn0bv+5hsDY=" crossorigin="anonymous" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.css"
			integrity="sha256-ECB9bbROLGm8wOoEbHcHRxlHgzGqYpDtNTgDTyDz0wg=" crossorigin="anonymous" />
	<style>
.no_form td {
	color: red;
}
.gotform td {
	font-weight: bold;
	font-size: 120%;
}
.align-right {
	text-align: right;
}
.align-center {
	text-align: center;
}
	</style>
</head>
<body>
<div class="container">
	<div class="row" style="text-align: center;">
		<img src='images/LBISLOGO.jpg' border="0" style="width: 287px; height: 64px;" />
		<div class="twelve columns"><h1>RELEASE  OF  INFORMATION</h1></div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns align-right">
		<a href="admin_foi.asp" style="font-style: italic;">&lt;&lt;&nbsp;back</a>
		</div></div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="six columns">
			<label for="empName">Name of Employee:</label>
			<input class="u-full-width" type="text" placeholder="Employee Name" id="empname" name="empname"
					value="<%= strEmpName%>"
					/>
			<input type="hidden" name="userid" id="userid" value="<%= lngUserID %>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
			<label for="addr">Employee Address:</label>
			<input class="u-full-width" type="text" placeholder="Employee Address" id="addr" name="addr" 
					value="<%= strAddr%>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="four columns">
			<label for="addr">Cell Phone Number:</label>
			<input class="u-full-width" type="tel" maxlength="20" placeholder="Cell#" id="cellno" name="cellno"
					value="<%= strCellno%>" />
		</div>
		<div class="six columns">
			<label for="addr">e-Mail Address:</label>
			<input class="u-full-width" placeholder="e-mail address" type="email" id="email" name="email"
					value="<%= strEmail%>" />
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
			<p>I authorize Language Bank, a member of Ascentria Care Alliance, to share my personal
			information described below with Language Bank customers upon reasonable request of
			a customer for the purpose of compliance audits, credentialing monitoring, or public
			health concerns.  Depending on the nature of the request, Language Bank may disclose
			the following information responsive to that request:</p>
<ul style="margin-left: 10%;">
	<li>resume</li>
	<li>educational background</li>
	<li>credentialing certifications for interpretation</li>
	<li>continuing education certifications</li>
	<li>performance evaluations and competency check list</li>
	<li>immunization records</li>
	<li>contact information (including mobile phone number and e-mail address)</li>
</ul>
			<p>I understand that Language Bank will make all reasonable efforts to maintain my privacy
			in connection with the disclosure of this information, and that customers will be advised
			not to contact me directly outside of the time of scheduled appointments or outside the
			scope of standard business relations.</p>

			<p>My signature below releases Language Bank and Ascentria Care Alliance and their
			authorized agents from all liability in connection with the disclosure of this information.
			I understand that I may revoke my consent by sending written notice of revocation to the
			Language Bank Program Manager or Assistant Manager.  Such revocation will apply only to
			the release of information regarding appointments which have yet to occur as of the date
			of receipt of the revocation.  My revocation will not affect the release of information
			about appointments that have already occurred.</p>
			<div id="sigplace">
				(signed electronically)<br />
				on <strong><%= strLast %></strong>
				from [<strong><%= strIP %></strong>]
				<br />
				user agent: <strong><%=strUA%></strong>
			</div>
		</div>
	</div>
	<div class="row" style="margin-top: 50px;">
		<div class="one column">&nbsp;</div>
		<div class="ten columns" style="border-top: 1px dotted #999;">
			&nbsp;
		</div>
	</div>
</div>
</body>
</html>