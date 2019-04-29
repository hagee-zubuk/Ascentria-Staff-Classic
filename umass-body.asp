<%language=vbscript%>
<!DOCTYPE html>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<html lang="en">
	<head>
		<meta charset="utf-8">
		<title>Language Bank - UMass Instructions</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<style>
td:first-child {
	text-align: right;
	vertical-align: top;
	padding-top: 2px;
}
		</style>
	</head>
	<body >
<%
	strSQL = "SELECT r.[index], r.[InstID]" & _
			", r.[langID], l.[Language]" & _
			", r.[appDate], r.[appTimeFrom], r.[appTimeTo]" & _
			", r.[deptID],  d.[dept], d.[Address], d.[InstAdrI], d.[City], d.[State], d.[Zip]" & _
			", r.[CliAdd], r.[Caddress], r.[Ccity], r.[cstate], r.[Czip], r.[Cphone]" & _
			", r.[clname], r.[cfname], r.[courtcall]" & _
			", r.[mrrec], r.[LBcomment] " & _
			"FROM [request_t] AS r " & _
			"INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
			"INNER JOIN [language_T] AS l ON r.[langID]=l.[index] " & _
			"WHERE r.[InstID]=860 AND " & _
			"r.[index]=" & Request("reqid")
	'Response.Write "<code>" & strSQL & "</code>"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open strSQL, g_strCONN, 3, 1
	If rsIntr.EOF Then
		Response.Write "<h1>INVALID</h1><p>The request ID is invalid, or something went wrong. Please contact LB Tech Support.</p></body></html>"
		Response.End
	End If
	If (rsIntr("cliAdd") ) Then
		'tmpAddress = Trim(rsIntr("cfname") & " " & rsIntr("clname"))
		'If Len(tmpAddress) > 1 Then tmpAddress = tmpAddress & " --<br />"
		tmpAddress = "(Client Address)<br />" & rsIntr("caddress")
		tmpStreet = rsIntr("ccity") & ", " & rsIntr("cstate") & " " & rsIntr("czip")
	Else
		tmpAddress = rsIntr("dept") & "<br />" & Trim(rsIntr("InstAdrI"))
		If (Len(tmpAddress) > 1) Then
			tmpAddress = rsIntr("address") & "<br />" & tmpAddress
		Else
			tmpAddress = rsIntr("address")
		End If
		tmpStreet = rsIntr("city") & ", " & rsIntr("state") & " " & rsIntr("zip")
	End If
	'strComment = Trim(rsIntr("LBcomment"))
	'If Len(strComment) > 1 Then
	'	strComment = "<tr><td>Comment:&nbsp;</td><td><pre>" & strComment & "</pre></td></tr>"
	'End If
	If (rsIntr("courtcall") = TRUE) Then 
		strPhone = Trim(rsIntr("Cphone"))
		If Len(strPhone) > 1 Then
			strPhone = "<tr><td>Phone:&nbsp;</td><td>" & strPhone & "</td></tr>"
		End If
	End If

%>
<div style="width: 80%; margin: 30px auto;">
<h1>DO NOT PRINT THIS PAGE</h1>
<p>
Report to:
<address style="margin-left: 50px; font-weight: bold; line-height: 120%; font-size: 150%;">
	<%= tmpAddress%><br />
	<%= tmpStreet%><br />
</address>
</p>
<table style="width: 90%;">
	<thead></thead>
	<tbody>
	<tr>
		<td style="text-align: right;">Appointment Date/Time:&nbsp;</td>
		<td><u><b style="font-size: 150%;"><%=WeekdayName(DatePart("w", rsIntr("appDate")))%>,&nbsp;<%=rsIntr("appDate")%>
&nbsp;<i><%=Z_FormatTime(rsIntr("appTimeFrom"), 4) %>
 - <%=Z_FormatTime(rsIntr("appTimeTo"), 4) %></i></b></u></td></tr>
 	<tr><td>Patient Name:&nbsp;</td>
 		<td><b style="font-size: 150%;"><%= rsIntr("cfname") %> <%= rsIntr("clname") %></b></td></tr>
<%=strPhone%>
	<tr><td>Language:&nbsp;</td>
		<td><%= rsIntr("Language") %></td></tr>
	<tr><td>Medical Record #:&nbsp;</td>
		<td><b><%= rsIntr("mrrec") %></b></td></tr>
<%=strComment%>
	</tbody>
</table>
<h1>=&nbsp;=&nbsp;=&nbsp;&nbsp;DO NOT PRINT&nbsp;&nbsp;=&nbsp;=&nbsp;=</h1>
<p>You are expected to use the UMass Memorial Medical Center Encounter Form, which you should find attached (see below). You can also pick these up at the hospital.</p>
<p>WHEN YOU ARE ON-SITE: Call UMass Interpreter Services Dispatching at 508-334-7651 or 774-441-
6793 or, from any hospital phone, Extension 1-6793</p>
<p><a style="color: #222;" href="https://interpreter.thelanguagebank.org/interpreter/Instructions for Interpreters at UMass pdf version 10.10.17.pdf">
	Instructions for using the attached </a><a style="color: #222;" href="https://interpreter.thelanguagebank.org/interpreter/umass_encounter_form.2018.pdf">
	UMass Verification/Encounter Form</a>
	<a style="color: #222;" href="https://interpreter.thelanguagebank.org/interpreter/Instructions for Interpreters at UMass pdf version 10.10.17.pdf">
	</a>are found in <u>https://interpreter.thelanguagebank.org/interpreter/Instructions for Interpreters at UMass pdf version 10.10.17.pdf</u>.</p>
</div>
	</body>
</html>
<%
	rsIntr.Close
	Set rsIntr = Nothing
%>
