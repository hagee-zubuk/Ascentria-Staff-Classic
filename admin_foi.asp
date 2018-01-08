<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
tmpPage = "document.frmReport."
%>
<!doctype html>
<html lang="en">
<%language=vbscript%>
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
	</div>
	<div class="row align-center">
		<div class="two columns">&nbsp;</div>
		<div class="eight columns"><strong style="font-size: 150%;">RELEASE  OF  INFORMATION FORMS</strong></div>
		<div class="two columns"><a class="u-full-width button button-secondary" href="dl_csv.asp?fn=<%=strFN%>">Get CSV</a></div>
	</div>
	<form id="frmROI" name="frmROI" method="post" action="admin_foi_proc.asp">
	<div class="row">
		<div class="twelve columns align-center">
	<%=Session("MSG")%>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns">
<table class="u-full-width" style="font-size: 80%;"> 			
	<thead>
		<tr>
			<th>Name</th>
			<th>Date</th>
			<th>Form</th>
		</tr>
	</thead>
	<tbody>
<%
Set FSO = CreateObject("Scripting.FileSystemObject")

tmpDate = Replace(FormatDateTime(Now, 0), ":", "")
tmpDate = Replace(tmpDate, "/", "")
tmpDate = Replace(tmpDate, " ", "-")
RepCSV =  "ConsentDocs." & tmpdate & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
strFN = Z_DoEncrypt( RepCSV )
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM [interpreter_T] AS i" & _
		"WHERE Active = 1 ORDER BY [Last Name], [First Name]"
strSQL = "SELECT i.[index], COALESCE(ir.[userid],-1) AS [userid], ir.[empname], ir.[addr], ir.[cellno]" & _
		", i.[Last Name], i.[First Name]" & _
		", ir.[email], ir.[fname], ir.[mname], ir.[lname], ir.[suffix]" & _
		", ir.[chksig], ir.[ip], ir.[useragent], ir.[last] " & _
		"FROM [interpreter_T] AS i " & _ 
		"INNER JOIN [user_t] AS u ON i.[index]=u.[intrid] " & _
		"LEFT JOIN [InfoRelease] AS ir ON u.[index]=ir.[userid] " & _
		"ORDER BY i.[Last Name] ASC"
rsIntr.Open strSQL, g_strCONN, 3, 1
Prt.WriteLine "LANGUAGE BANK - CONSENT FORM SUBMISSIONS"
CSVHead = """Last Name"",""First Name"",""Date"",""e-Mail"",""Last Name"",""Middle Name"",""First Name""" & _
		",""Suffix"",""Address"",""Cell#"",""IP"",""Useragent"""
CSVBody = ""
Do Until rsIntr.EOF
	If CLng(rsIntr("userid")) > 0 Then
		Response.Write "<tr class=""gotform"">"
	Else
		Response.Write "<tr class=""no_form"">"
	End If
	Response.Write "<td>" & rsIntr("Last Name") & ", " & rsIntr("First Name") 
	CSVBody = CSVBody & """" & rsIntr("Last Name") & """,""" & rsIntr("First Name") & ""","
	If CLng(rsIntr("userid")) > 0 Then
		Response.Write "</td><td>" & FormatDateTime(rsIntr("last"), 0)
		Response.Write "</td><td><a href=""foi_done.asp?fetchid=" & rsIntr("userid") & """>view form</a>"
		CSVBody = CSVBody & """" & FormatDateTime(rsIntr("last"), 0) & """,""" & _
				rsIntr("email") & """,""" & rsIntr("lname") & """,""" & rsIntr("mname") & """,""" & rsIntr("fname")
		CSVBody = CSVBody & """,""" & rsIntr("suffix") & """,""" & rsIntr("addr") & """,""" & rsIntr("cellno")
		CSVBody = CSVBody & """,""" & rsIntr("ip") & """,""" & rsIntr("useragent") & """"
	Else
		CSVBody = CSVBody & ",,,,,,,,,"
		Response.Write "</td><td>&nbsp;--"
		Response.Write "</td><td>&nbsp;"
	End If
	Response.Write "</td></tr>" & vbCrLf
	CSVBody = CSVBody & vbCrLf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsInt = Nothing
Prt.WriteLine strMSG
Prt.WriteLine CSVHead
Prt.WriteLine CSVBody
Prt.Close	
Set Prt = Nothing

%>
</table>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="ten columns"><a class="button button-secondary" href="dl_csv.asp?fn=<%=strFN%>">download CSV</a></div>
	</div>
</form></div>
</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>