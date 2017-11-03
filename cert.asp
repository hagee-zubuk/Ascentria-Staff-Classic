<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'GET LANG
Set rsLangs = Server.CreateObject("ADODB.RecordSet")
sqlLangs = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsLangs.Open sqlLangs, g_strCONN, 1, 3
If Not rsLangs.EOF THen
	tmpLang = UCase(GetLang(rsLangs("LangID")))
	tmpAppDate = rsLangs("AppDate")
End If
rsLangs.Close
Set rsLangs = Nothing
'GET INTERPRETERS
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE (Upper(Language1) = '" & tmpLang & "' OR Upper(Language2) = '" & tmpLang & "' OR Upper(Language3) = '" & tmpLang & _
		"' OR Upper(Language4) = '" & tmpLang & "' OR Upper(Language5) = '" & tmpLang & "') AND Active = 1 ORDER BY [Last Name], [First Name]" 
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	If isNull(rsIntr("vacfrom")) Then
		'GET CERT
		Set rsCert = Server.CreateObject("ADODB.RecordSet")
		sqlCert = "SELECT * FROM IntrTraining_T WHERE Type = 3 AND IntrID = " & rsIntr("index") & " ORDER BY date"
		rsCert.Open sqlCert, g_strCONN, 3, 1
		If Not rsCert.EOF then
			tmpCert = tmpCert & "<tr onclick='PassMe(" & rsIntr("index") & ");'><td align='center'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</td>" & _
				"<td align'center'><hr width='75%'></td></tr>"
			Do Until rsCert.EOF
				tmpCert = tmpCert & "<tr><td>&nbsp;</td><td align='center'>" & rsCert("cert") & "</td></tr>"
				rsCert.MoveNext
			Loop
		End If
		rsCert.Close
		Set rsCert = Nothing
	Else
		If Not (tmpAppDate >= rsIntr("vacfrom") And tmpAppDate <= rsIntr("vacto")) Then
		'GET CERT
		Set rsCert = Server.CreateObject("ADODB.RecordSet")
		sqlCert = "SELECT * FROM IntrTraining_T WHERE Type = 3 AND IntrID = " & rsIntr("index") & " ORDER BY date"
		rsCert.Open sqlCert, g_strCONN, 3, 1
		If Not rsCert.EOF then
			tmpCert = tmpCert & "<tr onclick='PassMe(" & rsIntr("index") & ");'><td align='center'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</td>" & _
				"<td align'center'><hr width='75%'></td></tr>"
			Do Until rsCert.EOF
				tmpCert = tmpCert & "<tr><td>&nbsp;</td><td align='center'>" & rsCert("cert") & "</td></tr>"
				rsCert.MoveNext
			Loop
		End If
		rsCert.Close
		Set rsCert = Nothing
		End If
	End If
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Certification</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function PassMe(xxx)
		{
			window.opener.document.frmAssign.selIntr.value = xxx;
			window.opener.document.frmAssign.selIntr.focus();
			self.close();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmHist' action=''>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					<td class='header' colspan='2'>
						<nobr>Certificate --&gt&gt
					</td>
				</tr>
				<tr>
					<td align='center'><u>Name</u></td><td align='center'><u>Certificate</u></td>
				</tr>
				<%=tmpCert %>
				
			</table>
		</form>
	</body>
</html>