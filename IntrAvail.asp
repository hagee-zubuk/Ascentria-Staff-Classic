<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
'lists available interpreter for given date, time and dialect
tmpAppInfo = Split(Request("AppInfo"), "|")
tmpAppDate = tmpAppInfo(0)
tmpAppTime = tmpAppInfo(1)
tmpAppLang = tmpAppInfo(2)

Set rsLangIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM Language_T WHERE [index] = " & tmpAppLang
rsLangIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsLangIntr.EOF Then
	IntrLang = UCase(Trim(rsLangIntr("Language")))
End If
rsLangIntr.Close
Set rsLangIntr = Nothing
If IntrLang <> "" Then
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntrLang = "SELECT * FROM interpreter_T WHERE Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & "' OR Upper(Language3) = '" & IntrLang & _
		"' OR Upper(Language4) = '" & IntrLang & "' OR Upper(Language5) = '" & IntrLang & "' And Active = 1 ORDER BY [Last Name], [First Name]" 
	rsIntr.Open sqlIntrLang, g_strCONN, 3, 1
	Do Until rsIntr.EOF
		Set rsAvail = Server.CreateObject("ADODB.RecordSet")
		sqlAvail = "SELECT * FROM Request_T WHERE appDate = '" & tmpAppDate & "' AND appTimeFrom = '" & tmpAppTime & "' AND IntrID = " & rsIntr("index")
		rsAvail.Open sqlAvail, g_strCONN, 3, 1

		If rsAvail.EOF Then
			If IsNull(rsIntr("vacfrom")) Then
				strAvail = strAvail & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
					"•&nbsp;&nbsp;" & rsIntr("Last Name") & ", " &  rsIntr("First Name") & "<br>"
			Else
				If Not(tmpAppDate >= rsIntr("vacfrom") And tmpAppDate <= rsIntr("vacto")) Then
					strAvail = strAvail & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
						"•&nbsp;&nbsp;" & rsIntr("Last Name") & ", " &  rsIntr("First Name") & "<br>"
				End If
			End If
		End If
		rsAvail.Close
		Set rsAvail = Nothing
		rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Interpreter Availability </title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmZip' action=''>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					
					<td class='header' colspan='6' align='center'>
						<nobr>Available Intepreters
					</td>
				</tr>
				<tr>
					<td colspan='6' align='center'>
						<nobr>&nbsp;<%=tmpAppDate%>&nbsp;&nbsp;•&nbsp;&nbsp;<%=tmpAppTime%>&nbsp;&nbsp;•&nbsp;&nbsp;<%=IntrLang%>&nbsp;
					</td>
				</tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
				<tr>
					<td colspan='6' align='left'>
						<font size='2'><b><%=strAvail%></b></font>
					</td>
				</tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
			</table>
		</form>
	</body>
</html>