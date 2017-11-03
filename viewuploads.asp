<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	If Request("ftype") = 0 Then
		ftype = 0
		subfold = "\vform\"
		msgtype = "Verification form"
	ElseIf Request("ftype") = 1 Then
		ftype = 1
		subfold = "\tolls\"
		msgtype = "Tolls and Parking Receipts"
	End If
	viewpath = ""
	If Z_fixNull(request("fname")) <> "" Then 
		viewpath = uploadpath & Request("reqid") & subfold & request("fname")
		strMSG = "Viewing " & msgtype & " uploaded on " & Request("ts")
		
	End If
	'Set rsFile = Server.CreateObject("ADODB.RecordSet")
	'ftype = 0
	'If Request("ftype") = 1 Then ftype = 1
	'sqlFile = "SELECT * FROM uploads WHERE RID = " & Request("reqid") & " AND [type] = " & ftype & " ORDER BY [timestamp] DESC"
	'rsFile.Open sqlFile, g_strCONNupload, 3, 1
	'Do Until rsFile.EOF
	'	strFile = strFile & "• <a style='text-decoration: none;' href='viewuploads.asp?ftype=" & request("ftype") & "&reqid=" & Request("reqid") & "&fname=" & rsFile("filename") & "'>" & rsFile("timestamp") & "</a><br>"
	'	rsFile.MoveNext
	'Loop
	'rsFile.Close
	'Set rsFile = Nothing
	
	Set rsVforms = Server.CreateObject("ADODB.RecordSet")
	rsVforms.Open "SELECT * FROM uploads WHERE RID = " & Request("reqid") & " AND [type] = 0 ORDER BY [timestamp] DESC", g_strCONNupload, 3, 1
	Do Until rsVforms.EOF
		strVform = strVform & "&nbsp;&nbsp;&nbsp;&nbsp;<a style='text-decoration: none;' href='viewuploads.asp?ftype=0&reqid=" & Request("reqid") & "&fname=" & rsVforms("filename") & "&ts=" & rsVforms("timestamp")  & "'><img src='images/zoom.gif'>" & rsVforms("timestamp") & "</a><br>"
		rsVforms.MoveNext
	Loop
	rsVforms.Close
	Set rsVforms = Nothing
	
	Set rsTolls = Server.CreateObject("ADODB.RecordSet")
	rsTolls.Open "SELECT * FROM uploads WHERE RID = " & Request("reqid") & " AND [type] = 1 ORDER BY [timestamp] DESC", g_strCONNupload, 3, 1
	Do Until rsTolls.EOF
		strTolls = strTolls & "&nbsp;&nbsp;&nbsp;&nbsp;<a style='text-decoration: none;' href='viewuploads.asp?ftype=1&reqid=" & Request("reqid") & "&fname=" & rsTolls("filename") & "&ts=" & rsTolls("timestamp")  & "'><img src='images/zoom.gif'>" & rsTolls("timestamp") & "</a><br>"
		rsTolls.MoveNext
	Loop
	rsTolls.Close
	Set rsTolls = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - VIEW UPLOADS</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		
		//-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form name="frmUpload" method="POST">
			<table border=0 style="width:100%;">
				<tr>
					<td class='header' colspan='2'>
						<nobr>View Uploads --&gt;&gt;
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table>
							<tr>
								<td align="left"><b><u>Verification Forms</u></b></td>
							</tr>
							<tr>
								<td align="left"><nobr><%=strVform%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align="left"><b><u>Tolls and Parking Receipts</u></b></td>
							</tr>
							<tr>
								<td align="left"><nobr><%=strTolls%></td>
							</tr>
						</table>
					</td>
					<td>
						<table>
							<tr>
								<td>
									<%=strMSG%>
								</td>
							</tr>
							<tr>
								<td align="center"  colspan='2'>
									<iframe src="files.asp?fpath=<%=viewpath%>" width="830" height="600"></iframe>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center"  colspan='2'>
						<input type="button" value="Close" class="btn" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="self.close();">
					</td>
				</tr>
			</table>
		</form>
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