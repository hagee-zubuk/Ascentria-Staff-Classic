<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
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
	
	Set rsUploads = Server.CreateObject("ADODB.RecordSet")
	strVForm = ""
	strTolls = ""
	viewpath = ""
	strRID = Z_CLng(Request("reqid"))
	lngUID = Z_CLng(Request("uid"))
	strSQL = "SELECT [timestamp], [filename], [rid], [type], [uid] FROM uploads WHERE RID=" & strRID & " ORDER BY [timestamp] DESC"
	rsUploads.Open strSQL, g_strCONNupload, 3, 1
	Do Until rsUploads.EOF
		If (Not Z_Blank(rsUploads("rid")) And strRID <> rsUploads("rid") ) Then strRID = rsUploads("rid")
		If ( rsUploads("type") = 0 ) Then
			strVform = strVform & "&nbsp;&nbsp;&nbsp;&nbsp;<a style=""text-decoration: none;"" href=""viewuploads.asp?reqid=" & _
					strRID & "&uid=" & rsUploads("uid") & """><img src='images/zoom.gif'>" & _
					rsUploads("timestamp") & "</a><br/>" & vbCrLf
		Else
			strTolls = strTolls & "&nbsp;&nbsp;&nbsp;&nbsp;<a style=""text-decoration: none;"" href=""viewuploads.asp?reqid=" & _
					strRID & "&uid=" & rsUploads("uid") & """><img src='images/zoom.gif'>" & _
					rsUploads("timestamp") & "</a><br/>" & vbCrLf
		End If

		If lngUID = rsUploads("uid") Then
			If ( rsUploads("type") = 0 ) Then
				subfold = "\vform\"
				msgtype = "Verification form"
			ElseIf ( rsUploads("type") = 1 ) Then
				subfold = "\tolls\"
				msgtype = "Tolls and Parking Receipts"
			End If
			viewpath = uploadpath & strRID & subfold & rsUploads("filename")
			strMSG = "Viewing " & msgtype & " uploaded " & rsUploads("timestamp")
		End If

		rsUploads.MoveNext
	Loop
	rsUploads.Close
	Set rsUploads = Nothing
	'strMSG = strSQL
	'Set rsTolls = Server.CreateObject("ADODB.RecordSet")
	'rsTolls.Open "SELECT * FROM uploads WHERE RID = " & Request("reqid") & " AND [type] = 1 ORDER BY [timestamp] DESC", g_strCONNupload, 3, 1
	'Do Until rsTolls.EOF
	''	strTolls = strTolls & "&nbsp;&nbsp;&nbsp;&nbsp;<a style='text-decoration: none;' href='viewuploads.asp?ftype=1&reqid=" & Request("reqid") & 
	'"&fname=" & rsTolls("filename") & "&ts=" & rsTolls("timestamp")  & "'><img src='images/zoom.gif'>" & rsTolls("timestamp") & "</a><br>"
	''	rsTolls.MoveNext
	'Loop
	'rsTolls.Close
	'Set rsTolls = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<!doctype html>
<html lang="en">

<head>
	<meta charset="utf-8">
	<meta http-equiv="x-ua-compatible" content="ie=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">

	<title>Language Bank - VIEW UPLOADS</title>
	<script language="JavaScript" type="text/JavaScript" src="js/jquery-3.3.1.min.js" ></script>
	<link href='style.css' type='text/css' rel='stylesheet' />
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
							<tr><td align="left"><b><u>Verification Forms</u></b></td>
							</tr>
							<tr><td align="left"><nobr><%=strVform%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td align="left"><b><u>Tolls and Parking Receipts</u></b></td></tr>
							<tr><td align="left"><nobr><%=strTolls%></td></tr>
<%	If False Then %>
							<tr><td align="left" style="height: 100px;">&nbsp;</td></tr>
							<tr><td align="left"><button type="button" class="btn" name="btnUp" id="btnUp" value="">Upload File</button></td>
								</tr>
<%	End If %>								
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
									<iframe id="viewer" name="viewer" src="files.asp?fpath=<%=viewpath%>" width="830" height="600"></iframe>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center"  colspan='2'>
						<input id="btnClose" name="btnClose" type="button" value="Close" class="btn" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="self.close();">
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
<script type="text/JavaScript" language='JavaScript'>

function goUpload() {
	document.location = "vf_upload.asp?"
}		

$( document ).ready(function() {
	
	
	$("#btnUp").on("mouseover", function() {
		$("#btnUp").toggleClass("hovbtn");
	});
	$("#btnClose").on("mouseover", function() {
		$("#btnUp").toggleClass("hovbtn");
	});
/*
	$("#btnUp").click(function () {
		console.log( "upload" );
		//document.location = "vf_upload.asp?";
		$("#viewer").attr("src", "vf_upload.asp?rid=<%=strRID%>");
		$("#btnUp").hide();
	});
*/
	console.log( "ready!" );
});

</script>
