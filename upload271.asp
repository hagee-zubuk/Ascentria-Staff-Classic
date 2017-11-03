<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmUpload."
enableme = "disabled"
nFileName = "none"
processme = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		Set oUpload = Server.CreateObject("SCUpload.Upload")
		oUpload.Upload
		If oUpload.Files.Count = 0 Then
			Set oUpload = Nothing
			Session("MSG") = "Please specify a file to import."
			Response.Redirect "upload271.asp"
		End If
		oFileSize = oUpload.Files(1).Item(1).Size
		If oFileSize >= 1500000 Then
			Set oUpload = Nothing
			Session("MSG") = "File is too large."
			Response.Redirect "upload271.asp"
		End If
		oFileName = oUpload.Files(1).Item(1).filename
		tmpext = Z_GetExt(oFileName)
		tmpFilename = Z_GenerateGUID()
		Do Until GUIDExists271(tmpFilename, tmpext) = False
			tmpFilename = Z_GenerateGUID()
		Loop
		nFileName = tmpFilename & "." & tmpext
		oUpload.Files(1).Item(1).Save f271Str, nFileName
		Set oUpload = Nothing
		processme = 1
		
	End If
%>
<html>
	<head>
		<title>Language Bank - Admin page - Upload 271</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			<!--
		<% if processme = 1 Then %>
			read271("<%=nFileName%>");
			
		<% End If %>
		function CalendarView(strDate)
			{
				document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
				document.frmConfirm.submit();
			}
			function uploadFile() {
			if (document.frmUpload.F1.value != "") {
					document.frmUpload.action = "upload271.asp";
					document.frmUpload.submit();
			}
			else {
				alert("ERROR: Please select a file.")
				return;
			}
			}
			function read271(cccc) {
				newwindow = window.open('print271.asp?fname=' + cccc,'','height=800,width=1000,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
				//popup here
			}
			-->
			function CalendarView(strDate)
		{
			document.frmUpload.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmUpload.submit();
		}
			</script>
	</head>
	<body>
		<form method='post' name='frmUpload' enctype="multipart/form-data">
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0' align='center' >
							<!-- #include file="_greetme.asp" -->
							<tr><td>&nbsp;</td></tr>
							<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
				<tr>
					<td align="center">
						<input  class='main' type="file" name="F1" size="30" class='main'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center">
						<input type="button" name="btnUp" value="READ 271 FILE" onclick="uploadFile();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign='bottom'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>