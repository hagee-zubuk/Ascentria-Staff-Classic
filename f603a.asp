<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	disAPP = ""
	If Z_CZero(Request("id")) > 0 Then
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT approvePDF FROM Request_T WHERE [index] = " & Request("id")
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		If Not rsTBL.EOF Then
			If rsTBL("approvePDF") Then 
				disAPP = "disabled"
				AppNote = "*This Form 604A has already been approved."
			End If
		End If
		rsTBL.Close
		Set rsTBL = Nothing
	End If
	Set rsFile = Server.CreateObject("ADODB.RecordSet")
	sqlFile = "SELECT * FROM Request_T WHERE [index] = " & Request("id")
	rsFile.Open sqlFile, g_strCONN, 3, 1
	If Not rsFile.EOF Then
		pdfFile = F604AStr & rsFile("filename")
		h_ID = Request("id")
		'ts = rsFile("datestamp")
		viewpath = "pdf/" & rsFile("filename")
	End If
	rsFile.Close
	Set rsFile = Nothing
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.CopyFile pdfFile, tmpF604AStr 
	Set fso = Nothing
%>
<html>
	<head>
		<title>Language Bank - VIEW 604A</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function AppMe(){
			document.frmUpload.action = "action.asp?ctrl=21";
			document.frmUpload.submit();
		}
		//-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form name="frmUpload" method="POST">
			<table border=0 style="width:100%;">
				<tr>
					<td class='header' colspan='2'>
						<nobr>FORM 604A --&gt&gt
					</td>
				</tr>
				<tr>
					<td align="center">
						<%=AppNote%>
					</td>
				</tr>
				<tr>
					<td align="center">
						<iframe src="<%=viewpath%>" width="525" height="600"></iframe>
					</td>
				</tr>
				<tr>
					<td align="center">
						<input type="button" value="Approve" class="btn" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="AppMe();" <%=disAPP%>>
						<input type="button" value="Close" class="btn" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="self.close();">
						<input type="hidden" name="h_ID" value="<%=h_ID%>">
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