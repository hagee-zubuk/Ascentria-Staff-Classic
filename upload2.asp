<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	h_filename = Request("hfname")
	lngReqID = Z_CLng(Request("RID"))
	disUpload = ""
	disAPP = ""
	If lngReqID <> "" Then
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT [index], uploadfile, approvePDF FROM Request_T WHERE [index] = " & 
		rsTBL.Open sqlTBL, g_strCONN, 3, 1
		If Not rsTBL.EOF Then
			If rsTBL("approvePDF") Then 
				disAPP = "disabled"
				Session("MSG") = "*This Form 604A has already been approved."
			End If
		End If
		rsTBL.Close
		Set rsTBL = Nothing
	End If
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		Set oUpload = Server.CreateObject("SCUpload.Upload")
		oUpload.Upload
		If oUpload.Files.Count = 0 Then
			Set oUpload = Nothing
			Session("MSG") = "Please specify a file to import."
			Response.Redirect "upload.asp"
		End If
		oFileName = oUpload.Files(1).Item(1).filename
		If Z_GetExt(oFileName) <> "PDF" Then
			Set oUpload = Nothing
			Session("MSG") = "Invalid File."
			Response.Redirect "upload.asp"
		End If
		' REMOVED ON 2019-03-24: no file size limitations in code'
		'oFileSize = oUpload.Files(1).Item(1).Size
		'If oFileSize > 2097152 Then
		''	Set oUpload = Nothing
		''	Session("MSG") = "File is too large."
		''	Response.Redirect "upload.asp"
		'End If
		nFileName = oUpload.Form("h_filename") & ".PDF"
		oUpload.Files(1).Item(1).Save F604AStr, nFileName
		Set oUpload = Nothing
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT [index], uploadfile, approvePDF FROM Request_T WHERE [index] = " & 
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		If Not rsTBL.EOF Then
			rsTBL("UploadFile") = True
			rsTBL("filename") = nFileName & ".PDF"
			rsTBL.Update
		End If
		rsTBL.Close
		Set rsTBL = Nothing
		Session("MSG") = "File Saved.<br>"
		disUpload = "disabled"
	End If
%>
<html>
	<head>
		<title>Language Bank - Upload</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function uploadFile() {
			if (document.frmUpload.F1.value != "") {
				filestr = document.frmUpload.F1.value.toUpperCase();
				if (filestr.indexOf(".PDF") == -1) {
					alert("ERROR: Incorrect file extension.")
					document.frmUpload.F1.value = "";
					return;
				}
				else {
					document.frmUpload.action = "upload.asp";
					document.frmUpload.submit();
				}
			}
			else {
				alert("ERROR: Please select a file.")
				return;
			}
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmUpload' enctype="multipart/form-data">
			<table align="center" border="0" width="100%" cellpadding="0" cellspacing="0">
				<tr>
					<td class='header' colspan='2'>
						<nobr>FORM 604A --&gt&gt
					</td>
				</tr>
				<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
				<tr>
					<td align="center">
						<input  class='main' type="file" name="F1" size="30" class='btn'>
					</td>
				</tr>
				<tr>
					<td align="left">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span><br>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*2 MB limit</span>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center">
						<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile();" <%=disUpload%> <%=disAPP%> class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						<input type="button" name="btnClose" value="Close" onclick="self.close();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						<input type="hidden" name="h_filename" value="<%=h_filename%>">
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
