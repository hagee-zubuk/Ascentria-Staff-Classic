<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
lngRID = Z_CLng(Request("rid"))
strRefresh = "window.opener.location.reload(false);";
%>
<!-- #include file="_closeSQL.asp" -->
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta http-equiv="x-ua-compatible" content="ie=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<title>File Upload</title>
	<link href='style.css' type='text/css' rel='stylesheet' />
<style>
#btnUpld:hover {
	color:#E8E8E8;
	font-family:'trebuchet ms',helvetica,sans-serif;
	font-size:8pt;
	font-weight:bold;
	background-color:#939393;
	width: 120px;
	height: 30px;
	text-align: center;
	border-radius: 5px;	
	border: 2px solid #939393;
}
</style>	
</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<div style="margin: 100px auto;">
			<form id="frmUpload" name="frmUpload" method="POST" enctype="multipart/form-data" action="vf_upload_do.asp">
				<div style="margin-left: 50px;">
					<p>Request ID: <%=lngRID%></p>
					<input type="hidden" name="rid" id="rid" value="<%=lngRID%>" />
					<label for="ufile"><strong>Choose a file</strong></label>
					<input type="file" name="ufile" id="ufile" />
					<br />
					<table><tbody>
						<tr><td>
						<label for="utype"><strong>Upload type:</strong></label></td>
						<td>
						<input type="radio" value="0" name="utype" id="type_v" checked="checked" />&nbsp;Verification Form<br />
						<input type="radio" value="1" name="utype" id="type_t" />&nbsp;Toll and Parking Receipt<br />
						</td></tr>
					</tbody></table>
					<br />
					<input type="submit" class="btn" name="btnUpld" id="btnUpld" value="Upload File" />
					<br /><br />
					<button type="button" class="btn" name="btnDn" id="btnDn" value="" onclick="self.close();">Close</button>
				</div>
			</form>
		</div>
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
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {
	<%= strRefresh %>
});
</script>