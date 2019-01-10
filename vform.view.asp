<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strVForm = ""
strTolls = ""
viewpath = ""
strRID = Z_CLng(Request("reqid"))
lngUID = Z_CLng(Request("uid"))
strSQL = "SELECT [timestamp], [filename], [rid], [type], [uid] FROM uploads WHERE [uid]=" & lngUID
Set rsUploads = Server.CreateObject("ADODB.RecordSet")
rsUploads.Open strSQL, g_strCONNupload, 3, 1
If rsUploads.EOF Then
	'shit! your file's missing
Else
	If ( rsUploads("type") = 0 ) Then
		subfold = "\vform\"
		msgtype = "Verification form"
	ElseIf ( rsUploads("type") = 1 ) Then
		subfold = "\tolls\"
		msgtype = "Toll/Parking Receipt"
	End If
	viewpath = uploadpath & strRID & subfold & rsUploads("filename")
	strMSG = "Viewing " & msgtype & " uploaded " & rsUploads("timestamp")
End If
rsUploads.Close
Set rsUploads = Nothing


%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<title>View File</title>
	<meta name="description" content="Verification Form QA">
	<meta name="author" content="Hagee@zubuk">
 	<link rel="stylesheet" type="text/css" href="css/normalize.css" />
 	<link rel="stylesheet" type="text/css" href="css/skeleton.css" />
	<link rel="stylesheet" type="text/css" href="style.css" />
  <!--[if lt IE 9]>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/html5shiv.js"></script>
  <![endif]-->
	<style>
	</style>
</head>
<body>
<div class="container" style="height: 100%;">
	<div class="row" style="height: 100%;">
		<div class="twelve columns">
			<p><%= strMSG %></p>
<iframe class="u-full-width" id="viewer" name="viewer" src="files.asp?fpath=<%=viewpath%>" width="100%" height="100%" style="height: 100%; width: 100%;"></iframe>
		</div>
	</div>
</div><!-- container -->
</body>
</html>
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {

});
// --></script>