<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function Z_MakeUniqueFileName()
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	tmpTime = replace(tmpTime, " ", "")
	Z_MakeUniqueFileName = tmpdate & tmpTime
End Function

lngRID	= Z_CLng( Request("rid")  )
lngType	= Z_CLng( Request("type") )

Set oUpload = Server.CreateObject("SCUpload.Upload")
oUpload.Upload
lngRID	= oUpload.Form("rid")
lngType	= oUpload.Form("utype")
If oUpload.Files.Count = 0 Then
	Set oUpload = Nothing
	Response.Write "<h1>Please specify a file to import (0" & lngType & "-" & lngRID & ").</h1>"
	Response.Write "<a href=""vf_upload.asp?rid=" & lngRID & """>try again</a>"
	Session("MSG") = ""
	Response.End
Else
	folderpath = uploadpath & lngRID
	folderpathvform =	uploadpath & lngRID & "\vform"
	folderpathtoll 	=	uploadpath & lngRID & "\tolls"
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(folderpath) Then fso.CreateFolder(folderpath)
	If Not fso.FolderExists(folderpathvform) Then fso.CreateFolder(folderpathvform)
	If Not fso.FolderExists(folderpathtoll) Then fso.CreateFolder(folderpathtoll)
	Set fso = Nothing

	oFileName = oUpload.Files(1).Item(1).filename
	strExt = LCase(Z_GetExt(oFileName))
	UniqueFilename = Z_MakeUniqueFileName()
	If (lngType = 0) Then
		filename = "vform"
		folder = folderpathvform
	Else
		filename = "tollsandpark"
		folder = folderpathtoll
	End If
	filename = filename & UniqueFilename & "." & strExt
	oUpload.Files(1).Item(1).Save folder, filename
	Session("MSG") = "File Saved."

	Set rsUpload = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM uploads"
	rsUpload.Open strSQL, g_strCONNupload, 1, 3
	rsUpload.AddNew
	rsUpload("RID") = lngRID
	rsUpload("type") = lngType
	rsUpload("filename") = FileName
	rsUpload("timestamp") = Now
	rsUpload("staff") = 1
	rsUpload.Update
	rsUpload.Close
	Set rsUpload = Nothing
	'Response.Redirect "viewuploads.asp?reqid=" & lngRID
%>
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
<h1>Uploaded</h1>
<script language="javascript" type="text/javascript">
	//parent.location.href = parent.location.href;
</script>
<a href="vf_upload.asp?rid=<%= lngRID %>">Upload again</a>
<br /><br /><button type="button" class="btn" name="btnDn" id="btnDn" value="" onclick="document.location='viewuploads.asp?reqid=<%= lngRID %>';">View Uploads</button>
&nbsp;&nbsp;&nbsp;&nbsp;
<button type="button" class="btn" name="btnDn" id="btnDn" value="" onclick="self.close();">Close</button>
	</div>
</body>
</html>
<%	
End If
%>