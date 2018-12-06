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
lngType	= oUpload.Form("type")
If oUpload.Files.Count = 0 Then
	Set oUpload = Nothing
	Response.Write "<h1>Please specify a file to import (0" & lngType & "-" & lngRID & ").</h1>"
	Response.Write "<a href=""vf_upload.asp?rid=" & lngRID & """>try again</a>"
	Response.End
End If

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
<h1>Uploaded</h1>
<script language="javascript" type="text/javascript">
	parent.location.href=parent.location.href;
</script>