<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
strFN = Request("FN") 
strNF = Request("NF")

'973CEF10E9090A841EA5A5320075BCF708E6278837AD01DDAA05A6D67CAAF4DBE7910FE0D101A27D
fname = Z_DoDecrypt(strFN)
strFN = RepPath & fname

f_fnm = Z_DoDecrypt(strNF)
If ( Len(f_fnm) > 0) Then
	strNm = Z_CleanName(f_fnm)
Else
	strNm = Z_CleanName(fname)	
End If

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strFN) Then
	Set objFile = objFSO.GetFile(strFN)
	intFSz = objFile.Size
	Set objFile = Nothing

	Response.Clear
	'Response.Status = "206 Partial Content"
	Response.Addheader "Content-Disposition", "attachment; filename=""" & strNm & """"
	Response.Addheader "Content-Length", intFSz 
	Response.Addheader "Accept-Ranges", "bytes"
	Response.Addheader "Content-Transfer-Encoding", "binary"
	Response.ExpiresAbsolute = #January 1, 2018 01:00:00#
	Response.CacheControl = "Private"
	Response.ContentType = "application/octet-stream"

	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = 1
	BinaryStream.Open
	BinaryStream.LoadFromFile strFN
	binCont = BinaryStream.Read
	BinaryStream.Close
	Response.BinaryWrite binCont
	Response.Flush()

Else
	Response.Clear
	Response.Status = "404 File Not Found"
End If
Set objFSO = Nothing
Response.End
%>
