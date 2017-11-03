<html>
<body>
<%
If Trim(Request("fpath")) <> "" Then
	Set oFileStream = Server.CreateObject("ADODB.Stream")
	oFileStream.Open
	oFileStream.Type = 1 'Binary
	oFileStream.LoadFromFile Request("fpath")
	Response.ContentType = "application/pdf"
	Response.AddHeader "Content-Disposition", "inline; filename=" & Request("fpath")
	Response.BinaryWrite oFileStream.Read
	oFileStream.Close
	Set oFileStream= Nothing
end if
%>
</body>
</html>