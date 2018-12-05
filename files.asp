<%
'<html>
'<body>

If Trim(Request("fpath")) <> "" Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	strPath = Request("fpath")
	If fso.FileExists(strPath) Then	
		Set oFileStream = Server.CreateObject("ADODB.Stream")
		oFileStream.Open
		oFileStream.Type = 1 'Binary
		On Error Resume Next
		oFileStream.LoadFromFile strPath
		If Err.Number <> 0 Then
			Response.Write "Error: " & Err.Number & " [" & Err.Message & "]"
		End If
		Response.Clear
		Response.ContentType = "application/pdf"
		Response.AddHeader "Content-Disposition", "inline; filename=v-form.pdf"

		Dim lSize, lBlocks
		'Const CHUNK = 2048
		Const CHUNK = 2048000
		lSize = oFileStream.Size
		Response.AddHeader "Content-Size", lSize
		lBlocks = 1
		Response.Buffer = False
		Do Until oFileStream.EOS Or Not Response.IsClientConnected
			Response.BinaryWrite(oFileStream.Read(CHUNK))
		Loop
		'Response.BinaryWrite oFileStream.Read

		oFileStream.Close
		Set oFileStream= Nothing
	Else
		Response.Write "<h1>Oops</h1><p>Unable to find the file, or access is denied accessing the file.</p>"
	End If
	Set fso = Nothing
end if

'</body>
'</html>
%>
