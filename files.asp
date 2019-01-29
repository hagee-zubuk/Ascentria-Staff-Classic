<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
If Trim(Request("fpath")) <> "" Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	strPath = Request("fpath")
	strExt = LCase(Z_GetExt(strPath))
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
		If (strExt = "pdf") Then
			Response.ContentType = "application/pdf"
			Response.AddHeader "Content-Disposition", "inline; filename=v-form.pdf"
		Else
			Response.ContentType = "image/" & strExt
			Response.AddHeader "Content-Disposition", "inline; filename=v-form." & strExt
		End If

		Dim lSize, lBlocks
		'Const CHUNK = 2048
		Const CHUNK = 204800
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
		Response.Write "<h1>Oops</h1><p>Unable to find the file, or access is denied accessing the file.</p><code>" & Request("fpath") & "</code>"
	End If
	Set fso = Nothing
End If
%>
