<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
lngIID = Z_CLng(Request("IID"))
lngUID = Z_CLng(Request.Cookies("UID"))
If lngIID < 1 or lngUID < 0 Then 
	Session("MSG") = "an error occurred trying to save the survey. please try again."
	Response.Redirect "survey2018.asp"
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT * FROM [survey2018] WHERE [iid]=" & lngIID & " AND [uid]=" & lngUID
rsSurv.Open strSQL, g_strCONN, 1, 3
If rsSurv.EOF Then
	rsSurv.AddNew
	rsSurv("iid") = lngIID
	rsSurv("uid") = lngUID
	rsSurv("txtDate") = Date
End If

For Each item In Request.Form
	If (item <> "iid") And (item <> "uid") And (item <> "txtDate") Then
		' Response.Write "Key: " & item & " --> "
		rsSurv(item) = Request.Form(item)
		' Response.Write Request.Form(item) & "<br />" & vbCrLf
	End If
Next
rsSurv.Update
rsSurv.Close
Set rsSurv = Nothing

strSQL = "SELECT [index] AS [id], [First Name] + ' ' + [Last Name] AS [name] " & _
		"FROM [interpreter_T] " & _
		"WHERE [index]=" & lngIID
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
rsIntr.Open strSQL, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	Session("SURVEY") = "Survey response for interpreter: <b style=""font-size: 150%;""><u>" & rsIntr("name") & "</u></b> saved."
End If
rsIntr.Close

Response.Redirect "survey.done.asp"
%>