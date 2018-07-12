<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
lngID = Request("ix")
If lngID < 1 Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.list.asp"
End If
Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT TOP 1 y.[index], y.[uid]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	", u.[Fname] + ' ' + u.[Lname] AS [reviewer] " & _
	"FROM [survey2018] AS y " & _
	"INNER JOIN [interpreter_T] AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T] AS u ON y.[uid]=u.[index] " & _
	"WHERE y.[index]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
If rsSurv.EOF Then
	Session("MSG") = "survey response index was not found"
	Response.Redirect "survey.list.asp"
End If
lngUID = rsSurv("uid")
Session("MSG") = "Deleted response for " & rsSurv("inter_name") & " by " & rsSurv("reviewer")
rsSurv.Close

strSQL = "UPDATE [survey2018] SET [uid]=-" & lngUID & " WHERE [index]=" & lngID
rsSurv.Open strSQL, g_strCONN, 1, 3
Set rsSurv = Nothing

Response.Redirect "survey.list.asp"
%>