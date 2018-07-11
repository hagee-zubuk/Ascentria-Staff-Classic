<%Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_JSON.asp" -->
<%
strTerm = Trim( Request("term") )
If Len(strTerm) < 3 Then Response.End
strSrc = "'%" & strTerm & "%'"
strSQL = "SELECT [index] AS [id], [First Name] + ' ' + [Last Name] AS [label], [First Name] + ' ' + [Last Name] AS [value] " & _
		"FROM [interpreter_T] " & _
		"WHERE [Active]=1 AND (" & _
		"[First Name] LIKE " & strSrc & " OR [Last Name] LIKE " & strSrc & _
		") ORDER BY [First Name] ASC"
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
Set jsa = jsArray()
rsIntr.Open strSQL, g_strCONN, 3, 1
Do Until rsIntr.EOF
	Set jsa(Null) = jsObject()
    For Each col In rsIntr.Fields
		jsa(Null)(col.Name) = col.Value
	Next
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing  
Response.ContentType="application/json"
' now output id, label and value for all
jsa.Flush
%>