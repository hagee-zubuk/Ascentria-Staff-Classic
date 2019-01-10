<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
lngQ = Z_CLng(Request("q"))
strSQL = "SELECT [index], [dept] FROM [Dept_T] WHERE [InstID]=" & lngQ & " ORDER BY [dept]"
Set rsTmp = Server.CreateObject("ADODB.Recordset")
rsTmp.Open strSQL, g_strCONN, 3, 1
'Response.ContentType="application/json"
Response.Write "<option value=""-1""></option>"
Do Until rsTmp.EOF
	Response.Write "<option value=""" & rsTmp("index") & """>" & rsTmp("dept") & "</option>" & vbCrLf
	rsTmp.MoveNext
Loop
rsTmp.Close
Set rsTmp = Nothing
%>