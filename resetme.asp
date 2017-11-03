<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	rsRes.Open "UPDATE appt_T SET accept = 0 WHERE UID = " & Request("appID"), g_strCONN, 1, 3
	Set rsRes = Nothing

	Session("MSG") = "Interpreter availability reset."
	Response.Redirect "openappts.asp?reload=1&frmdte=" & Request("frmdte") & "&todte=" & Request("todte") & "&selLang=" & Request("selLang")
%>