<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
	If Z_Czero(Request("ReqID")) = 0 Then 
		Session("MSG") = "ID not Found."
		Response.Redirect "openappts.asp"
	End If
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	rsApp.Open "SELECT emergency, emerfee, LBComment FROM Request_T WHERE [index] = " & Request("reqID"), g_strCONN, 1, 3
	If Not rsApp.EOF Then
		rsApp("emergency") = False
		If Request("chkEmer" & request("ctr")) = 1 Then rsApp("emergency") = True
		rsApp("Emerfee") = False
		If Request("chkEmerfee" & request("ctr")) = 1 Then rsApp("Emerfee") = True
		rsApp("LBcomment") = Trim(Request("txtLBcom" & request("ctr")))
		rsApp.Update
	End If
	rsApp.Close
	Set rsApp = Nothing
	Call SaveHist(Request("reqID"), "openappts.asp")
	Session("MSG") = "ID: " & Request("reqID") & " saved."
	Response.Redirect "openappts.asp?reload=1&frmdte=" & Request("frmdte") & "&todte=" & Request("todte") & "&selLang=" & Request("selLang")
%>