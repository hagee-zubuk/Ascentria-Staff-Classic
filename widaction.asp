<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
server.scripttimeout = 360000
Set rsCID = Server.CreateObject("ADODB.RecordSet")
	sqlCID = "SELECT WID, PID, XID FROM Interpreter_T WHERE [index] = " & Request("selIntr")
	rsCID.Open sqlCID, g_strCONN, 1, 3
	If Not rsCID.EOF Then
		rsCID("WID") = Request("txtwid")
		rsCID("PID") = Request("txtpid")
		rsCID("XID") = Request("txtxid")
		rsCID.Update
	End If
	rsCID.Close
	Set rsCID = Nothing
	Session("MSG") = "Data saved."

response.redirect "wid.asp"
%>