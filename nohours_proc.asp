<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_UtilsMedicaid.asp" -->
<%
server.scripttimeout = 360000
Dim ArrSun(), ArrMon(), ArrTue(), ArrWed(), ArrThu(), ArrFri(), ArrSat()


arrIntr = Split(Request("selIntr"), ",")
x = 0
tmpNow = Now

Do Until x = Ubound(arrIntr) + 1
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT * FROM request_T WHERE timestamp = '" & date & "'"
	'rsApp.Open "[request_T]", g_strCONN, 1, 3
	rsApp.Open sqlApp, g_strCONN, 1, 3
	rsApp.AddNew
	rsApp("timestamp")		= tmpNow
	rsApp("reqID")			= Request("selRP")
	rsApp("appdate")		= Request("txtAppDate")
	dtADF = Z_CDate(Request("txtAppDate")) & " " & Request("txtAppTFrom")
	dtADT = Z_CDate(Request("txtAppDate")) & " " & Request("txtAppTTo")
	
	'Response.Write "From: " & dtADF & "<br />"
	'Response.Write "To  : " & dtADT & "<br />"

	rsApp("apptimefrom")	= Z_CDate(dtADF)
	rsApp("apptimeto")		= Z_CDate(dtADT)
	rsApp("langID")			= 95 'change to live
	rsApp("clname")			= GetReqlname(Request("selRP"))
	rsApp("cfname")			= GetReqfname(Request("selRP"))
	rsApp("InstID")			= 479
	rsApp("InstRate")		= 1
	If (Request("selTrain")) = 3 Then	'.. Interpreter Training Hours
		rsApp("IntrRate")	= 24
	Else
		rsApp("IntrRate")	= 20
	End If
	rsApp("deptID")			= Request("selDept")
	rsApp("IntrID")			= arrIntr(x)
	rsApp("training")		= Request("selTrain")
	rsApp.Update
	rsApp.Close
	Set rsApp = Nothing
	x = x + 1
Loop

Session("MSG") = "Appointment Saved."
Response.Redirect "nohours.asp"
%>