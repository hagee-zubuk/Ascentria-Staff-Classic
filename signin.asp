<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Files.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
ValidAko = False
ChangePass = False

Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_t WHERE upper(Username) = '" & ucase(Request("txtUN")) & "' "
rsUser.Open sqlUser, g_strCONN, 3, 1
If Not rsUser.EOF Then
	Response.Cookies("LBUSER") = Request("txtUN")
	If Request("txtPW") = Z_DoDecrypt(rsUser("password")) Then 
		Response.Cookies("LBUSERTYPE") = rsUser("type") 'gets user type - admin/default
		If rsUser("type") <> 2 Then 
			Session("UIntr") = rsUser("IntrID")
		'End If
			Session("UsrName") = rsUser("Fname") & " " & rsUser("Lname")
			Response.Cookies("LBUsrName") = rsUser("Fname") & " " & rsUser("Lname")
			If rsUser("Lname") = "" Then 
				Session("UsrName") = Session("UsrName") & " " & rsUser("Lname")
				Response.Cookies("LBUsrName") = Session("UsrName") & " " & rsUser("Lname")
			End If
			Response.Cookies("UID") = rsUser("index")
			If rsUser("reset") Then ChangePass = True
			ValidAko = True
		Else
			Session("MSG") = "Please Use https://lbis.lssne.org/interpreter to enter your Timesheet/Mileage."
		End If
	Else
		Session("MSG") = "ERROR: Invalid username and/or password."
	End If
Else
	Session("MSG") = "ERROR: Invalid username and/or password."
End If
rsUser.Close
Set rsUser = Nothing
<!-- #include file="_closeSQL.asp" -->
If ValidAko = True Then
	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(LoginLog, 8, True)
	strLog = Now & vbtab & "Successful Sign in :: User: " & Session("UsrName")
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
	'If Request.Cookies("LBUSERTYPE") <> 2 Then
	'	Response.Redirect "calendarview2.asp"
	'Else
		If ChangePass Then
			Response.Redirect "chngpass.asp"
		Else	
			Response.Redirect "calendarview2.asp"
		End If
	'End If
Else
	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(LoginLog, 8, True)
	strLog = Now & vbtab & "Error in Sign in :: User: " & Request("txtUN")
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
	Response.Redirect "default.asp"
End If
%>