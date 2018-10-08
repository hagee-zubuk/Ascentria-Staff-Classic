<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Server.ScriptTimeout = 360000

If Request("selcli") > 0 Then
	cliid = Request("selcli")

	If Request.Cookies("LBUSERTYPE") <> 1 Then 
		Session("MSG") = "Account cannot edit information"
		Response.Redirect "client.asp?cliid=" & cliid
	End If

	Set rsCli = Server.CreateObject("ADODB.RecordSet")
	sqlCli = "SELECT * FROM c_need_T WHERE UID = " & cliid
	rsCli.Open sqlCli, g_strCONN, 1, 3
	If Not rsCli.EOF Then
		rsCli("clname") = Trim(Request("txtClilname"))
		rsCli("cfname") = Trim(Request("txtClifname"))
		rsCli("dob") = Z_CDate(Request("txtDOB"))
		rsCli("email") = Trim(Request("txtemail"))
		rsCli("comment") = Trim(Request("txtcom"))
		rsCli("asl") = False
		If Request("chk1") = 1 Then rsCli("asl") = True
		rsCli("senglish") = False
		If Request("chk2") = 1 Then rsCli("senglish") = True
		rsCli("cdeaf") = False
		If Request("chk3") = 1 Then rsCli("cdeaf") = True
		rsCli("dblind") = False
		If Request("chk4") = 1 Then rsCli("dblind") = True
		rsCli("cart") = False
		If Request("chk5") = 1 Then rsCli("cart") = True
		rsCli("cspeech") = False
		If Request("chk6") = 1 Then rsCli("cspeech") = True
		rsCli("dlow") = False
		If Request("chk7") = 1 Then rsCli("dlow") = True
		rsCli("lprint") = False
		If Request("chk8") = 1 Then rsCli("lprint") = True
		rsCli("cd") = False
		If Request("chk9") = 1 Then rsCli("cd") = True
		rsCli("alist") = False
		If Request("chk10") = 1 Then rsCli("alist") = True
		rsCli("braille") = False
		If Request("chk11") = 1 Then rsCli("braille") = True
		rsCli("laptop") = False
		If Request("chk12") = 1 Then rsCli("laptop") = True
		rsCli("other") = False
		If Request("chk13") = 1 Then rsCli("other") = True
		rsCli.Update
	End If
	rsCli.Close
	Set rsCli = Nothing
	'save pref list
	If Request("selintr") > 0 Then
		Set rsCli = Server.CreateObject("ADODB.RecordSet")
		sqlCli = "SELECT * FROM c_need_intr_T"
		rsCli.Open sqlCli, g_strCONN, 1, 3
		rsCli.AddNew
		rsCli("CID") = cliid
		rsCli("intrID") = Request("selintr")
		rsCli.Update
		rsCli.Close
		Set rsCli = Nothing
	End If
	'save exclude list
	If Request("selintrex") > 0 Then
		Set rsCli = Server.CreateObject("ADODB.RecordSet")
		sqlCli = "SELECT * FROM c_need_intr_no_T"
		rsCli.Open sqlCli, g_strCONN, 1, 3
		rsCli.AddNew
		rsCli("CID") = cliid
		rsCli("intrID") = Request("selintrex")
		rsCli.Update
		rsCli.Close
		Set rsCli = Nothing
	End If
	'delete checked pref
	y = Request("ctr")
	ctr = 0
	Do Until ctr = y + 1
		tmpID = Request("chkpref" & ctr)
		If tmpID <> "" Then
			Set rsCli = Server.CreateObject("ADODB.RecordSet")
			sqlCli = "DELETE FROM c_need_intr_T WHERE UID = " & tmpID
			rsCli.Open sqlCli, g_strCONN, 1, 3
			Set rsCli = Nothing
		End If
		ctr = ctr + 1
	Loop
	'delete checked exclu
	y = Request("ctr2")
	ctr = 0
	Do Until ctr = y + 1
		tmpID = Request("chkprefex" & ctr)
		If tmpID <> "" Then
			Set rsCli = Server.CreateObject("ADODB.RecordSet")
			sqlCli = "DELETE FROM c_need_intr_no_T WHERE UID = " & tmpID
			rsCli.Open sqlCli, g_strCONN, 1, 3
			Set rsCli = Nothing
		End If
		ctr = ctr + 1
	Loop	
	Session("MSG") = "Client saved"
Else
	Set rsCli = Server.CreateObject("ADODB.RecordSet")
	sqlCli = "SELECT * FROM c_need_T"
	rsCli.Open sqlCli, g_strCONN, 1, 3
	rsCli.AddNew
	rsCli("clname") = Trim(Request("txtClilname"))
	rsCli("cfname") = Trim(Request("txtClifname"))
	rsCli("dob") = Z_CDate(Request("txtDOB"))
	rsCli("email") = Trim(Request("txtemail"))
	rsCli("asl") = False
	If Request("chk1") = 1 Then rsCli("asl") = True
	rsCli("senglish") = False
	If Request("chk2") = 1 Then rsCli("senglish") = True
	rsCli("cdeaf") = False
	If Request("chk3") = 1 Then rsCli("cdeaf") = True
	rsCli("dblind") = False
	If Request("chk4") = 1 Then rsCli("dblind") = True
	rsCli("cart") = False
	If Request("chk5") = 1 Then rsCli("cart") = True
	rsCli("cspeech") = False
	If Request("chk6") = 1 Then rsCli("cspeech") = True
	rsCli("dlow") = False
	If Request("chk7") = 1 Then rsCli("dlow") = True
	rsCli("lprint") = False
	If Request("chk8") = 1 Then rsCli("lprint") = True
	rsCli("cd") = False
	If Request("chk9") = 1 Then rsCli("cd") = True
	rsCli("alist") = False
	If Request("chk10") = 1 Then rsCli("alist") = True
	rsCli("braille") = False
	If Request("chk11") = 1 Then rsCli("braille") = True
	rsCli("laptop") = False
	If Request("chk12") = 1 Then rsCli("laptop") = True
	rsCli("other") = False
	If Request("chk13") = 1 Then rsCli("other") = True
	rsCli.Update
	cliid = rsCli("UID")
	rsCli.Close
	Set rsCli = Nothing
	'save pref list
	If Request("selintr") > 0 Then
		Set rsCli = Server.CreateObject("ADODB.RecordSet")
		sqlCli = "SELECT * FROM c_need_intr_T"
		rsCli.Open sqlCli, g_strCONN, 1, 3
		rsCli.AddNew
		rsCli("CID") = cliid
		rsCli("intrID") = Request("selintr")
		rsCli.Update
		rsCli.Close
		Set rsCli = Nothing
	End If
	'save exclude list
	If Request("selintrex") > 0 Then
		Set rsCli = Server.CreateObject("ADODB.RecordSet")
		sqlCli = "SELECT * FROM c_need_intr_no_T"
		rsCli.Open sqlCli, g_strCONN, 1, 3
		rsCli.AddNew
		rsCli("CID") = cliid
		rsCli("intrID") = Request("selintrex")
		rsCli.Update
		rsCli.Close
		Set rsCli = Nothing
	End If
	Session("MSG") = "New Client saved"
End If
Response.Redirect "client.asp?cliid=" & cliid
%>