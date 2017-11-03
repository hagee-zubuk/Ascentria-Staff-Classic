<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If

Function IsActive(xxx)
'cehck if interpreter is active
	IsActive = True
	Set rsAct = Server.CreateObject("ADODB.RecordSet")
	sqlAct = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsAct.Open sqlAct, g_strCONN, 3, 1
	If Not rsAct.EOF Then
		If rsAct("Active") = 0 Then IsActive = False
	End If
	rsAct.Close
	Set rsAct = Nothing	
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetPrime(xxx)
	'get primary number of requesting person
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			'GetPrime = rsRP("Phone")
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function GetPrime2(xxx)
	'get primary number of interpreter
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 1 Or rsRP("prime") = 2 Then
			'GetPrime = rsRP("Phone")
			GetPrime2 = ""
		ElseIf rsRP("prime") = 3 Then
			GetPrime2 = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function

tmpPage = "document.frmAssign."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpStat = rsConfirm("Status")
	If rsConfirm("Status") = 0 Then stat = "checked"
	If rsConfirm("Status") = 1 Then comp = "checked"
	If rsConfirm("Status") = 2 Then misd = "checked"
	If rsConfirm("Status") = 3 Then canc = "checked"
	If rsConfirm("Status") = 4 Then canc2 = "checked"
	tmpMiss = rsConfirm("Missed")
	tmpCancel = rsConfirm("Cancel")
	tmpClient = ""
	If rsConfirm("client") = 1 Then tmpClient = " (LSS Client)"
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	tmpAddr = rsConfirm("CliAdrI") & " " & rsConfirm("caddress") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	tmpFon = rsConfirm("Cphone")
	tmpAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = Ctime(rsConfirm("appTimeFrom"))
	tmpAppTFrom2 = rsConfirm("appTimeFrom")
	tmpAvail = Weekday(tmpAppDate) & "," & Hour(tmpAppTFrom)
	tmpAppTTo = Ctime(rsConfirm("appTimeTo"))
	tmpAppTTo2 = rsConfirm("appTimeTo")
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = 1 Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	chkVer = ""
	If rsConfirm("Verified") = 1 Then chkVer = "checked"
	chkPaid = ""
	If Not IsNull(rsConfirm("Processed")) Or rsConfirm("Processed") <> "" Then chkPaid = "checked"
	tmpBilHrs = rsConfirm("Billable")
	'tmpActDate = rsEdit("adate")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpcomintr = rsConfirm("intrcomment")
	tmpCombil = rsConfirm("bilcomment")
	tmpLBcom = rsConfirm("LBcomment")
	tmpHPID = Z_CZero(rsConfirm("HPID"))
End If
rsConfirm.Close
Set rsConfirm = Nothing
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		Session("MSG") = Session("MSG") & "<br>" & rsHP("lbcom") & "<br>" & rsHP("comment")
	End If
	rsHp.Close
	Set rsHp = Nothing
End If
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & RP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(RP)
End If
rsReq.Close
Set rsReq = Nothing
'GET AVAILABLE LANGUAGES
'Set rsLang = Server.CreateObject("ADODB.RecordSet")
'sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
'rsLang.Open sqlLang, g_strCONN, 3, 1
'Do Until rsLang.EOF
'	tmpL = ""
'	If tmpLang = "" Then tmpLang = -1
'	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
'	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
'	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
'		"return " & rsLang("index") & ";}"
'	rsLang.MoveNext
'Loop
'rsLang.Close
'Set rsLang = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
	
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE [index] = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
End If
rsLang.Close
Set rsLang = Nothing 
'GET INTERPRETER RATES
Set rsRate2 = Server.CreateObject("ADODB.RecordSet")
sqlRate2 = "SELECT * FROM rate2_T ORDER BY Rate2"
rsRate2.Open sqlRate2, g_strCONN, 3, 1
Do Until rsRate2.EOF
	RateKo2 = ""
	strRate2 = strRate2 & "<option value='" & rsRate2("Rate2") & "'>$" & Z_FormatNumber(rsRate2("Rate2"), 2) & "</option>" & vbCrLf
	rsRate2.MoveNext
Loop
rsRate2.Close
Set rsRate2 = Nothing
'GET INTERPRETER INFO
'Set rsIntr2 = Server.CreateObject("ADODB.RecordSet")
'sqlInst2 = "SELECT * FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
'rsIntr2.Open sqlInst2, g_strCONN, 3, 1
'Do Until rsIntr2.EOF
'	CtrLang = 0
'	If rsIntr2("Language1") <> "" Then CtrLang =  CtrLang + 1 
'	If rsIntr2("Language2") <> "" Then CtrLang =  CtrLang + 1
'	If rsIntr2("Language3") <> "" Then CtrLang =  CtrLang + 1
'	If rsIntr2("Language4") <> "" Then CtrLang =  CtrLang + 1
'	If rsIntr2("Language5") <> "" Then CtrLang =  CtrLang + 1
'	strJScript2 = strJScript2 & "if (Intr == " & rsIntr2("Index") & ") " & vbCrLf & _
'		"{document.frmAssign.selIntr.value = """ & rsIntr2("Index") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrEmail.value = """ & rsIntr2("E-mail") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrP1.value = """ & rsIntr2("phone1") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrExt.value = """ & rsIntr2("P1Ext") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrP2.value = """ & rsIntr2("phone2") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrFax.value = """ & rsIntr2("fax") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrAddr.value = """ & rsIntr2("address1") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrCity.value = """ & rsIntr2("City") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrState.value = """ & rsIntr2("State") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrZip.value = """ & rsIntr2("Zip Code") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrAddrI.value = """ & rsIntr2("IntrAdrI") &"""; " & vbCrLf & _
'		"document.frmAssign.txtIntrRate.value = """ & rsIntr2("Rate") &"""; " & vbCrLf & _
'		"document.frmAssign.LangCtr.value = " & CtrLang &"; " & vbCrLf & _
'		"document.frmAssign.Lang1.value = GetLangID(""" & Trim(rsIntr2("Language1")) & """); " & vbCrLf & _
'		"document.frmAssign.Lang2.value = GetLangID(""" & Trim(rsIntr2("Language2")) &"""); " & vbCrLf & _
'		"document.frmAssign.Lang3.value = GetLangID(""" & Trim(rsIntr2("Language3")) &"""); " & vbCrLf & _
'		"document.frmAssign.Lang4.value = GetLangID(""" & Trim(rsIntr2("Language4")) &"""); " & vbCrLf & _
'		"document.frmAssign.Lang5.value = GetLangID(""" & Trim(rsIntr2("Language5")) &"""); " & vbCrLf 
'		If rsIntr2("InHouse") = 1 Then 
'			strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = true; " & vbCrLf 
'		Else
'			strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = false; " & vbCrLf 
'		End If
'		If rsIntr2("prime") = 0 Then
'			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[0].checked = true;" & vbCrLf 
'		ElseIf rsIntr2("prime") = 1 Then
'			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[1].checked = true;" & vbCrLf 
'		ElseIf rsIntr2("prime") = 2 Then
'			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[3].checked = true;" & vbCrLf 
'		ElseIf rsIntr2("prime") = 3 Then
'			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[2].checked = true;" & vbCrLf 
'		End If
'		strJScript2 = strJScript2 & "}"
		
'		IntrSel = ""
'		If CInt(tmpIntr) = rsIntr2("index") Then IntrSel = "selected"
'		strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntr2("Index") & "'>" & rsIntr2("last name") & ", " & rsIntr2("first name") & "</option>" & vbCrlf
'		tmpIntrName = CleanMe(rsIntr2("last name")) & ", " & CleanMe(rsIntr2("first name"))
'		strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
'			"ChoiceReq.value = " & rsIntr2("index") & ";" & vbCrLf & _
'			"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
'			"document.frmAssign.selIntr.appendChild(ChoiceReq);}" & vbCrLf
'		
'		'strIntrCHK = strIntrCHK & "if (document.frmAssign.txtIntrLname.value == """ & Trim(rsIntr2("last name")) & """ && document.frmAssign.txtIntrFname.value == """ & Trim(rsIntr2("First name")) & """) " & vbCrLf & _
'		'"{var ans = window.confirm(""Interpreter's name already exists. Click on Cancel to rename. Click on OK to continue.""); " & vbCrLf & _
'		'"{if (ans){ " & vbCrLf & _
		'"pnt = 1; " & vbCrLf & _
		'"} " & vbCrLf & _
		'"else " & vbCrLf & _
		'"{ " & vbCrLf & _
		'"return; " & vbCrLf & _
		'"} " & vbCrLf & _
		'"} " & vbCrLf & _
		'"} " & vbCrLf & _
		'"else " & vbCrLf & _
		'"{pnt = 1; " & vbCrLf & _
		'"} " & vbCrLf
			
'		rsIntr2.MoveNext
'Loop
'rsIntr2.Close
'Set rsIntr2 = Nothing

'GET ALLOWED INTERPRETER
Set rsLangIntr = Server.CreateObject("ADODB.RecordSet")
sqlLangIntr = "SELECT * FROM language_T WHERE [index] = " & tmpLang
rsLangIntr.Open sqlLangIntr, g_strCONN, 3, 1
If Not rsLangIntr.EOF Then
	IntrLang = UCase(rsLangIntr("Language"))
	strIntrLang = strIntrLang & "if (dialect == " & rsLangIntr("index") & "){" & vbCrLf
	Set rsIntrLang = Server.CreateObject("ADODB.RecordSet")
	sqlIntrLang = "SELECT * FROM interpreter_T WHERE (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & "' OR Upper(Language3) = '" & IntrLang & _
		"' OR Upper(Language4) = '" & IntrLang & "' OR Upper(Language5) = '" & IntrLang & "' OR Upper(Language6) = '" & IntrLang & "') AND Active = 1 ORDER BY [Last Name], [First Name]" 
	rsIntrLang.Open sqlIntrLang, g_strCONN, 3, 1
	Do Until rsIntrLang.EOF
		IntrSel = ""
		If CInt(tmpIntr) = rsIntrLang("index") Then IntrSel = "selected"
		mark = 0
		tmpIntrName = CleanMe(rsIntrLang("last name")) & ", " & CleanMe(rsIntrLang("first name"))
		If OnVacation(rsIntrLang("index"), tmpAppDate) = False Then 'If isNull(rsIntrLang("vacfrom")) Then
			If tmpIntr = rsIntrLang("index") Or (Avail(rsIntrLang("index"), tmpAvail) And NotRestrict(rsIntrLang("index"), tmpInst, tmpDept)) Then
				mark = 1
				'strIntrLang = strIntrLang	& "if(intr != "& rsIntrLang("index") & ")" & vbCrLf & _
				'	"{var ChoiceIntr = document.createElement('option');" & vbCrLf & _
				'	"ChoiceIntr.value = " & rsIntrLang("index") & ";" & vbCrLf & _
				'	"ChoiceIntr.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
				'	"document.frmAssign.selIntr.appendChild(ChoiceIntr);}" & vbCrLf
				rest = ""
				If NotRestrict(rsIntrLang("index"), tmpInst, tmpDept) = false Then rest = " (restricted)"
				strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntrLang("Index") & "'>" & rsIntrLang("last name") & ", " & rsIntrLang("first name") & rest & "</option>" & vbCrlf
				tmpIntrName = CleanMe(rsIntrLang("last name")) & ", " & CleanMe(rsIntrLang("first name"))
				'strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
				'	"ChoiceReq.value = " & rsIntrLang("index") & ";" & vbCrLf & _
				'	"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
				'	"document.frmAssign.selIntr.appendChild(ChoiceReq);}" & vbCrLf
			End If
		'Else
		'	If Not (tmpAppDate >= rsIntrLang("vacfrom") And tmpAppDate <= rsIntrLang("vacto")) Then
		'		If tmpIntr = rsIntrLang("index") Or (Avail(rsIntrLang("index"), tmpAvail) And NotRestrict(rsIntrLang("index"), tmpInst)) Then
		'			mark = 1
		'			'strIntrLang = strIntrLang	& "if(intr != "& rsIntrLang("index") & ")" & vbCrLf & _
		'			'	"{var ChoiceIntr = document.createElement('option');" & vbCrLf & _
		'			'	"ChoiceIntr.value = " & rsIntrLang("index") & ";" & vbCrLf & _
		'			'	"ChoiceIntr.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
		'			'	"document.frmAssign.selIntr.appendChild(ChoiceIntr);}" & vbCrLf
		'			rest = ""
		'			If NotRestrict(rsIntrLang("index"), tmpInst) = false Then rest = " (restricted)"
		'			strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntrLang("Index") & "'>" & rsIntrLang("last name") & ", " & rsIntrLang("first name") & rest & "</option>" & vbCrlf
		'			tmpIntrName = CleanMe(rsIntrLang("last name")) & ", " & CleanMe(rsIntrLang("first name"))
					'strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
					'	"ChoiceReq.value = " & rsIntrLang("index") & ";" & vbCrLf & _
					'	"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
					'	"document.frmAssign.selIntr.appendChild(ChoiceReq);}" & vbCrLf	
		'		End If		
		'	End if
		End If
		If mark = 0 And NotRestrict(rsIntrLang("index"), tmpInst, tmpDept) And OnVacation(rsIntrLang("index"), tmpAppDate) = False Then
			strIntr3 = strIntr3 & "<option value='" & rsIntrLang("index") & "' " & IntrSel & ">" & tmpIntrName & "</option>" & vbCrlf
		End If
		CtrLang = 0
	If rsIntrLang("Language1") <> "" Then CtrLang =  CtrLang + 1 
	If rsIntrLang("Language2") <> "" Then CtrLang =  CtrLang + 1
	If rsIntrLang("Language3") <> "" Then CtrLang =  CtrLang + 1
	If rsIntrLang("Language4") <> "" Then CtrLang =  CtrLang + 1
	If rsIntrLang("Language5") <> "" Then CtrLang =  CtrLang + 1
	If rsIntrLang("Language6") <> "" Then CtrLang =  CtrLang + 1
	intrRate = rsIntrLang("Rate")
	If Z_EligibleHigherPay(tmpLang) Then IntrRate = Z_GetHigherPay(rsIntrLang("Rate"), rsIntrLang("index"))
	strJScript2 = strJScript2 & "if (Intr == " & rsIntrLang("Index") & ") {" & vbCrLf & _
		"document.frmAssign.txtIntrEmail.value = """ & rsIntrLang("E-mail") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrP1.value = """ & rsIntrLang("phone1") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrExt.value = """ & rsIntrLang("P1Ext") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrP2.value = """ & rsIntrLang("phone2") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrFax.value = """ & rsIntrLang("fax") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrAddr.value = """ & rsIntrLang("address1") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrCity.value = """ & rsIntrLang("City") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrState.value = """ & rsIntrLang("State") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrZip.value = """ & rsIntrLang("Zip Code") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrAddrI.value = """ & rsIntrLang("IntrAdrI") &"""; " & vbCrLf & _
		"document.frmAssign.txtIntrRate.value = """ & IntrRate &"""; " & vbCrLf & _
		"document.frmAssign.LangCtr.value = " & CtrLang &"; " & vbCrLf & _
		"document.frmAssign.chkAvail.value = " & SkedCheck(rsIntrLang("Index"), Request("ID"), tmpAppDate, tmpAppTFrom2, tmpAppTTo2) &"; " & vbCrLf & _
		"document.frmAssign.Lang1.value = GetLangID(""" & Trim(rsIntrLang("Language1")) & """); " & vbCrLf & _
		"document.frmAssign.Lang2.value = GetLangID(""" & Trim(rsIntrLang("Language2")) &"""); " & vbCrLf & _
		"document.frmAssign.Lang3.value = GetLangID(""" & Trim(rsIntrLang("Language3")) &"""); " & vbCrLf & _
		"document.frmAssign.Lang4.value = GetLangID(""" & Trim(rsIntrLang("Language4")) &"""); " & vbCrLf & _
		"document.frmAssign.Lang5.value = GetLangID(""" & Trim(rsIntrLang("Language5")) &"""); " & vbCrLf & _
		"document.frmAssign.Lang6.value = GetLangID(""" & Trim(rsIntrLang("Language6")) &"""); " & vbCrLf 
		If rsIntrLang("InHouse") = 1 Then 
			strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = true; " & vbCrLf 
		Else
			strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = false; " & vbCrLf 
		End If
		If rsIntrLang("prime") = 0 Then
			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[0].checked = true;" & vbCrLf 
		ElseIf rsIntrLang("prime") = 1 Then
			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[1].checked = true;" & vbCrLf 
		ElseIf rsIntrLang("prime") = 2 Then
			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[3].checked = true;" & vbCrLf 
		ElseIf rsIntrLang("prime") = 3 Then
			strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[2].checked = true;" & vbCrLf 
		End If
		strJScript2 = strJScript2 & "}"
		rsIntrLang.MoveNext
	Loop
	rsIntrLang.Close
	Set rsIntrLang = Nothing
	rsLangIntr.MoveNext
	strIntrLang = strIntrLang & "}"
End If
rsLangIntr.Close
Set rsLangIntr = Nothing
'GET INTERPRETER INFO IF INACTIVE BUT ASSIGNED
If (Z_CZero(tmpIntr) <> 0 Or tmpIntr <> "-1") And Request("ID") <> "" Then
	'CHECK IF INACTIVE
	If IsActive(tmpIntr) = False Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
		rsIntr.Open sqlIntr, g_strCONN, 3, 1
		If Not rsIntr.EOF Then
			strIntr = strIntr	& "<option selected value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "(INACTIVE)</option>" & vbCrlf
			tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
			'strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			'	"ChoiceReq.value = " & rsIntr("index") & ";" & vbCrLf & _
			'	"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
			'	"document.frmAssign.selIntr.appendChild(ChoiceReq);}" & vbCrLf
			CtrLang = 0
			If rsIntr("Language1") <> "" Then CtrLang =  CtrLang + 1 
			If rsIntr("Language2") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language3") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language4") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language5") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language6") <> "" Then CtrLang =  CtrLang + 1
			intrRate = rsIntr("Rate")
			If Z_EligibleHigherPay(tmpLang) Then IntrRate = Z_GetHigherPay(rsIntr("Rate"), rsIntr("Index"))
			strJScript2 = strJScript2 & "if (Intr == " & rsIntr("Index") & ") {" & vbCrLf & _
				"document.frmAssign.txtIntrEmail.value = """ & rsIntr("E-mail") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrP1.value = """ & rsIntr("phone1") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrExt.value = """ & rsIntr("P1Ext") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrP2.value = """ & rsIntr("phone2") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrFax.value = """ & rsIntr("fax") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrAddr.value = """ & rsIntr("address1") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrCity.value = """ & rsIntr("City") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrState.value = """ & rsIntr("State") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrZip.value = """ & rsIntr("Zip Code") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrAddrI.value = """ & rsIntr("IntrAdrI") &"""; " & vbCrLf & _
				"document.frmAssign.txtIntrRate.value = """ & intrRate &"""; " & vbCrLf & _
				"document.frmAssign.LangCtr.value = " & CtrLang &"; " & vbCrLf & _
				"document.frmAssign.chkAvail.value = " & SkedCheck(rsIntr("Index"), Request("ID"), tmpAppDate, tmpAppTFrom2, tmpAppTTo2) &"; " & vbCrLf & _
				"document.frmAssign.Lang1.value = GetLangID(""" & Trim(rsIntr("Language1")) & """); " & vbCrLf & _
				"document.frmAssign.Lang2.value = GetLangID(""" & Trim(rsIntr("Language2")) &"""); " & vbCrLf & _
				"document.frmAssign.Lang3.value = GetLangID(""" & Trim(rsIntr("Language3")) &"""); " & vbCrLf & _
				"document.frmAssign.Lang4.value = GetLangID(""" & Trim(rsIntr("Language4")) &"""); " & vbCrLf & _
				"document.frmAssign.Lang5.value = GetLangID(""" & Trim(rsIntr("Language5")) &"""); " & vbCrLf & _
				"document.frmAssign.Lang6.value = GetLangID(""" & Trim(rsIntr("Language6")) &"""); " & vbCrLf 
			If rsIntr("InHouse") = 1 Then 
				strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = true; " & vbCrLf 
			Else
				strJScript2 = strJScript2 & "document.frmAssign.chkInHouse.checked = false; " & vbCrLf 
			End If
			If rsIntr("prime") = 0 Then
				strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[0].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 1 Then
				strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[1].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 2 Then
				strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[3].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 3 Then
				strJScript2 = strJScript2 & "document.frmAssign.radioPrim2[2].checked = true;" & vbCrLf 
			End If
			strJScript2 = strJScript2 & "}"
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End If
If strIntr3 <> "" Then strIntr = strIntr & "<option value='0'>-----</option>" & vbCrlf & strIntr3
'GET INSTITUTION RATES
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM rate_T ORDER BY Rate"
rsRate.Open sqlRate, g_strCONN, 3, 1
Do Until rsRate.EOF
	RateKo = ""
	If tmpInstRate = rsRate("Rate") Then Rateko = "selected"
	strRate1 = strRate1 & "<option " & Rateko & " value='" & rsRate("Rate") & "'>$" & Z_FormatNumber(rsRate("Rate"), 2) & "</option>" & vbCrLf
	rsRate.MoveNext
Loop
rsRate.Close
Set rsRate = Nothing
'INTERPRETER CHECKER INACTIVE
Set rsIntrCHK = Server.CreateObject("ADODB.RecordSet")
sqlIntrCHK = "SELECT * FROM interpreter_T WHERE Active = 0"
rsIntrCHK.Open sqlIntrCHK, g_strCONN, 3, 1
Do Until rsIntrCHK.EOF
	strIntrCHK = strIntrCHK & "if (document.frmAssign.txtIntrLname.value == """ & Trim(rsIntrCHK("last name")) & """ && document.frmAssign.txtIntrFname.value == """ & Trim(rsIntrCHK("First name")) & """) " & vbCrLf & _
		"{alert(""Interpreter's name already exists but is currently inactive.\nPlease contact your system adminitrator.""); " & vbCrLf & _
		"return; " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{pnt = 1; " & vbCrLf & _
		"} " & vbCrLf
	rsIntrCHK.MoveNext 
Loop
rsIntrCHK.Close
Set rsIntrCHK = Nothing
'INTERPRETER FEEDBACK
Set rsIntrFB = Server.CreateObject("ADODB.RecordSet")
sqlIntrFB = "SELECT * FROM InterpreterEval_T WHERE intrID = " & tmpIntr & " AND appID = " & Request("ID") & " AND UID = " & Request.Cookies("UID")
rsIntrFB.Open sqlIntrFB, g_strCONN, 3, 1
If Not rsIntrFB.EOF Then
	tmpintrFeed = rsIntrFB("comment")
End If
rsIntrFB.Close
Set rsIntrFB = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Edit Interpreter - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function GetLangID(xxx)
		{
			if (xxx == "" )
			{
				return 0;
			}
			<%=strLangChk%>
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function CalendarView(strDate)
		{
			document.frmAssign.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAssign.submit();
		}
		function SaveAss(xxx, yyy, zzz)
		{
			if (document.frmAssign.selIntr.value != 0)
			{
				if (document.frmAssign.radioPrim2[0].checked == true && document.frmAssign.txtIntrEmail.value == "")
				{
					alert("ERROR: Please supply an E-mail address to interpreter."); 
					document.frmAssign.txtIntrEmail.focus();
					return;
				}
				if (document.frmAssign.radioPrim2[1].checked == true && document.frmAssign.txtIntrP1.value == "")
				{
					alert("ERROR: Please supply a Home Number to interpreter."); 
					document.frmAssign.txtIntrP1.focus();
					return;
				}
				if (document.frmAssign.radioPrim2[3].checked == true && document.frmAssign.txtIntrP2.value == "")
				{
					alert("ERROR: Please supply a Mobile Number to interpreter."); 
					document.frmAssign.txtIntrP2.focus();
					return;
				}
				if (document.frmAssign.radioPrim2[2].checked == true && document.frmAssign.txtIntrFax.value == "")
				{
					alert("ERROR: Please supply a Fax Number to interpreter."); 
					document.frmAssign.txtIntrFax.focus();
					return;
				}
				//CHECK VALID FAX
				if (document.frmAssign.radioPrim2[2].checked == true && document.frmAssign.txtIntrFax.value != "")
				{
					var tmpFax =  document.frmAssign.txtIntrFax.value
					tmpFax = tmpFax.replace("-", "")
					if (tmpFax.length < 10) 
					{
						alert("ERROR: Please include area code in Fax Number to interpreter."); 
						document.frmAssign.txtIntrFax.focus();
						return;
					}
				}
			}
			<%=strIntrCHK%>
			//CHECK IF INTERPRETER ALLOWED
			<% If Cint(Request.Cookies("LBUSERTYPE")) = 1 Then %>
				if (document.frmAssign.chkAvail.value == 1) {
					var ans = window.confirm("WARNING: Interpreter already has an appointment for this date and time range.\nPlease check the calendar. \nClick OK to override. \nClick Cancel to stop."); 
					if (ans) {
						pnt = 1;
					}
					else {
						return;
					}
				}
			<% Else %>
				if (document.frmAssign.chkAvail.value == 1) {
					var ans = window.confirm("WARNING: Interpreter already has an appointment for this date and time range.\nPlease check the calendar. \nClick OK to override. \nClick Cancel to stop."); 
					if (ans) {
						pnt = 1;
					}
					else {
						return;
					}
				}
			<% End If %>
				pnt = 1;
	
			if (pnt == 1)
			{
				document.frmAssign.action = "action.asp?ctrl=11&ReqID=" + xxx + "&LangID=" + yyy + "&IntrID=" + zzz;
				document.frmAssign.submit();
			}
		}
		function IntrInfo(Intr)
		{	
			if (Intr == -1 || Intr == 0)
			{
				//document.frmAssign.selIntr.value = -1;
				document.frmAssign.txtIntrEmail.value = ""; 
				document.frmAssign.txtIntrP1.value = ""; 
				document.frmAssign.txtIntrExt.value = ""; 
				document.frmAssign.txtIntrP2.value = ""; 
				document.frmAssign.txtIntrFax.value = ""; 
				document.frmAssign.txtIntrAddr.value = ""; 
				document.frmAssign.txtIntrCity.value = ""; 
				document.frmAssign.txtIntrState.value = ""; 
				document.frmAssign.txtIntrZip.value = "";
				document.frmAssign.txtIntrAddrI.value = "";  
				document.frmAssign.chkInHouse.checked = false; 
				document.frmAssign.txtIntrRate.value = 0;
				document.frmAssign.radioPrim2[2].checked = true;
				document.frmAssign.chkAvail.value = 0;
				document.frmAssign.LangCtr.value = 0;
				document.frmAssign.Lang1.value = "";
				document.frmAssign.Lang2.value = "";
				document.frmAssign.Lang3.value = "";
				document.frmAssign.Lang4.value = "";
				document.frmAssign.Lang5.value = "";
				document.frmAssign.Lang6.value = "";
			}
			<%=strJScript2%>
			chkPrim2();
		}
		function chkPrim2()
		{
			if (document.frmAssign.radioPrim2[0].checked == true)
			{
				document.frmAssign.txtPRim2.value = "E-Mail";
			}
			if (document.frmAssign.radioPrim2[1].checked == true)
			{
				document.frmAssign.txtPRim2.value = "Home Phone";
			}
			if (document.frmAssign.radioPrim2[2].checked == true)
			{
				document.frmAssign.txtPRim2.value = "Fax";
			}
			if (document.frmAssign.radioPrim2[3].checked == true)
			{
				document.frmAssign.txtPRim2.value = "Mobile Phone";
			}
		}
		function LangChoice(dialect, intr)
		{
			 var i;
			for(i=document.frmAssign.selIntr.options.length-1;i>=1;i--)
			{
				if (intr != "undefined"  )
				{
					if (document.frmAssign.selIntr.options[i].value != intr)
					{
						document.frmAssign.selIntr.remove(i);
					}
				}
				else
				{
					document.frmAssign.selIntr.remove(i);
				}
			}
			<%=strIntrLang%>
		}
		function hideNewIntr() 
		{
			if (document.frmAssign.txtIntrLname.value == "" && document.frmAssign.txtIntrFname.value == "")
			{	
				document.frmAssign.txtIntrLname.style.visibility = 'hidden';
				document.frmAssign.txtIntrFname.style.visibility = 'hidden';
				document.frmAssign.txtcoma.style.visibility = 'hidden';
				document.frmAssign.txtformat.style.visibility = 'hidden';
				//document.frmAssign.btnNewIntr.value = 'NEW';
				document.frmAssign.txtIntrLname.value = "";
				document.frmAssign.txtIntrFname.value = "";
				document.frmAssign.selIntrRate.style.visibility = 'hidden';
				document.frmAssign.HnewIntr.value = 'NEW';
			}
			else
			{
				document.frmAssign.txtIntrLname.style.visibility = 'visible';
				document.frmAssign.txtIntrFname.style.visibility = 'visible';
				document.frmAssign.txtcoma.style.visibility = 'visible';
				document.frmAssign.txtformat.style.visibility = 'visible';
				document.frmAssign.selIntrRate.style.visibility = 'visible';
				//document.frmAssign.btnNewIntr.value = 'BACK';
				document.frmAssign.selIntr.disabled = true;
				document.frmAssign.txtIntrLname.value = '<%=tmpIntrLname%>';
				document.frmAssign.txtIntrFname.value = '<%=tmpIntrFname%>';
				document.frmAssign.txtIntrEmail.value = '<%=tmpIntrEmail%>';
				document.frmAssign.txtIntrP1.value = '<%=tmpIntrP1%>';
				document.frmAssign.txtIntrExt.value = '<%=tmpIntrExt%>';
				document.frmAssign.txtIntrFax.value = '<%=tmpIntrFax%>';
				document.frmAssign.txtIntrP2.value = '<%=tmpIntrP2%>';
				document.frmAssign.txtIntrAddr.value = '<%=tmpIntrAddr%>';
				document.frmAssign.txtIntrCity.value = '<%=tmpIntrCity%>';
				document.frmAssign.txtIntrState.value = '<%=tmpIntrState%>';
				document.frmAssign.txtIntrZip.value = '<%=tmpIntrZip%>';
				document.frmAssign.txtIntrAddrI.value = '<%=tmpNewIntrAddrI%>';  
				document.frmAssign.selIntrRate.value = '<%=tmpIntrRate%>';
				document.frmAssign.LangCtr.value = 0;
				document.frmAssign.Lang1.value = "";
				document.frmAssign.Lang2.value = "";
				document.frmAssign.Lang3.value = "";
				document.frmAssign.Lang4.value = "";
				document.frmAssign.Lang5.value = "";
				document.frmAssign.Lang6.value = "";
				document.frmAssign.HnewIntr.value = 'BACK';
			}
		}
		function IntrShowMe()
		{
			if (document.frmAssign.chkAll2.checked == true) 
			{
				for(i=document.frmAssign.selIntr.options.length-1;i>=1;i--)
				{
					document.frmAssign.selIntr.remove(i);
				}
				<%=strIntr2%>
			}
			else
			{
				LangChoice(<%=tmpLang%>);
			}
		}
		function textboxchangeIntr() 
		{
			
	
				//document.frmAssign.btnNewIntr.value = 'NEW';
				document.frmAssign.selIntr.disabled = false;
				document.frmAssign.txtNewIntr.value = "";
				document.frmAssign.selIntrRate.value = 0;
				document.frmAssign.txtIntrLname.style.visibility = 'hidden';
				document.frmAssign.txtIntrFname.style.visibility = 'hidden';
				document.frmAssign.txtcoma.style.visibility = 'hidden';
				document.frmAssign.txtformat.style.visibility = 'hidden';
				document.frmAssign.selIntrRate.style.visibility = 'hidden';
				IntrInfo(document.frmAssign.selIntr.value);
				document.frmAssign.HnewIntr.value = 'NEW';
		
		}
		function PopMe2(tmpDate, tmpTime, tmpLang)
		{
			if (tmpDate == "" || tmpTime == "" || tmpLang == "")
			{
				alert("Error: Appointment Date/Time/Language is Required.")
				return;
			}
			else
			{
				var zDate = tmpDate; 
				var zTime = tmpTime;
				var zLang = tmpLang
				var zDT = zDate + "|" + zTime + "|" + zLang
				//alert(zipus);
				newwindow2 = window.open('IntrAvail.asp?AppInfo=' + zDT,'name','width=400,scrollbars=0,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow2.focus()}
			}
		}
		function CertMe(xxx)
		{
			newwindow2 = window.open('cert.asp?ID=' + xxx,'name','width=400,scrollbars=0,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow2.focus()}
		}
		-->
		</script>
		<body onload='IntrInfo(<%=tmpIntr%>); hideNewIntr();'>
			<form method='post' name='frmAssign'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Edit Interpreter</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='center' colspan='10'>
										<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
											<table border='0' cellspacing='1'>		
												<tr>
													<td><span class='error'><%=Session("MSG")%></span></td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%></td>
								</tr>
								<tr>
									<td align='right'>Timestamp:</td>
									<td class='confirm' width='300px'><%=TS%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%></td>
								</tr>
								<tr>
									<td align='right'>Department:</td>
									<td class='confirm'><%=tmpDname%></td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
								<tr>
									<td align='right' width='15%'>Rate:</td>
									<td>
										<input class='main' size='5' maxlength='5'  readonly name='txtInstRate' value='<%=tmpInstRate%>'>
										<select class='seltxt' style='width: 70px;' name='selInstRate'>
											<option value='0' >&nbsp;</option>
											<%=strRate1%>
										</select>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">
											*Rate varies per request
										</span>
									</td>
								</tr>
								<% Else %>
												<input class='main' size='5' maxlength='5' type='hidden' name='txtInstRate' value='<%=tmpInstRate%>'>
											<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Requesting Person:</td>
									<td class='confirm'><%=tmpRP%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>
								<tr>
									<td align='right'>Client Name:</td>
									<td class='confirm'><%=tmpName%></td>
								</tr>
								<tr>
									<td align='right'>Client Address:</td>
									<td class='confirm'><%=tmpAddr%></td>
								</tr>
								<tr>
									<td align='right'>Language:</td>
									<td class='confirm'><%=tmpSalita%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Date:</td>
									<td class='confirm'><%=tmpAppDate%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Time:</td>
									<td class='confirm'><%=tmpAppTFrom%> - <%=tmpAppTTo%></td>
								</tr>
								<tr>
									<td align='right'>Docket Number:</td>
									<td class='confirm'><%=tmpDoc%></td>
								</tr>
								<tr>
									<td align='right'>Court Room No:</td>
									<td class='confirm'><%=tmpCRN%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Appointment Comment:</td>
									<td class='confirm'><%=tmpCom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information</td>
								</tr>
									<tr>
										<td align='right'>Interpreter:</td>
										<td>
											<select class='seltxt' name='selIntr' style='width: 200px;' onchange='JavaScript:IntrInfo(document.frmAssign.selIntr.value);' onfocus='JavaScript:IntrInfo(document.frmAssign.selIntr.value);'>
												<option value='0'>&nbsp;</option>
												<%=strIntr%>
											</select>
											<input type='hidden' name='txtAppDate' value="<%=tmpAppDate%>">
											<input type='hidden' name='txtAppTFrom' value="<%=tmpAppTFrom%>">
											<input type='hidden' name='selLang' value="<%=tmpLang%>">
											<!--<input class='btnLnk' type='button' name='btnNewIntr' value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeIntr();'>//-->
											<input type="button" value="AVAIL" name="btnchkavail"
											onclick='PopMe2(document.frmAssign.txtAppDate.value, document.frmAssign.txtAppTFrom.value, document.frmAssign.selLang.value);' title='Check Availability' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input class='btnLnk' type='button' name='btnCert' value='CERT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='CertMe(<%=Request("ID")%>);'>
											<input type='hidden' name='chkAvail'>	
											<input type='hidden' name='HnewIntr'>
											<input type='hidden' name='Lang1'>
											<input type='hidden' name='Lang2'>
											<input type='hidden' name='Lang3'>
											<input type='hidden' name='Lang4'>
											<input type='hidden' name='Lang5'>
											<input type='hidden' name='Lang6'>
											<input type='hidden' name='LangCtr'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtIntrLname' value='<%=tmpIntrLname%>' onkeyup='bawal(this);'>
											<input class='trans' size='1' style='width: 5px;' name='txtcoma' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtIntrFname' value='<%=tmpIntrFname%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim2'  readonly size='12'>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='0'  <%=selIntrEmail%> onclick='chkPrim2();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtIntrEmail' value='<%=tmpIntrEmail%>' onkeyup='bawal(this);' readonly></td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='1' <%=selIntrP1%> onclick='chkPrim2();'>&nbsp;
											Home Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrP1' value='<%=tmpIntrP1%>' onkeyup='bawal(this);' readonly>
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtIntrExt' value='<%=tmpIntrExt%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='3' <%=selIntrFax%> onclick='chkPrim2();'>
											&nbsp;&nbsp;&nbsp;Fax:
											</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrFax' value='<%=tmpIntrFax%>' onkeyup='bawal(this);' readonly>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Please include area code on fax number</span>
										</td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='2' <%=selIntrP2%> onclick='chkPrim2();'>
											Mobile Phone:
											</td>
										<td><input class='main' size='12' maxlength='12' name='txtIntrP2' value='<%=tmpIntrP2%>' onkeyup='bawal(this);' readonly></td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddrI' value='<%=tmpNewIntrAddrI%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddr' value='<%=tmpIntrAddr%>' onkeyup='bawal(this);' readonly>
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtIntrCity' value='<%=tmpIntrCity%>' onkeyup='bawal(this);' readonly>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtIntrState' value='<%=tmpIntrState%>' onkeyup='bawal(this);' readonly>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtIntrZip' value='<%=tmpIntrZip%>' onkeyup='bawal(this);' readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>In-House Interpreter:</td>
										<td><input type='checkbox' name='chkInHouse' value='1' <%=tmpInHouse%>></td>
									</tr>
									<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
										<tr>
											<td align='right' width='15%'>Default Rate:</td>
											<td>
												<input type="text" class='main' size='5' maxlength='5'  readonly  name='txtIntrRate' value='<%=tmpIntrRate%>'>
												<select class='seltxt' style='width: 70px;' name='selIntrRate'>
													<option value='0' >&nbsp;</option>
													<%=strRate2%>
												</select>
											</td>
										</tr>
									<% Else %>
										<tr><td>
											<input type="hidden" class='main' size='5' maxlength='5'  readonly  name='txtIntrRate' value='<%=tmpIntrRate%>'>
											<select class='seltxt' style='width: 70px;' name='selIntrRate'>
													<option value='0' >&nbsp;</option>
													<%=strRate2%>
												</select>
										</td></tr>
									<% End If %>
									<% If Request("ID") <> "" Then %>
										<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
											<tr>
												<td align='right' width='15%'>Request Rate:</td>
												<td>
													<input type="text" class='main' size='5' maxlength='5'  readonly  name='txtReqIntrRate' value='<%=tmpIntrRate%>'>
												</td>
											</tr>
										<% Else %>
											<tr>
												
												<td>
													<input type="hidden" class='main' size='5' maxlength='5'  readonly  name='txtReqIntrRate' value='<%=tmpIntrRate%>'>
												</td>
											</tr>
										<% End If %>
									<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>	
									<td align='right' valign='top'>Interpreter Comment:</td>
									<td>
										<textarea name='txtcomintr' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpComintr%></textarea>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right' valign='top'><b>Interpreter Feedback/Evaluation:</b></td>
									<td>
										<textarea name='txtintrfeed' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpintrFeed%></textarea>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
									<td colspan='10' class='header'><nobr>Billing Information</td>
								</tr>
								<tr>
									<td align='right'>Billable Hours:</td>
									<td class='confirm'><%=tmpBilHrs%></td>
								</tr>
								<tr>
									<td align='right'>Actual Time:</td>
									<td class='confirm'><%=tmpActTFrom%> - <%=tmpActTTo%></td>
								</tr>
								<tr>
									<td align='right'>&nbsp;</td>
									<td rowspan='3' valign='top'>
										<table cellSpacing='2' cellPadding='0' border='0'>
											<tr>
												<td align='left'>Bill To Institution </td>
												<td>|</td>
												<td>Pay To Interpreter</td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilTInst%></td>
												<td>|</td>
												<td class='confirm' align='center'><%=tmpBilTIntr%></td>
											</tr>
											<tr>
												<td class='confirm' align='center'><%=tmpBilMInst%> </td>
												<td>|</td>
												<td class='confirm' align='center'> <%=tmpBilMIntr%></td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td align='right'>Travel Time:</td>
								</tr>
								<tr>
									<td align='right'>Mileage:</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Billing Comment:</td>
									<td class='confirm'><%=tmpCombil%></td>
								</tr>	
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Language Bank Notes	
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Notes:</td>
									<td class='confirm'>
										<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;' readonly><%=tmpLBcom%></textarea>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
												<input type='hidden' name="HID" value='<%=Request("ID")%>'>
												<input type='hidden' name="hidInstRate" value='<%=tmpInstRate%>'>
												<input type='hidden' name="hidIntrRate" value='<%=tmpIntrRate%>'>
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveAss(<%=Request("ID")%>, <%=tmpLang%>, <%=Z_CZero(tmpIntr)%>) ;'>
												<input class='btn' type='button' value='Back' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>';">
											</td>
									</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>
<%
tmpMSG = Replace(Session("MSG"), "<br>", "")
If tmpMSG <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
//--></script>
<%
End If
Session("MSG") = ""
%>