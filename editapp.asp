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
Function GetMyStatus(xxx)
	Select Case xxx
		Case 1
			GetMyStatus = "COMPLETED"
		Case 2
			GetMyStatus = "MISSED"
		Case 3
			GetMyStatus = "CANCELED"
		Case 4
			GetMyStatus = "CANCELED (BILLABLE)"
		Case Else
			GetMyStatus = "PENDING"
	End Select
End Function
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", "''")
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
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
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
server.scripttimeout = 360000
tmpPage = "document.frmAssign."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	tmplName = Z_RemoveDlbQuote(rsConfirm("clname")) 
	tmpfName = Z_RemoveDlbQuote(rsConfirm("cfname"))
	chkClient = ""
	If rsConfirm("Client") = True Then chkClient = "checked"
	chkUClientadd = ""
	If  rsConfirm("CliAdd")  = True Then chkUClientadd = "checked"
	tmpAddr = rsConfirm("caddress") 
	tmpCity = rsConfirm("cCity") 
	tmpState = rsConfirm("cstate") 
	tmpZip = rsConfirm("czip")
	tmpCAdrI = rsConfirm("CliAdrI")
	tmpCFon = rsConfirm("Cphone")
	tmpCAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = Z_FormatTime(rsConfirm("appTimeFrom"))
	tmpAppTTo = Z_FormatTime(rsConfirm("appTimeTo"))
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("deptID")
	tmpInstRate = rsConfirm("InstRate")
	tmpJudge = rsConfirm("judge")
	tmpClaim = rsConfirm("claimant")
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	tmpComintr = rsConfirm("IntrComment")
	tmpcombil = rsConfirm("bilComment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpGender	= Z_CZero(rsConfirm("Gender"))
	tmpMale = ""
	tmpFemale = ""
	If tmpGender = 0 Then 
		tmpMale = "SELECTED"
	Else
		tmpFemale = "SELECTED"
	End If
	chkMinor = ""
	If rsConfirm("Child") Then chkMinor = "CHECKED"
	chkcall = ""
	If rsConfirm("courtcall") Then chkcall = "CHECKED"
	chkleave = ""
	If rsConfirm("leavemsg") Then chkleave = "CHECKED"
	chkout = ""
	If rsConfirm("outpatient") Then chkout = "CHECKED"
	chkmed = ""
	If rsConfirm("hasmed") Then chkmed = "CHECKED"
	If Trim(rsConfirm("medicaid")) <> "" Then
			MCNum = rsConfirm("medicaid")
			'radiomed4 = "checked"
		End If
		If Trim(rsConfirm("meridian")) <> "" Then
			MHPnum = rsConfirm("meridian")
			radiomed1 = "checked"
		End If
		If Trim(rsConfirm("nhhealth")) <> "" Then
			NHHFnum = rsConfirm("nhhealth")
			radiomed2 = "checked"
		End If
		If Trim(rsConfirm("wellsense")) <> "" Then
			WSHPnum = rsConfirm("wellsense")
			radiomed3 = "checked"
		End If
	If Trim(MCNum) <> "" And Trim(MHPnum) = "" And Trim(NHHFnum) = "" And _
		Trim(WSHPnum) = "" THen radiomed4 = "checked"
		chkawk = ""
		If rsConfirm("acknowledge") Then chkawk = "Checked"
	chkAppMed = ""
	If rsConfirm("vermed") = 1 Then chkAppMed = "CHECKED"
	chkdisappmed = ""
	If rsConfirm("vermed") = 2 Then chkdisappmed = "CHECKED"
	chkacc = ""
	If rsConfirm("autoacc") Then chkacc = "CHECKED"
	chkcomp = ""
	If rsConfirm("wcomp") Then chkcomp = "CHECKED"
	tmpsecIns = rsConfirm("secins")
	'timestamp on sent/print
	tmpSentReq = "Request email has not been sent to Requesting Person."
	If rsConfirm("SentReq") <> "" Then tmpSentReq = "Request email was last sent to Requesting Person on <b>" & rsConfirm("SentReq") & "</b>."
	tmpSentIntr = "Request email has not been sent to Interpreter."
	If rsConfirm("SentIntr") <> "" Then tmpSentIntr = "Request email was last sent to Interpreter on <b>" & rsConfirm("SentIntr") & "</b>."
	tmpPrint = "Request has not been printed."
	If rsConfirm("Print") <> "" Then tmpPrint = "Request was last printed on<b> " & rsConfirm("Print") & "</b>."
	tmpHPID = Z_CZero(rsConfirm("HPID"))
	tmpLBcom = rsConfirm("LBcomment")
	tmpLBAmount = rsConfirm("PDamount")
	tmpLate = Z_Czero(rsConfirm("Late"))
	tmpLateRes = Z_Czero(rsConfirm("Lateres"))
	uploadfileview = ""
	If rsConfirm("uploadfile") Then uploadfileview = "MERON"
	disApp = ""
	If rsConfirm("approvePDF") Then disApp = "disabled"
	tmpfileuploaded = ""
	If rsConfirm("uploadfile") Then tmpfileuploaded = "*Form 604A already uploaded. Uploading another file will remove the previous uploaded file." 
	disUpload = ""	
	If rsConfirm("approvePDF") Then 
		disUpload = "disabled"	
		tmpfileuploaded = "*Form 604A already approved."
	End If
	train0 = ""
	train1 = ""
	train2 = ""
	train3 = ""
	If rsConfirm("training") = 0 Then train0 = "selected"
	If rsConfirm("training") = 1 Then train1 = "selected"
	If rsConfirm("training") = 2 Then train2 = "selected"
	If rsConfirm("training") = 3 Then train3 = "selected"
	mrrec = rsConfirm("mrrec")
	tmpblock = ""
	If rsConfirm("blocksched") Then tmpblock = "checked"
End If
rsConfirm.Close
Set rsConfirm = Nothing
'RESUPPLY DATA IN ERROR EVENT
If Session("MSG") <> "" And Request("ID") = "" Then
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
	tmpNewInst = Split(Z_DoDecrypt(Request.Cookies("LBINST")), "|")
	tmpNewDept = Split(Z_DoDecrypt(Request.Cookies("LBDEPT")), "|")
	tmpNewReq = Split(Z_DoDecrypt(Request.Cookies("LBREQ")), "|")
	tmpNewIntr = Split(Z_DoDecrypt(Request.Cookies("LBINTR")), "|")
	tmpNewIntrBTN = tmpNewIntr(10)
	If tmpNewReq(6) = "BACK" Then
		tmpNewReqLN = tmpNewReq(0)
		tmpNewReqFN = tmpNewReq(1)
		tmpReqExt = tmpNewReq(8)
		tmpNewReqPhone = tmpNewReq(2)
		tmpNewReqEmail = tmpNewReq(3)
		tmpNewReqFax = tmpNewReq(4)
		tmpNewReqPrim = tmpNewReq(7)
		selRPEmail = ""
		selRPPhone = ""
		selRPFax = ""
		if tmpNewReqPrim = 0 Then selRPEmail = "checked"
		if tmpNewReqPrim = 1 Then selRPPhone = "checked"
		if tmpNewReqPrim = 2 Then selRPFax = "checked"
	Else
		tmpReqP = tmpEntry(1)
	End If
	tmplName = tmpEntry(2)
	tmpfName = tmpEntry(3)
	chkClient = ""
	If tmpEntry(20) <> "" Then chkClient = "checked"
	tmpAddr = tmpEntry(4)
	chkUClientadd = ""
	If tmpEntry(29) <> "" Then chkUClientadd = "checked"
	tmpCity = tmpEntry(5)
	tmpState = tmpEntry(6)
	tmpZip = tmpEntry(7)
	tmpCAdrI = tmpEntry(30)
	tmpDir = tmpEntry(8)
	tmpSC = tmpEntry(9)
	tmpDOB = tmpEntry(10)
	tmpLang = tmpEntry(11)
	tmpAppDate = tmpEntry(12)
	tmpAppTFrom = Z_FormatTime(tmpEntry(13))
	tmpAppTTo = Z_FormatTime(tmpEntry(14))
	tmpAppLoc = tmpEntry(15)
	If tmpNewInst(6) = "BACK" Then 
		tmpNewInstchk = tmpNewInst(6)
		tmpNewInstTxt = tmpNewInst(0)
	Else
		tmpInst = tmpEntry(16)
	End If
	If tmpNewDept(6) = "BACK" Then
		tmpNewInstchk = tmpNewDept(6)
		tmpNewInstDept = tmpNewDept(0)
		SocSer = ""
		Priv = ""
		Legal = ""
		Med = ""
		tmpClass = tmpNewDept(7)
		If tmpNewDept(7) = 1 Then SocSer = "selected"
		If tmpNewDept(7) = 2 Then Priv = "selected"
		If tmpNewDept(7) = 3 Then Court = "selected"
		If tmpNewDept(7) = 4 Then Med = "selected"
		If tmpNewDept(7) = 5 Then Legal = "selected"
		tmpNewInstAddr = tmpNewDept(2)
		tmpNewInstCity = tmpNewDept(3)
		tmpNewInstState = tmpNewDept(4)
		tmpNewInstZip = tmpNewDept(5)
		tmpNewInstAddrI = tmpNewDept(16)
		If tmpNewInst(8) <> "" Then 
			chkBillMe = "checked"
		Else
			tmpBilInstAddr = tmpNewDept(9)
			tmpBilInstCity = tmpNewDept(10)
			tmpBilInstState = tmpNewDept(11)
			tmpBilInstZip = tmpNewDept(12)
		End If	
		tmpBLname =  tmpNewDept(13)
	Else
		tmpDept = tmpEntry(26)
	End If
	tmpRate = tmpEntry(17)
	tmpDoc = tmpEntry(18)
	tmpCRN = tmpEntry(19)
	tmpCFon = tmpEntry(21)
	If Request.Cookies("LBACTION") = 1 Then
		tmpCAFon = tmpEntry(27)
	Else
		tmpCAFon = tmpEntry(38)
	End If
	If tmpNewIntrBTN = "BACK" Then
		tmpIntrLname = tmpNewIntr(0)
		tmpIntrFname = tmpNewIntr(1)
		tmpIntrEmail = tmpNewIntr(2)
		tmpIntrP1 = tmpNewIntr(3)
		tmpIntrExt = tmpNewIntr(13)
		tmpIntrFax = tmpNewIntr(4)
		tmpIntrP2 = tmpNewIntr(5)
		tmpIntrAddr = tmpNewIntr(6)
		tmpIntrCity = tmpNewIntr(7)
		tmpIntrState = tmpNewIntr(8)
		tmpNewIntrAddrI = tmpNewIntr(16)
		tmpIntrZip = tmpNewIntr(9)
		tmpInHouse = ""
		If tmpNewIntr(11) <> "" Then tmpInHouse = "checked"
		tmpIntrPrim = tmpNewIntr(12)
		selIntrFax = ""
		selIntrP2 = ""
		selIntrP1 = ""
		selIntrEmail = ""
		if tmpIntrPrim = 0 Then selIntrEmail = "checked"
		if tmpIntrPrim = 1 Then selIntrP1 = "checked"
		if tmpIntrPrim = 2 Then selIntrP2 = "checked"
		if tmpIntrPrim = 0 Then selIntrFax = "checked"
		tmpIntrRate =  tmpNewIntr(14)
	Else
		tmpIntr = tmpEntry(22)
	End If
	tmpEmer = ""
	If tmpEntry(24) <> "" Then tmpEmer = "checked"
	If Request.Cookies("LBACTION") = 1 Then	
		tmpCom =  tmpEntry(25)
	Else
		tmpCom =  tmpEntry(30)
	End If
End If
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = "(In-House)"
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name") & " " & tmpInHouse
Else
	tmpIntrName = "<i>To be assigned.</i>"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'HP DATA
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		tmpCallMe = ""
		If rsHP("callme") = True Then tmpCallMe = "* Call patient to remind of appointment"
		if rsHP("reason") <> "" Then mytmpReas = Z_Replace(rsHP("reason"),", ", "|")
		tmpReason = GetReas(mytmpReas)
		tmpClin = rsHP("clinician")  
		tmpReqname = rsHP("reqName")  
		InHP = 0
		tmpMeet = ""
		If rsHP("mwhere") = 1 Then
			InHP = 1
			tmpMeet = UCase(GetLoc(rsHP("mlocation")))
			If tmpMeet = "OTHER" Then tmpMeet = rsHP("mother")
		End If
		tmpParents = ""
		If rsHP("parents") <> "" Then tmpParents = rsHP("parents") 
		tmpHPAmount = rsHP("PDamount")
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
	aFon = rsReq("aphone") 
End If
rsReq.Close
Set rsReq = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
	PubDef = 0
	If rsInst("PD") Then PubDef = 1
End If
rsInst.Close
Set rsInst = Nothing 
'GET allowed mco
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM mco_T"
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	Do Until rsInst.EOF 
		If rsInst("mco") = "Medicaid" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmAssign.radiomed[3].disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmAssign.radiomed[3].disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "Meridian" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmAssign.radiomed[0].disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmAssign.radiomed[0].disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "NHhealth" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmAssign.radiomed[1].disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmAssign.radiomed[1].disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "WellSense" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmAssign.radiomed[2].disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmAssign.radiomed[2].disabled = true; " & vbCrLf 
			End If
		End If
		rsInst.MoveNext
	Loop
End If
rsInst.Close
Set rsInst = Nothing
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	mydrg = rsDept("drg")
	myclass = rsDept("class")
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpHistaddr = rsDept("address") & "|" & rsDept("City") & "|" &  rsDept("state") & "|" & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
	
End If
rsDept.Close
Set rsDept = Nothing 
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T WHERE [index] <> 95 ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = -1
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
		"return " & rsLang("index") & ";}"
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'set late
Select Case tmpLate
	Case 3 l3 = "Selected"
	Case 5 l5 = "Selected"
	Case 7 l7 = "Selected"
	Case 10 l10 = "Selected"
	Case 15 l15 = "Selected"
	Case 20 l20 = "Selected"
End Select
'set late

Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM tardy_T ORDER BY UID"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpT = ""
	If tmpLateRes = "" Then tmpTardy = 0
	If CInt(tmpLateRes) = rsLang("UID") Then tmpT = "selected"
	strTar = strTar	& "<option " & tmpT & " value='" & rsLang("UID") & "'>" &  rsLang("lateres") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
Select Case tmpLateRes
	Case 1 lr1 = "Selected"
	Case 2 lr2 = "Selected"
	Case 3 lr3 = "Selected"
	Case 4 lr4 = "Selected"
	Case 5 lr5 = "Selected"
End Select
tmpFilename = Z_GenerateGUID()
Do Until GUIDExists(tmpFilename) = False
	tmpFilename = Z_GenerateGUID()
Loop
If mydrg Then 'secondary insur
	Const adOpenForwardOnly = 0
	Const adOpenKeyset      = 1
	Const adOpenDynamic     = 2
	Const adOpenStatic      = 3
	mySheet = "Alphabetical Order"
	my1stCell = "B3"
	myLastCell = "B900"
	my1stCell2 = "A3"
	myLastCell2 = "A900"
	strHeader = "HDR=NO;"
	myXlsFile = secinsPath & "CARRIER CODE LIST.xls"
	Set objExcel = CreateObject( "ADODB.Connection" )
	 objExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
	    myXlsFile & ";Extended Properties=""Excel 8.0;IMEX=1;" & _
	    strHeader & """"
	 Set objRS = CreateObject( "ADODB.Recordset" )
	    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
	    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
	 Set objRS2 = CreateObject( "ADODB.Recordset" )
	    strRange2 = mySheet & "$" & my1stCell2 & ":" & myLastCell2
	    objRS2.Open "Select * from [" & strRange2 & "]", objExcel, adOpenStatic
	 i = 0
	    Do Until objRS.EOF
				
	      '  If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
	
	        For j = 0 To objRS.Fields.Count - 1
	        		selsecins = ""
	            If Not IsNull( objRS.Fields(j).Value ) Or Trim(objRS.Fields(j).Value) <> "" Then
	           		If tmpsecIns <> "" Then
	           			If tmpsecIns = objRS2.Fields(j).Value then selsecins = "SELECTED"
	           		End If
	            	stroption =stroption & "<option value='" & objRS2.Fields(j).Value & "' " & selsecins & ">" & objRS.Fields(j).Value & "</option>" & vbCrlf
	               'arrData( j, i ) = Trim( objRS.Fields(j).Value )
	            End If
	        Next
	        ' Move to the next row
	        objRS.MoveNext
	        objRS2.MoveNext
	        ' Increment the array "row" number
	        i = i + 1
	    Loop
	 ' Close the file and release the objects
	 	objRS2.Close
    objRS.Close
    objExcel.Close
    Set objRS    = Nothing
    Set objRS2   = Nothing
    Set objExcel = Nothing
End If
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 And mydrg Then
	If tmpHPID > 0 Then hpid = "readonly"
	If tmpHPID > 0 Then hpid2 = "disabled"
End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Edit Appointment - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function Left(str, n){
			if (n <= 0)
			    return "";
			else if (n > String(str).length)
			    return str;
			else
			    return String(str).substring(0,n);
		}
		function CalendarView(strDate)
		{
			document.frmAssign.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAssign.submit();
		}
		function bawal2(tmpform)
		{
			var iChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-,.\'"; //",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		tmp = tmp + tmpform.value.charAt(i);
		  		}
			  	else
		  		{
		  			alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  			
		  		}
		  	}
		}
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
		function RTrim(str)
    {
            var whitespace = new String(" \t\n\r");

            var s = new String(str);

            if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
               

                var i = s.length - 1;       
                while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
                    i--;


              
                s = s.substring(0, i+1);
            }

            return s;
    }
    function LTrim(str)
    {
            var whitespace = new String(" \t\n\r");

            var s = new String(str);

            if (whitespace.indexOf(s.charAt(0)) != -1) {
                
                var j=0, i = s.length;

                while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
                    j++;

                s = s.substring(j, i);
            }

            return s;
    }
    function Trim(str)
    {
            return RTrim(LTrim(str));
    }
    function bawalletters(tmpform) {
			var iChars = "0123456789";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		tmp = tmp + tmpform.value.charAt(i);
		  		}
			  	else
		  		{
		  			alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  			
		  		}
		  	}
		}
		function SaveAss(xxx)
		{
			<% If mydrg Then %>
				if (document.frmAssign.chkmed.checked == true) {
						if (document.frmAssign.txtDOB.value == "") {
							alert("Please input client's date of birth.")
							return;
						}
						if (document.frmAssign.radiomed[0].checked == false && document.frmAssign.radiomed[1].checked == false && document.frmAssign.radiomed[2].checked == false &&
							document.frmAssign.radiomed[3].checked == false) {
							alert("Please select a Medicaid/MCO.")
							return;
						}
						if (Trim(document.frmAssign.MHPnum.value) == "" && document.frmAssign.radiomed[0].checked == true) {
							alert("Please input client's Meridian Health Plan number.")
							return;
						}
						if (Trim(document.frmAssign.NHHFnum.value) == "" && document.frmAssign.radiomed[1].checked == true) {
							alert("Please input client's NH Healthy Families number.")
							return;
						}
						else {
							if (Trim(document.frmAssign.NHHFnum.value) != "") {
								var chrmed = Trim(document.frmAssign.NHHFnum.value);
								if (chrmed.length != 11) {
									alert("Invalid NH Healthy Families number length(11).")
									return;
								}
							}
						}
						if (Trim(document.frmAssign.WSHPnum.value) == "" && document.frmAssign.radiomed[2].checked == true) {
							alert("Please input client's Well Sense Health Plan number.")
							return;
						}
						else {
							if (Trim(document.frmAssign.WSHPnum.value) != "") {
								var chrmed = Trim(document.frmAssign.WSHPnum.value);
								if (chrmed.length != 9) {
									alert("Invalid Well Sense Health Plan number length(9).")
									return;
								}
								var str = Left(document.frmAssign.WSHPnum.value, 2)
								var res = str.toUpperCase(); 
								if (res != 'NH') {
									alert("Well Sense number MUST contain NH (eg: NHXXXXXXX).")
									return;
								}
							}
						}
						if (Trim(document.frmAssign.MCnum.value) == "" && document.frmAssign.radiomed[3].checked == true) {
							alert("Please input client's Medicaid number.")
							return;
						}
						else {
							if (Trim(document.frmAssign.MCnum.value) != "") {
								var chrmed = Trim(document.frmAssign.MCnum.value);
								if (chrmed.length != 11) {
									alert("Invalid Medicaid number length(11).")
									return;
								}
							}
						}
						if (document.frmAssign.chkawk.checked == false) {
							alert("Acknowledge statement is required.")
							return;
						}
					}
			<% End If %>
			if (Trim(document.frmAssign.txtCliAdd.value) != "" || Trim(document.frmAssign.txtCliCity.value) != "" || Trim(document.frmAssign.txtCliState.value) != "" || Trim(document.frmAssign.txtCliZip.value) != "") {
				if (document.frmAssign.chkClientAdd.checked == false) {
					alert("Alternate Appointment Address detected. If you wish to make this address as the appointment address, please check the checkbox beside it.")
					return;
				}
			}
			if (document.frmAssign.chkClientAdd.checked == true)
			{
				if (Trim(document.frmAssign.txtCliAdd.value) == "" || Trim(document.frmAssign.txtCliCity.value) == "" || Trim(document.frmAssign.txtCliState.value) == "" || Trim(document.frmAssign.txtCliZip.value) == "")
				{
					alert("Please input Alternate Appointment's full address.")
					return;
				}
			}
			if (Trim(document.frmAssign.txtAppTFrom.value) == "")
			{
				alert("ERROR: Appointment Time (From:) is Required."); 
				return;
			}
			if (document.frmAssign.txtAppTFrom.value == "24:00")
			{
				alert("ERROR: Appointment Time (From:) is invalid (24:00 not accepted)."); 
				return;
			}
			if (Trim(document.frmAssign.txtAppTTo.value) == "")
			{
				alert("ERROR: Appointment Time (To:) is Required."); 
				return;
			}
			if (document.frmAssign.txtAppTTo.value == "24:00")
			{
				alert("ERROR: Appointment Time (To:) is invalid (24:00 not accepted)."); 
				return;
			}
			if (document.frmAssign.sellate.value > 0 && document.frmAssign.sellateres.value == 0) {
				alert("ERROR: Please select a reason for being tardy."); 
				return;
			}
			<% If PubDef = 1 Then %>
				if (document.frmAssign.txtDocNum.value == "")
				{
					alert("ERROR: Docket Number is Required."); 
					return;
				}
				if (document.frmAssign.txtPDamount.value == "")
				{
					alert("ERROR: Amount requested from court is Required."); 
					return;
				}
			<% End If %>
			if ((document.frmAssign.chkcall.checked == true || document.frmAssign.chkleave.checked == true) && document.frmAssign.txtCliFon.value == "") {
				alert("Please input client's phone number.");
				return;
			}
			if (document.frmAssign.txtCliFon.value != "") {
				document.frmAssign.chkcall.checked = true;
			}
			if (document.frmAssign.myLang.value == document.frmAssign.selLang.value)
			{
				document.frmAssign.action = "action.asp?ctrl=13&ReqID=" + xxx;
				document.frmAssign.submit();
			}
			else
			{
				var ans = window.confirm("Changing Language will result in assigned Interpreter being removed.\nClick Cancel to stop.");
				if (ans)
				{
					document.frmAssign.action = "action.asp?ctrl=13&Intr=1&ReqID=" + xxx;
					document.frmAssign.submit();
				}
			}
		}
		function ViewFile(xxx) {
			newwindow = window.open('f603a.asp?id=' + xxx ,'name','height=750,width=650,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		<% If mydrg Then %>
		function OutPatient() {
				if (document.frmAssign.chkout.checked == true) {
					document.frmAssign.chkmed.disabled = false;
				}
				else {
					document.frmAssign.chkmed.checked = false;
					document.frmAssign.chkmed.disabled = true;
					document.frmAssign.radiomed[3].disabled = true;
					document.frmAssign.radiomed[2].disabled = true;
					document.frmAssign.radiomed[1].disabled = true;
					document.frmAssign.radiomed[0].disabled = true;
					document.frmAssign.radiomed[3].checked = false;
					document.frmAssign.radiomed[2].checked = false;
					document.frmAssign.radiomed[1].checked = false;
					document.frmAssign.radiomed[0].checked = false;
					document.frmAssign.MHPnum.value = "";
					document.frmAssign.NHHFnum.value = "";
					document.frmAssign.WSHPnum.value = "";
					document.frmAssign.MCnum.value = "";
					document.frmAssign.chkawk.disabled = true;
					document.frmAssign.MHPnum.disabled = true;
					document.frmAssign.NHHFnum.disabled = true;
					document.frmAssign.WSHPnum.disabled = true;
					document.frmAssign.MCnum.disabled = true;
				}
			}
		function HasMedicaid(dept) {
			if (document.frmAssign.chkmed.checked == true) {
				document.frmAssign.MCnum.disabled = false;
				<%=allowMCO%>
				document.frmAssign.chkawk.disabled = false;
			}
			else {
				document.frmAssign.radiomed[3].disabled = true;
				document.frmAssign.radiomed[2].disabled = true;
				document.frmAssign.radiomed[1].disabled = true;
				document.frmAssign.radiomed[0].disabled = true;
				document.frmAssign.radiomed[3].checked = false;
				document.frmAssign.radiomed[2].checked = false;
				document.frmAssign.radiomed[1].checked = false;
				document.frmAssign.radiomed[0].checked = false;
				document.frmAssign.MHPnum.value = "";
				document.frmAssign.NHHFnum.value = "";
				document.frmAssign.WSHPnum.value = "";
				document.frmAssign.MCnum.value = "";
				document.frmAssign.chkawk.disabled = true;
				document.frmAssign.MHPnum.disabled = true;
				document.frmAssign.NHHFnum.disabled = true;
				document.frmAssign.WSHPnum.disabled = true;
				document.frmAssign.MCnum.disabled = true;
			}
		}
		<% End If %>
		function uploadFile(xxx)
		{
			var tmpfname = "<%=tmpFilename%>";
			newwindow = window.open('upload.asp?ID=' + xxx + '&hfname=' + tmpfname ,'name','height=150,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		function tardyako() {
			if (document.frmAssign.sellate.value > 0) {
				document.frmAssign.sellateres.disabled = false;
			}
			else {
				document.frmAssign.sellateres.disabled = true;
				document.frmAssign.sellateres.value = 0;
			}
		}
		function PDchk() {
			if (document.frmAssign.h_PD.value == 1) {
				document.frmAssign.btnUp.disabled = false;
			}
			else {
				document.frmAssign.btnUp.disabled = true;
			}				
		}	
		function SelPlan() {
				document.frmAssign.MHPnum.disabled = true;
				document.frmAssign.NHHFnum.disabled = true;
				document.frmAssign.WSHPnum.disabled = true;
				document.frmAssign.MCnum.disabled = true;
				if (document.frmAssign.radiomed[0].checked == true) {
					document.frmAssign.MHPnum.disabled = false;
					document.frmAssign.NHHFnum.value = "";
					document.frmAssign.WSHPnum.value = "";
					//document.frmAssign.MCnum.value = "";
					document.frmAssign.MCnum.disabled = false;
				}
				if (document.frmAssign.radiomed[1].checked == true) {
					document.frmAssign.NHHFnum.disabled = false;
					document.frmAssign.MHPnum.value = "";
					document.frmAssign.WSHPnum.value = "";
					//document.frmAssign.MCnum.value = "";
					document.frmAssign.MCnum.disabled = false;
				}
				if (document.frmAssign.radiomed[2].checked == true) {
					document.frmAssign.WSHPnum.disabled = false;
					document.frmAssign.NHHFnum.value = "";
					document.frmAssign.MHPnum.value = "";
					//document.frmAssign.MCnum.value = "";
					document.frmAssign.MCnum.disabled = false;
				}
				if (document.frmAssign.radiomed[3].checked == true) {
					//document.frmAssign.MCnum.disabled = false;
					document.frmAssign.NHHFnum.value = "";
					document.frmAssign.WSHPnum.value = "";
					document.frmAssign.MHPnum.value = "";
					document.frmAssign.MCnum.disabled = false;
				}
			}
			function Chkdrg(tmpdept) {
				<% If Not myDRG Then %>
					//document.frmAssign.chkmed.checked = false;
					document.frmAssign.MCnum.value = "";
					//document.frmAssign.chkmed.disabled = true;
					document.frmAssign.MCnum.disabled = true;
					document.frmAssign.chkacc.disabled = true;
					document.frmAssign.chkcomp.disabled = true;
					//document.frmAssign.selIns.disabled = true;
					document.frmAssign.chkacc.checked = false;
					document.frmAssign.chkcomp.checked = false;
					//document.frmAssign.btnSec.disabled = true;
					document.frmAssign.selIns.value = "";
					document.frmAssign.chkout.checked = false;
					document.frmAssign.chkout.disabled = true;
						document.frmAssign.chkawk.disabled = true;
				<% Else %>
					document.frmAssign.chkout.disabled = false;
					document.frmAssign.chkmed.disabled = false;
					//document.frmAssign.MCnum.disabled = false;
					document.frmAssign.chkacc.disabled = false;
					document.frmAssign.chkcomp.disabled = false;
						document.frmAssign.chkawk.disabled = false;
					//document.frmAssign.selIns.disabled = false;
					//document.frmAssign.btnSec.disabled = false;
					OutPatient();
					HasMedicaid(tmpdept);
				<% End If %>
			}
			function DpwedeMed() {
				if (document.frmAssign.chkacc.checked == true || document.frmAssign.chkcomp.checked == true) {
					alert("This appointment is not eligible for Medicaid/MCO.");
					document.frmAssign.chkout.checked = false;
					return;
				}
			}
			function DpwedeIba() {
				if (document.frmAssign.chkout.checked == true) {
					if (document.frmAssign.chkacc.checked == true || document.frmAssign.chkcomp.checked == true) {
						alert("This appointment is not eligible for Auto Accident and/or Worker's compensation.");
						document.frmAssign.chkacc.checked = false;
						document.frmAssign.chkcomp.checked = false;
						return;
					}
				}
			}
			function chkleavemsg() {
				if (document.frmAssign.chkcall.checked == false) {
					document.frmAssign.chkleave.checked = false;
				}
			}
		//-->
		</script>
		<body onload="tardyako();PDchk();
			<% If mydrg Then %> 
				 Chkdrg(<%=Z_CZero(tmpdept)%>); SelPlan();
			<% End If %>
			">
			<form method='post' name='frmAssign'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Edit Appointment</td>
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
								<tr>
									<td align='right'>Status:</td>
									<td class='confirm' width='300px'><%=Statko%></td>
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
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpDeptaddr%></td>
								</tr>
								<tr>
									<td align='right'>Billed To:</td>
									<td class='confirm'><%=tmpBContact%></td>
								</tr>
								<tr>
									<td align='right'>Billing Address:</td>
									<td class='confirm'><%=tmpBaddr%></td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<!--<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>//-->
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Requesting Person:</td>
									<td class='confirm'><%=tmpRP%></td>
								</tr>
								<tr>
									<td align='right'>Phone:</td>
									<td class='confirm'><%=fon%></td>
								</tr>
								<tr>
									<td align='right'>Fax:</td>
									<td class='confirm'><%=fax%></td>
								</tr>
								<tr>
									<td align='right'>E-Mail:</td>
									<td class='confirm'><%=email%></td>
								</tr>
								<tr>
									<td align='right'>Alternate Phone:</td>
									<td class='confirm'><%=afon%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>

								<tr>
									<td align='right' valign='top'>&nbsp;</td>
									<td colspan='2'>
										<input type='checkbox' name='chkblock' value='1' <%=tmpblock%>>
											&nbsp;Block Schedule
											&nbsp;&nbsp;
											<input type='checkbox' name='chkClient' value='1' <%=chkClient%>>&nbsp;LSS Client
									</td>
								</tr>
									
								<tr>
										<td align='right'>*Client Last Name:</td>
										<td>
											<input class='main' size='20' maxlength='25' name='txtClilname' <%=hpid%> value="<%=tmplname%>" onkeyup='bawal2(this);'>&nbsp;First Name:
											<input class='main' size='20' maxlength='25' name='txtClifname' <%=hpid%> value="<%=tmpfname%>" onkeyup='bawal2(this);'>
											<% If tmpInst = 479 Then %>
	&nbsp;&nbsp;&nbsp;&nbsp;Training:
	<select name='selTrain' class='seltxt' style='width:150px;'>
		<option value='0' <%=train0%>>Regular</option>
		<option value='1' <%=train1%>>MIT Interpreter 1</option>
		<option value='2' <%=train2%>>Continuing Education Hours</option>
		<option value='3' <%=train3%>>Trainers Hours</option>
	</select>
											<% End If %>
										</td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAddrI' value='<%=tmpCAdrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'><nobr>Alternate Appointment Address:</td>
										<td colspan='3'><nobr>
											<input class='main' size='50' maxlength='50' name='txtCliAdd' value='<%=tmpAddr%>' onkeyup='bawal(this);'>
											<input type='checkbox' name='chkClientAdd' value='1' <%=chkUClientadd%>>CHECK this box and FILL these fields if appointment address is different from department address
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td>
											<input class='main' size='25' maxlength='25' name='txtCliCity' value='<%=tmpCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtCliState' value='<%=tmpState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtCliZip' value='<%=tmpZip%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Client Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkcall' value='1' <%=chkcall%>  onclick='chkleavemsg();'>
											Language Bank Interpreter to provide courtesy reminder call (Please note that this is ONLY courtesy reminder call and patient/client may still not show up to his/her appointment).
											<br><br>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkleave' value='1' <%=chkleave%> onclick='chkleavemsg();'>
											If a patient/client does not answer the phone and his answering machine/voice mail picks up a call or family member answers the phone, can interpreter provide/give full appointment<br>
											info (date, time, location, name of hospital/clinic/department, providers name) on patient/client voice message or give this info to patient/client’s family member?
											<br><br>
										</td>
									</tr>
									<tr>
										<td align='right' valign='top'>Alter. Phone:</td>
										<td align='left'>
											<textarea name='txtAlter' class='main' onkeyup='bawal(this);' ><%=tmpCAFon%></textarea>
										</td>
									</tr>
									<tr>
										<td align='right'>Gender:</td>
										<td>
											<select class='seltxt' name='selGender' style='width: 75px;'>
												<option value='0' <%=tmpMale%>>Male</option>
												<option value='1' <%=tmpfeMale%>>Female</option>
											</select>
											&nbsp;&nbsp;
											Minor:
											<input type='checkbox' name='chkMinor' value='1' <%=chkMinor%>>
										</td>
									</tr>
									<tr>
										<td align='right'>Directions / Landmarks:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliDir' value='<%=tmpDir%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right' valign='top'>Special Circumstances/Precautions:</td>
										<td>
											<textarea name='txtCliCir' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpSC%></textarea>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Precautions (infections, safety, etc.) for this appointment.</span>
										</td>
										<td>&nbsp;</td>
										<%If tmpHPID <> 0 Then%>
											<td><u>HospitalPilot Information:</u></td>
										<%End If%>
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
										<%If tmpHPID <> 0 Then%>
											<td>&nbsp;</td>
											<td rowspan='4' valign='top'>
												<table cellSpacing='2' cellPadding='0'  border='0' style='border:2px solid;'>
													<tr>
														<td align='right'>ID:</td>
														<td>
															<input class='main' size='11' maxlength='10' name='txtHPID' readonly value='<%=Z_ZeroToNull(tmpHPID)%>' onkeyup='bawal(this);'>
															<input type='hidden' name='hideHPID' value='<%=Z_CZero(tmpHPID)%>'>
														</td>
													</tr>
													<tr>
														<td align='right' valign='top'>Requester's Name:</td>
														<td class='confirm'><%=tmpReqname%></td>
													</tr>
													<tr>
														<td align='right' valign='top'><nobr>Reason:</td>
														<td class='confirm'><%=tmpReason%></td>
													</tr>
													<tr>
														<td align='right'><nobr>Clinician:</td>
														<td class='confirm'><%=tmpClin%></td>
													</tr>
													<%If InHP = 1 Then%>
														<tr>
															<td align='right'><nobr>Meeting Place:</td>
															<td class='confirm'><%=tmpMeet%></td>
														</tr>
													<%End If%>
													<%If tmpParents <> "" Then%>
														<tr>
															<td align='right'><nobr>Parents:</td>
															<td class='confirm'><%=tmpParents%></td>
														</tr>
													<%End If%>
													<% If tmpHPAmount > 0 Then %>
														<tr>
															<td align='right'><nobr>Amount requested from court:</td>
															<td class='confirm'>$<%=Z_FormatNumber(tmpHPAmount, 2)%></td>
														</tr>
													<% End If %>
													<tr>
														<td class='confirm' colspan='2'><%=tmpMinor%></td>
													</tr>
													<tr>
														<td class='confirm' colspan='2'><%=tmpCallMe%></td>
													</tr>
												</table>
											</td>
										<%End If%>
									</tr>
									<tr>
											<td align='right'>Patient MR #:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='mrrec' value='<%=mrrec%>' onkeyup='bawal(this);'>
											</td>
										</tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<select class='seltxt' name='selLang'  style='width:100px;' onchange=''>
												<option value='-1'>&nbsp;</option>
												<%=strLang%>
											</select>
											<input type='hidden' name='myLang' value='<%=tmpLang%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Date:</td>
										<td>
											<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
											<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;" <%=hpid2%>
											onclick="showCalendarControl(document.frmAssign.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input type='hidden' name='mydate' value='<%=tmpAppDate%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Time:</td>
										<td>
											&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpAppTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpAppTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
											<input type='hidden' name='mystime' value='<%=tmpAppTFrom%>'>
										</td>
									</tr>
									<!--<tr>
										<td align='right'>Appointment Location:</td>
										<td><input class='main' size='50' maxlength='50' name='txtAppLoc' value='<%=tmpAppLoc%>' onkeyup='bawal(this);'></td>
									</tr>//-->
									<tr><td>&nbsp;</td></tr>
									<tr><td>&nbsp;</td></tr>
									<% If mydrg Then %>
										<tr>
											<td align='right'><b>For Medicaid/MCO billing:</b></td>
											<td><b>(also fill in)</b></td>
										</tr>
										<tr>
											<td align='right'>Auto Accident:</td>
											<td><input type='checkbox' name='chkacc' value='1' <%=chkacc%> onclick="DpwedeIba();"></td>
										</tr>
										<tr>
											<td align='right'>Worker's Compensation:</td>
											<td><input type='checkbox' name='chkcomp' value='1' <%=chkcomp%> onclick="DpwedeIba();"></td>
										</tr>
										<tr>
											<td align='right'>&nbsp;</td>
											<td>
											
												Approve:
												<input type='checkbox' name='chkAppMed' value='1' <%=chkappmed%> onclick='if(this.checked) {document.frmAssign.chkDisAppMed.checked=false;};'>
												Disaaprove:
												<input type='checkbox' name='chkDisAppMed' value='1' <%=chkdisappmed%> onclick='if(this.checked) {document.frmAssign.chkAppMed.checked=false;};'>
											</td>
										</tr>
										<tr>
											<td align='right'>Outpatient:</td>
											<td><input type='checkbox' name='chkout' value='1' <%=chkout%> onclick="DpwedeMed(); OutPatient();"></td>
										</tr>
										<tr>
											<td align='right'>Has Medicaid/MCO:</td>
											<td>
												<input type='checkbox' name='chkmed' value='1' <%=chkmed%> onclick="HasMedicaid(<%=tmpDept%>);">
												Medicaid:<input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>">
											</td>
										</tr>
										<tr>
											<td align='right'></td>
											<td colspan='3'>
												<input type='radio' name='radiomed' <%=radiomed1%> value='1' onclick='SelPlan();'>
												Meridian Health Plan
												<input type='text' class='main' maxlength='14' name='MHPnum' value="<%=MHPnum%>"><br>
												<input type='radio' name='radiomed' <%=radiomed2%> value='2' onclick='SelPlan();'>
												NH Healthy Families
												<input type='text' class='main' maxlength='14' name='NHHFnum' value="<%=NHHFnum%>" onkeyup='bawalletters(this);'><br>
												<input type='radio' name='radiomed' <%=radiomed3%> value='3' onclick='SelPlan();'>
												Well Sense Health Plan
												<input type='text' class='main' maxlength='14' name='WSHPnum' value="<%=WSHPnum%>"><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Well Sense number MUST contain NH (eg: NHXXXXXXX).)</span> <br>
												<input type='radio' name='radiomed' <%=radiomed4%> value='4' onclick='SelPlan();'>
												Medicaid
												<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Directly Billed to Medicaid/Straight Medicaid/Non-MCO)</span> 
												<br><br>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td colspan="3" align="left">
												<input type='checkbox' name='chkawk' value='1' <%=chkawk%> >
												Acknowledgement Statement:<br> On behalf of my organization/institution, I/we agree to accept financial responsibility for this appointment and agree to pay Language Bank for interpretation services provided to us, if MCO or Medicaid declines to pay/cover this appointment.<br><br>
													 I acknowledge that appointment entered is NOT Auto Accident or Workers Compensation case. On behalf of my organization/institution, I/we agree to reimburse/pay Language Bank if the state or MCO request repayment (if case is to be Auto Accident or Workers Compensation case). 
													<br><br>
											</td>
											<!--<td><input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>"></td>//-->
										</tr>
										<tr>
											<td align='right'>Secondary Insurance:</td>
											<td>	
												<select class='seltxt' name='selIns'  style='width:200px;' onchange=''>
													<option value='0'>&nbsp;</option>
													<%=stroption%>
												</select>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
									<% End If %>
									<tr>
										<td align='right'><b>For legal appointments:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Judge:</td>
										<td><input class='main' size='50' maxlength='50' name='txtjudge' value='<%=tmpJudge%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Claimant:</td>
										<td><input class='main' size='50' maxlength='50' name='txtClaim' value='<%=tmpClaim%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<% If tmpInst = 757 Or tmpInst = 777 Then %>
											<td align='right'>Delivery Ticket:</td>
										<% Else %>	
											<td align='right'>Docket Number:</td>
										<% End If %>
										<td><input class='main' size='50' maxlength='50' name='txtDocNum' value='<%=tmpDoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Court Room No:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCrtNum' value='<%=tmpCRN%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'><b>For Public Defender:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Amount requested from court:</td>
										<td>
											$<input class='main' size='8' maxlength='7' name='txtPDamount' value='<%=tmpLBAmount%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>Form 604A:</td>
										<td colspan="3">
											<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile(<%=Request("ID")%>);" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=disUpload%>>
											<!--<input  class='main' type="file" name="F1" size="20" class='btn'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span>//-->
											<input type="hidden" name="h_tmpfilename" value='<%=tmpFilename%>'><%=tmpfileuploaded%>
											<% If uploadfileview <> "" Then %>
												<a href="#" onclick="ViewFile(<%=Request("ID")%>);" style="text-decoration: none;">[view uploaded file]</a>
											<% End If %>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>TARDY:</td>
										<td>
											<select class='seltxt' name='sellate' style='width: 75px;' onchange="tardyako();">
												<option value='0'>0</option>
												<option value='3' <%=l3%>>3</option>
												<option value='5' <%=l5%>>5</option>
												<option value='7' <%=l7%>>7</option>
												<option value='10' <%=l10%>>10</option>
												<option value='15' <%=l15%>>15</option>
												<option value='20' <%=l20%>>>20</option>
											</select> Mins.
											&nbsp;&nbsp;
											Reason: <select class='seltxt' name='sellateres' style='width: 200px;'>
												<option value='0'>&nbsp;</option>
												<!--<option value='1' <%=lr1%>>Previous appointment ran over </option>
												<option value='2' <%=lr2%>>Wrong address on VForm</option>
												<option value='3' <%=lr3%>>Traffic/accident</option>
												<option value='4' <%=lr4%>>Family emergency</option>
												<option value='5' <%=lr5%>>No excuse</option>//-->
												<%=strTar%>
											</select>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Appointment Comment:</td>
										<td colspan='3' >
											<textarea name='txtcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCom%></textarea>
										</td>
									</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information</td>
								</tr>
								<tr>
									<td align='right'>Interpreter:</td>
									<td class='confirm'><%=tmpIntrName%>
										<input type='hidden' name='myint' value='<%=tmpIntr%>'>
										</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Interpreter Comment:</td>
									<td class='confirm'><%=tmpComintr%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
										<tr><td>&nbsp;</td></tr>
									<tr>
									<td colspan='10' class='header'><nobr>Billing Information</td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
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
								<% end if %>
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
												<input type='hidden' name="hidHistaddr" value='<%=tmpHistaddr%>'>
												<input type='hidden' name='h_PD' value='<%=PubDef%>'>
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveAss(<%=Request("ID")%>) ;'>
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
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>