<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetMyStatus(xxx)
	Select Case xxx
		Case 8
			GetMyStatus = "UNFULFILLED"
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
Function GetLoc(xxx)
	Select Case xxx
		Case 0 
			GetLoc = "Front Door"
		Case 1
			GetLoc = "Cafeteria"
		Case 2
			GetLoc = "Registration Desk"
		Case 3
			GetLoc = "Department"
		Case 4
			GetLoc = "OTHER"
	End Select
End Function
Function Z_strLate(intRes)
	Z_strLate = "N/A"
	If Z_CZero(intRes) = 0 Then Exit Function
	Set rsLate = Server.CreateObject("ADODB.RecordSet")
	rsLate.Open "SELECT lateres FROM Tardy_T WHERE UID = " & intRes, g_strCONN, 3, 1
	If Not rsLate.EOF Then
		Z_strLate = Trim(rsLate("lateres"))
	End If
	rsLate.Close
	Set rsLate = Nothing
End Function

TotalEmails = ""
'get mileage rate
set rsmile = server.createobject("adodb.recordset")
sqlmile = "select * from mileageRate_T"
rsmile.open sqlmile, g_strconn, 3, 1
if not rsmile.eof then
	tmpmilerate = Z_czero(rsmile("mileagerate"))
end if
rsmile.close
set rsmile = nothing
tmpPage = "document.frmConfirm."
uploadfileview = ""
chk_noChange = "checked"
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	If Z_DateNull(rsConfirm("med_billInst")) <> Empty Then
		chk_noChange = ""
		chk_billInst = "checked"
		med_billInst = rsConfirm("med_billInst")
	ElseIf  Z_DateNull(rsConfirm("med_billmed")) <> Empty Then
		chk_noChange = ""
		chk_billMed = "checked"
		med_billMed = rsConfirm("med_billMed")
		billmedreas = rsConfirm("billmedreas")
	End If
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	tmpmed = ""
	If rsConfirm("outpatient") And rsConfirm("hasmed") Then
		tmpmed = rsConfirm("medicaid")
		tmpmer = rsConfirm("meridian")
		tmpnh = rsConfirm("nhhealth")
		tmpwell = rsConfirm("wellsense")
	End If
	If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	tmpName = Z_RemoveDlbQuote(rsConfirm("clname")) & ", " & Z_RemoveDlbQuote(rsConfirm("cfname")) & tmpClient
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	If rsConfirm("CliAdd") = True Then 
		tmpDeptaddrG = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
		tmpZipInstg = rsConfirm("czip")
	End If
	tmpFon = rsConfirm("Cphone")
	tmpAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = CTime(rsConfirm("appTimeFrom"))
	tmpAppTTo = CTime(rsConfirm("appTimeTo"))
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpJudge = rsConfirm("judge")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	assigned = ""
	If tmpIntr > 0 Then assigned = "disabled"
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	If rsConfirm("emerFEE") = True Then tmpEmer = "(EMERGENCY - Fee applied)"
	tmpCom = rsConfirm("Comment")
	Statko = GetMyStatus(rsConfirm("Status"))
	tmpstat = rsConfirm("Status")
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = Z_FormatNumber( rsConfirm("TT_Inst"), 2)
	tmpBilTIntr = Z_FormatNumber( Z_CZero(rsConfirm("actTT")) * Z_CZero(rsConfirm("intrrate")) , 2)
	tmpBilMInst = Z_FormatNumber( rsConfirm("M_Inst"), 2)
	tmpBilMIntr = Z_FormatNumber( rsConfirm("actMil") * tmpmilerate, 2)
	tmpComintr = rsConfirm("intrcomment")
	tmpcombil = rsConfirm("bilcomment")
	tmpLBcom = rsConfirm("LBcomment")
	tmpHPID = Z_CZero(rsConfirm("HPID"))
	mrrec = rsConfirm("mrrec")
	cc_email = Z_FixNull( rsConfirm("cc_addr") )
	chkPaid = ""
	HideFin = True
	If Not IsNull(rsConfirm("Processed")) Or rsConfirm("Processed") <> "" Or Not IsNull(rsConfirm("Processedmedicaid")) Or _
		rsConfirm("Processedmedicaid") <> "" Then HideFin = False
	If Not IsNull(rsConfirm("Processed")) Or rsConfirm("Processed") <> "" Then chkPaid = "<i>BILLED to Instituion on " & rsConfirm("Processed") & "</i>"
	If Not isNull(rsConfirm("med_billInst")) Or rsConfirm("med_billInst") <> "" Then chkPaid = chkPaid & "<br><i>Billing change " & rsConfirm("med_billInst") & "</i>"
	If Not isNull(rsConfirm("med_billmed")) Or rsConfirm("med_billmed") <> "" Then chkPaid = chkPaid & "<br><i>Billing change " & rsConfirm("med_billmed") & "</i>"
	If Not IsNull(rsConfirm("Processedmedicaid")) Or rsConfirm("Processedmedicaid") <> "" Then chkPaid = chkPaid & "<br><i>BILLED to Medicaid on " & rsConfirm("Processedmedicaid") & "</i>"
	If Left(chkPaid, 4) = "<br>" Then chkPaid = Mid(chkPaid, 5)
	reqTrail = Trim(rsConfirm("billingTrail"))
	If Left(reqTrail, 4) = "<br>" Then reqTrail = Mid(reqTrail, 5)
	chkVer = ""
	If rsConfirm("verified") = True Then chkVer = "(CONFIRMED)"
	tmpGender	= Z_CZero(rsConfirm("Gender"))
	If tmpGender = 0 Then 
		tmpSex = "MALE"
	Else
		tmpSex = "FEMALE"
	End If
	tmpMinor2 = ""
	If rsConfirm("Child") Then tmpMinor2 = "*MINOR"	
	tmpDecTT = z_fixNull(rsConfirm("actTT"))
	tmpDecMile = z_fixNull(rsConfirm("actMil"))
	tmpSent = Z_FixNull(rsConfirm("SentIntr"))
	tmpSentReq = Z_FixNull(rsConfirm("SentReq"))
	tmpLBAmount = rsConfirm("PDamount")
	If rsConfirm("uploadfile") Then 
		uploadfileviewLB = "<a href='#' onclick='ViewFile(" & Request("ID") & ");' style='text-decoration: none;'>[view uploaded file]</a>"'"*To view uploaded file, open 'Appointment Information'."
	Else
		uploadfileviewLB = "*Form 604A has not been uploaded."
	End If
	tmpsyscom = rsConfirm("syscom")
	tmpTrain = ""
	If rsConfirm("training") Then tmpTrain = " (Training Appointment)"
	lateres = ""
	If Z_CZero(rsConfirm("lateres")) > 0 Then lateres = Z_strLate(rsConfirm("lateres"))
	tmpLate = ""
	If Z_Czero(rsConfirm("late")) > 0 Then tmpLate = " --> " & rsConfirm("late") & " MINS. TARDY (" & lateres & ")"
End If
rsConfirm.Close
Set rsConfirm = Nothing
'Intrpreter login check
If Request.Cookies("LBUSERTYPE") = 2 Then
	If tmpIntr <> Session("UIntr") Then
		Session("MSG") = "You cannot view this Appointment."
		Response.Redirect "calendarview2.asp"
	End If
End If
'HP DATA
If tmpHPID <> 0  THen
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_StrCONNHP, 3, 1
	If Not rsHP.EOF Then
		tmpCallMe = ""
		If rsHP("callme") = True Then tmpCallMe = "* Call patient to remind of appointment"
		'tmpReason = rsHP("reason")
		tmpReason = GetReas(Z_Replace(rsHP("reason"),", ", "|"))
		tmpClin = rsHP("clinician")
		tmpReqname = rsHP("reqName")  
		tmpRPhone = rsHP("rPhone")  
		InHP = 0
		tmpMeet = ""
		If rsHP("mwhere") = 1 Then
			InHP = 1
			tmpMeet = UCase(GetLoc(rsHP("mlocation")))
			If tmpMeet = "OTHER" Then tmpMeet = rsHP("mother")
		End If
		tmpMinor = ""
		If rsHP("minor") = True Then tmpMinor = "* Minor"
		tmpParents = ""
		If rsHP("parents") <> "" Then tmpParents = rsHP("parents") 
		tmpHPCom = rsHP("lbcom")
		tmpoLang = rsHP("oLang")
		If tmpInst = 108 Then tmpDHHS = GetUserID(rsHP("deptID"))
		tmpblock = ""
		if rsHP("block") Then tmpblock = "BLOCK SCHEDULE"
		tmpHPAmount = rsHP("PDamount")
		If rsHp("uploadfile") Then uploadfileview = "<a href='#' onclick='ViewFile(" & Request("ID") & ");' style='text-decoration: none;'>[view uploaded file]</a>"'"*To view uploaded file, open 'Appointment Information'."
	End If
	rsHp.Close
	Set rsHp = Nothing

End If
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & RP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = ""
	If tmpHPID <> 0  THen
		If tmpReqname <> "" Then tmpRP = tmpReqname
	End If
	If tmpRP = "" Then tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(RP)
	TotalEmails = Pcon
	aFon = rsReq("aphone") 
End If
rsReq.Close
Set rsReq = Nothing
If (cc_email <> "") Then
	TotalEmails = TotalEmails & ";" & cc_email
	If (InStr(cc_email, "@")<2) Then TotalEmails = TotalEmails & "@emailfaxservice.com"
End If
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
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpOtherInfo = rsDept("otherInfo")
	tmpZipInst = ""
	tmpClass = rsDept("Class")
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpZipInstg = rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE [index] = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
	if tmpoLang <> "" Then tmpSalita = tmpSalita & " (" & tmpoLang & ")"
End If
rsLang.Close
Set rsLang = Nothing 
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = "(In-House)"
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name") & " " & tmpInHouse
	tmpIntrEmail = rsIntr("E-mail")
	tmpIntrP1 = rsIntr("Phone1")
	If rsIntr("P1Ext") <> "" Then tmpIntrP1 = tmpIntrP1 & " ext. " &  rsIntr("P1Ext")
	tmpIntrP2 = rsIntr("Phone2")
	tmpIntrFax = rsIntr("Fax")
	tmpIntrAdd = rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code") 'rsIntr("address1") & ", " & rsIntr("IntrAdrI") & ", " &
	tmpIntrAddG = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrZip = rsIntr("Zip Code")
	PconIntr = GetPrime2(tmpIntr)
	tmpZipIntr = ""
	If rsIntr("Zip Code") <> "" Then tmpZipIntr = rsIntr("Zip Code")
	intrPID = rsIntr("PID")
	intrWID = rsIntr("WID")	
	intrXID = rsIntr("XID")	
	TTM = ""
Else
	tmpIntrName = "<i>To be assigned.</i>"
	TTM = "disabled"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'get mileage cap for interpreters
Set rsmile = Server.CreateObject("adodb.recordset")
sqlmile = "select * from travel_t"
rsmile.open sqlmile, g_strconn, 3, 1
If Not rsmile.EOF Then
	tmpmilecap = Z_czero(rsmile("milediff"))
End If
rsmile.close
Set rsmile = Nothing
'get mileage cap for institutions except courts
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travelInst_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecapinst = Z_czero(rsmile("milediffinst"))
	end if
	rsmile.close
	set rsmile = nothing
'get mileage cap for institutions courts only
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travelInstCourt_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecapcourts = Z_czero(rsmile("milediffcourt"))
	end if
	rsmile.close
	set rsmile = nothing
'get eval/feedback
strFB = ""
Set rsFB = Server.CreateObject("ADODB.RecordSet")
sqlFB = "SELECT * FROM InterpreterEval_T WHERE appID = " & Request("ID") 
If Request.Cookies("LBUSERTYPE") <> 1 Then
	sqlFB = sqlFB & " AND UID = " & Request.Cookies("UID")
End If
sqlFB = sqlFB & " ORDER BY date DESC"
rsFB.Open sqlFB, g_strCONN, 3, 1
Do Until rsFB.EOF
	strFB = strFB & rsFB("date") & ": " & rsFB("comment") & " - " & GetUsername(rsFB("UID")) & "<br>"
	rsFB.MoveNext
Loop
rsFB.Close
Set rsFB = Nothing
canremove = "Disabled"
If Z_CZero(tmpIntr) > 0 Then canremove = ""
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Request Confirmation - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		<% If tmpIntr <> 0 Then %>
    var origindef = "<%=tmpIntrAddG%>";
				var desinationdef = "<%=tmpDeptaddrG%>";
				var originzip = "<%=tmpIntrZip%>";
				var desinationzip = "<%=tmpZipInstg%>";
				function initMap() {
  			var service = new google.maps.DistanceMatrixService();
				calculateDistances(service, origindef, desinationdef);	
      	}
				function  calculateDistances(directionsService, from, to) {
	        directionsService.getDistanceMatrix({
	          origins: [from],
	          destinations: [to],
	          travelMode: 'DRIVING',
	          unitSystem: google.maps.UnitSystem.METRIC,
	          avoidHighways: false,
	          avoidTolls: false
	        }, callback);
	      }
       function callback(response, status) {
				  var origins = response.originAddresses;
				  var destinations = response.destinationAddresses;
				  if (origins != '' && destinations != '') {
				    for (var i = 0; i < origins.length; i++) {
				      var results = response.rows[i].elements;
				      for (var j = 0; j < results.length; j++) {
				      	var element = results[j];
				        var distance = element.distance.value;
				        var duration = element.duration.value;
				        getDistanceValues(distance, duration);
				      }
				    }
				  }
				  else {
				  	alert('Error: One of the addresses is invalid. System used ZIP CODES to calculate Travel Time and Mileage');
				  	var service = new google.maps.DistanceMatrixService();
				  	calculateDistances(service, originzip, desinationzip);
				  }
				}
		function getDistanceValues(dista, dura){ 
	      // Use this function to access information about the latest load()
	      // results.
				duree = dura;
				dist = dista;
				dureeHrs = ((duree) / 60) / 60;
				distMile = dist / 1609.344;
	   		decHrs = dureeHrs;
				decMile = distMile;
	   		tmpRate = decMile / decHrs;
	   		document.frmConfirm.txtRTravel.value = Math.round((dureeHrs * 2) * 100)/100;
	   		document.frmConfirm.txtRMile.value = Math.round((distMile * 2) * 100)/100; 
	   		
		  	if (decMile > <%=tmpmilecap%>) //interpreter
		  	{
					bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2); //billable mileage (2 way)
					bilTravel = bilMile / tmpRate; //billable travel time (2 way)
				
		  		document.frmConfirm.txtTravel.value = Math.round(bilTravel * 100)/100;
		  		document.frmConfirm.txtMile.value = Math.round(bilMile * 100)/100; 
		 		}
		 		else
		 		{
		 			document.frmConfirm.txtTravel.value = 0
		  		document.frmConfirm.txtMile.value = 0
		 		}
		  	//institution
				<% If tmpClass = 3 Or tmpClass = 5 Then %> //bill institution
					if (decMile > <%=tmpmilecapcourts%>)
					{
	  				bilMileInst = (decMile * 2) - (<%=tmpmilecapcourts%> * 2); //billable mileage (2 way)
	  				document.frmConfirm.txtMileInst.value = Math.round(bilMileInst * 100)/100;
	  			}
	  			else
	  			{
	  				document.frmConfirm.txtMileInst.value = 0;
	  			}
	  		<% Else %>
	  			if (decMile > <%=tmpmilecapinst%>)
					{
	  				bilMileInst = (decMile * 2) - (<%=tmpmilecapinst%> * 2); //billable mileage (2 way)
	  				document.frmConfirm.txtMileInst.value = Math.round(bilMileInst * 100)/100;
	  			}
	  			else
	  			{
	  				document.frmConfirm.txtMileInst.value = 0;
	  			}
	  		<% End If %>
	  		if (document.frmConfirm.txtMileInst.value > 0)
	  		{
	  			bilTravelInst = bilMileInst / tmpRate; //billable travel time (2 way)
	  			document.frmConfirm.txtTravelInst.value = Math.round(bilTravelInst * 100)/100;
				}
				else
				{
					document.frmConfirm.txtMileInst.value = 0;
				}
	   	}
	  <% End If %>
		function chkEmail(tmpemail)
		{
			if (tmpemail == undefined || tmpemail == "")
				{
					alert("ERROR: Primary Contact is blank or invalid.");
				}
			else
				{
					<% If tmpSentReq = "" Then %>	
						var ans = window.confirm("This action will send an email/fax to the requesting person.\nClick Cancel to stop.");
						if (ans)
						{
							document.frmConfirm.action = "email.asp?sino=0&emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
							document.frmConfirm.submit();
						}
					<% Else %>
						var ans = window.confirm("This Request has already been sent to the Requesting person.\nPlease double check this request to avoid error.\nThis action will send an email/fax to the Requesting person.\nClick Cancel to stop.");
						if (ans)
						{
							document.frmConfirm.action = "email.asp?sino=0&emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
							document.frmConfirm.submit();
						}
					<% End If %>
				}
		}
		function chkEmail2(tmpemail, tmpM, tmpTT)
		{
			if (tmpemail == undefined || tmpemail == "")
				{
					alert("ERROR: Primary Contact is blank or invalid.");
				}
			else
				{
					<% If tmpDecTT <> "" Or tmpDecMile <> "" Then %>
						var MileTT = tmpM + "|" + tmpTT;
						<% If tmpSent = "" Then %>	
							var ans = window.confirm("This action will send an email/fax to the interpreter.\nClick Cancel to stop.");
							if (ans)
							{
								document.frmConfirm.action = "email.asp?sino=1&MileTT='" + MileTT + "' &emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
								document.frmConfirm.submit();
							}
						<% Else %>
							var ans = window.confirm("This Request has already been sent to an Interpreter.\nPlease double check this request to avoid error.\nThis action will send an email/fax to the interpreter.\nClick Cancel to stop.");
							if (ans)
							{
								document.frmConfirm.action = "email.asp?sino=1&MileTT='" + MileTT + "' &emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
								document.frmConfirm.submit();
							}
						<% End If %>
					<% Else %>
						alert("Please save Travel Time and Mileage First.");
					<% End If %>
				}
		}
		function EditMe2()
		{
			<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
				document.frmConfirm.action = "mainbill.asp?ID=" + <%=Request("ID")%>;
			<% Else %>
				document.frmConfirm.action = "mainbill2.asp?ID=" + <%=Request("ID")%>;
			<% End If %>
			document.frmConfirm.submit();
		}
		function EditMe3()
		{
			document.frmConfirm.action = "LBcom.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function AssignMe()
		{
			document.frmConfirm.action = "assign.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function EditContact()
		{
			document.frmConfirm.action = "editcontact.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function EditApp()
		{
			document.frmConfirm.action = "editapp.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function PopMe(zzz)
		{
			if (zzz !== undefined)
				{
				newwindow = window.open('print.asp?ID=' + zzz,'name','height=1056,width=816,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
				}
		}
		function CalendarView(strDate)
		{
			document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmConfirm.submit();
		}
		function CopyMe()
		{
			var ans = window.confirm("Clone this appointment?\nClick Cancel to stop.");
			if (ans)
			{
				document.frmConfirm.action = "clone.asp?Clone=" + <%=Request("ID")%>;
				document.frmConfirm.submit();
			}
		}
		function PopMe2(xxx)
		{
			//if (instzip == "" || intrzip == "")
			//{
			//	alert("Error: Institution's zip code and/or Interpreter's zip code is blank.")
			//	return;
			//}
			//else
			//{
			//	var zip1 = instzip; 
			//	var zip2 = intrzip;
			//	var zipus = zip1 + "|" + zip2
				//alert(zipus);
				newwindow2 = window.open('travel.asp?ReqID=' + xxx,'name','height=600,width=650,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow2.focus()}
			//}
		}
		function MyHist(xxx)
		{
			newwindow3 = window.open('history.asp?ReqID=' + xxx,'name','height=500,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow3.focus()}
		}
		function ConfirmMe(xxx)
		{
			var ans = window.confirm("Confirm appointment?\nClick Cancel to stop.");
			if (ans)
			{
				document.frmConfirm.action = "action.asp?ctrl=14&ReqID=" + xxx;
				document.frmConfirm.submit();
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
		function SaveTTM(xxx)
		{
			<% If tmpDecTT = "" Or tmpDecMile = "" Then %>
				document.frmConfirm.action = "action.asp?ctrl=18&ReqID=" + xxx;
				document.frmConfirm.submit();
				//alert('actual travel: ' + document.getElementById('txtRTravel').value + '\nactual mileage: ' + document.getElementById('txtRMile').value);
			<% Else %>
				var ans = window.confirm("Travel Time and Mileage has been saved already. Do you want to save it again?\nClick Cancel to stop.\n\n*Travel Time and Mileage may change from time to time which may cause a different value to be produced when saved again.\n*Monetary values to be paid for travel time and mileage will be reset to zero(0).");
				if (ans)
				{
					document.frmConfirm.action = "action.asp?ctrl=18&ReqID=" + xxx;
					document.frmConfirm.submit();
					//alert('actual travel: ' + document.getElementById('txtRTravel').value + '\nactual mileage: ' + document.getElementById('txtRMile').value);
				}
			<% End If %>
		}
		function AssignMe2(xxx)
		{
			newwindow = window.open('emailIntr.asp?ID=' + xxx,'','height=250,width=500,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		function KillMe(xxx)
		{
			var ans = window.confirm("Delete Request? Click Cancel to stop.");
			if (ans){
				document.frmConfirm.action = "action.asp?ctrl=9&ReqID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		function ViewFile(xxx) {
			newwindow = window.open('f603a.asp?id=' + xxx ,'name','height=750,width=650,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
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
		function enablebutton() {
			document.frmConfirm.btnSaveTTM.disabled = false;
		}
		<% If Not HideFin Then %>
		function financeonly() {
			if (document.frmConfirm.radioFin[0].checked == true) {
				document.frmConfirm.btnsave.disabled = true;
				document.frmConfirm.med_billInst.disabled = true;
				document.frmConfirm.med_billMed.disabled = true;
				document.frmConfirm.billmedreas.disabled = true;
				document.frmConfirm.med_billInst.value = '';
				document.frmConfirm.med_billMed.value = '';
				document.frmConfirm.billmedreas.value = '';
			}
			else if (document.frmConfirm.radioFin[1].checked == true) {
				document.frmConfirm.btnsave.disabled = false;
				document.frmConfirm.med_billInst.disabled = false;
				document.frmConfirm.med_billMed.disabled = true;
				document.frmConfirm.billmedreas.disabled = true;
				//document.frmConfirm.med_billInst.value = '';
				document.frmConfirm.med_billMed.value = '';
				document.frmConfirm.billmedreas.value = '';
			}
			else if (document.frmConfirm.radioFin[2].checked == true) {
				document.frmConfirm.btnsave.disabled = false;
				document.frmConfirm.med_billInst.disabled = true;
				document.frmConfirm.med_billMed.disabled = false;
				document.frmConfirm.billmedreas.disabled = false;
				document.frmConfirm.med_billInst.value = '';
				//document.frmConfirm.med_billMed.value = '';
				//document.frmConfirm.billmedreas.value = '';
			}
		}
		function changebill(xxx) {
			if (document.frmConfirm.radioFin[1].checked == true) {
				if (Trim(document.frmConfirm.med_billInst.value) == '') {
					alert("Date of Medicaid Claim Denied is required.")
					return;
				}
				else {
					if (isDate(document.frmConfirm.med_billInst.value) == false) {
						alert("Invalid Date of Medicaid Claim Denied.")
						return;
					}
				}
			}
			if (document.frmConfirm.radioFin[2].checked == true) {
				if (Trim(document.frmConfirm.med_billMed.value) == '') {
					alert("Date of Institution Billing change to Medicaid is required.")
					return;
				}
				else {
					if (isDate(document.frmConfirm.med_billMed.value) == false) {
						alert("Invalid Institution Billing change to Medicaid.")
						return;
					}
				}
				if (Trim(document.frmConfirm.billmedreas.value) == '') {
					alert("Reason for Institution Billing change to Medicaid is required.")
					return;
				}
			}
			document.frmConfirm.action = "action.asp?ctrl=24&ReqID=" + xxx;
			document.frmConfirm.submit();
		}
		
	<% Else %>
		function financeonly() {
			
		}
	<% end If %>
		function RTrim(str) {
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
	function LTrim(str) {
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
	function Trim(str) {
		return RTrim(LTrim(str));
	}
	function isDate(dateStr) {
		var datePat = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
		var matchArray = dateStr.match(datePat); // is the format ok?
		if (matchArray == null) {
			alert("Please enter date as mm/dd/yyyy");
			return false;
		}
		month = matchArray[1]; // p@rse date into variables
		day = matchArray[3];
		year = matchArray[5];
		if (month < 1 || month > 12) { // check month range
			alert("Month must be between 1 and 12.");
			return false;
		}
		if (day < 1 || day > 31) {
			alert("Day must be between 1 and 31.");
			return false;
		}
		if ((month==4 || month==6 || month==9 || month==11) && day==31) {
			alert("Month "+month+" doesn`t have 31 days!")
			return false;
		}
		if (month == 2) { // check for february 29th
			var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
			if (day > 29 || (day==29 && !isleap)) {
				return false;
			}
		}
		return true; // date is valid
		}
		function ViewUploads(xxx) {
			newwindow = window.open('viewuploads.asp?ftype=0&reqid=' + xxx,'name','height=750,width=1000,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			if (window.focus) {newwindow.focus()}
		}
		function removeme(xxx) {
			var ans = window.confirm("Remove Interpreter? Click Cancel to stop.");
			if (ans){
				document.frmConfirm.action = "action.asp?ctrl=27&ReqID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		function CancelMe(xxx, yyy)
		{
			if (yyy == 3)
			{
				alert("This request has been canceled already.")
				return;
			}
			var ans = window.confirm("Cancel Appointment.\nAn E-mail will be sent to a LB staff for notification.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmConfirm.action = "action.asp?ctrl=28&ID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		-->
		</script>
		<% If tmpIntr <> 0 Then %>
			<script src="https://maps.googleapis.com/maps/api/js?key=<%=googlemapskey%>" type="text/javascript"></script>
		<% End If %>
		<% If tmpIntr > 0 Then %>
			<body onload='PopMe(<%=Request("PID")%>); financeonly();initMap();enablebutton();'>
		<% Else %>
			<body onload='PopMe(<%=Request("PID")%>); financeonly();'>
		<% End If %>
			<form method='post' name='frmConfirm'>
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
									<td class='title' colspan='10' align='center'><nobr>Request Confirmation</td>
								</tr>
								<tr>
									<td align='center' colspan='10' class='RemME'>
										<a style='text-decoration: none;' href="JavaScript: MyHist(<%=Request("ID")%>);">[History]</a>
										
									</td>
								</tr>
								<tr>
									<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<% If Cint(Request.Cookies("LBUSERTYPE")) = 4 Then %>
										<tr><td>&nbsp;<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
											<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
											<input type="hidden" name="txtRTravel" id="txtRTravel" >
											<input type="hidden" name="txtRMile" id="txtRMile">
											<input type='hidden' name='txtMile'>
											<input type='hidden' name='txtTravel'>
											<input type='hidden' name='txtMileInst'>
											<input type='hidden' name='txtTravelInst'></td></tr>
										<tr><td>&nbsp;</td></tr>
								<% ElseIf Cint(Request.Cookies("LBUSERTYPE")) = 2 Then%>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
											<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
												<input type="hidden" name="txtRTravel" id="txtRTravel" >
											<input type="hidden" name="txtRMile" id="txtRMile">
											<input type='hidden' name='txtMile'>
											<input type='hidden' name='txtTravel'>
											<input type='hidden' name='txtMileInst'>
											<input type='hidden' name='txtTravelInst'>
										</td>
									</tr>
								<% Else %>
									<tr>
										<td colspan='10' align='center' height='30px' valign='bottom'>
											<input class='btn' type='button' id='btnSaveTTM' name='btnSaveTTM' disabled='disabled' style='width: 253px;' <%=TTM%> value='Save Travel Time and Mileage' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveTTM('<%=Request("ID")%>');">
											<script>
												document.frmConfirm.btnSaveTTM.disabled = true;
											</script>
											<input class='btn' type='button' style='width: 253px;' <%=assigned%> value='Email' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="AssignMe2('<%=Request("ID")%>');">
										</td>
									</tr>
									<tr>
										<td colspan='10' align='center' valign='bottom'>
											<input class='btn' type='button' style='width: 125px;' value='View in Calendar' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='calendarview2.asp?appdate=<%=tmpAppDate%>'">
											<input class='btn' type='button' style='width: 125px;' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PopMe(<%=Request("ID")%>);'>
											<input class='btn' type='button' style='width: 125px;' value='Clone Appt.' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  onclick='CopyMe();'>
											<input class='btn' type='button' style='width: 125px;' value="Directions" name="zipcalc"
												onclick='PopMe2(<%=Request("ID")%>);' title='Zip code calculator' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
										</td>
									</tr>
									<tr>
										<td colspan='10' align='center' valign='bottom'>
											<input class='btn' type='button' style='width: 253px;' value='Send to Requesting Person' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail('<%=TotalEmails%>');">
											<input class='btn' type='button' style='width: 253px;' value='Send to Interpreter' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail2('<%=PconIntr%>', document.frmConfirm.txtMile.value, document.frmConfirm.txtTravel.value);">
											<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
											<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
											<input type="hidden" name="txtRTravel" id="txtRTravel" >
											<input type="hidden" name="txtRMile" id="txtRMile">
											<input type='hidden' name='txtMile'>
											<input type='hidden' name='txtTravel'>
											<input type='hidden' name='txtMileInst'>
											<input type='hidden' name='txtTravelInst'>
										</td>
									</tr>
									<tr>
										<td colspan='10' align='center' valign='bottom'>
											<input class='btn' type='button' style='width: 510px;' value='Cancel Appointment' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='CancelMe(<%=Request("ID")%>, <%=tmpstat%>);'>
										</td>	
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' valign='bottom'>
											<input class='btn' type='button' style='width: 510px;' value='Remove Interpreter' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=canremove%> onclick='removeme(<%=Request("ID")%>);'>
										</td>	
									</tr>
									<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
										<tr>
											<td colspan='10' align='center' valign='bottom'>
												<input class='btn' type='button' style='width: 510px;' value='View Uploads' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ViewUploads(<%=Request("ID")%>);'>
											</td>	
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan='10' align='center' valign='bottom'>
												<input class='btn' type='button' style='width: 510px;' value='Delete Appointment' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe(<%=Request("ID")%>);'>
											</td>	
										</tr>
									<% End If %>	
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Language Bank Notes
									<% If Cint(Request.Cookies("LBUSERTYPE")) = 0 Or Cint(Request.Cookies("LBUSERTYPE")) = 1 _
										Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Then %>
									<input class='btnLnk' type='button' name='btnEditNotes' value='EDIT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
										<%=disableMe%> onclick='EditMe3();' title='Language Bank Notes'>		
									</td>
								<% End If %>
								</tr>
								<tr>
									<td align='right' valign='top'>Notes:</td>
									<td class='confirm'>
										<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;' rows='6' readonly><%=tmpLBcom%></textarea>
									</td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information
									<% If Cint(Request.Cookies("LBUSERTYPE")) = 0 Or Cint(Request.Cookies("LBUSERTYPE")) = 1 _
										Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Then %>
										<input class='btnLnk' type='button' name='btnEditContact' value='EDIT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
												<%=disableMe%> onclick='EditContact();' title='Edit Contact Information'>
									<% End If %>
									</td>
								</tr>
								<tr>
									<td align='right'>Request ID:</td>
									<td class='confirm' width='300px'><%=Request("ID")%>&nbsp;<%=tmpEmer%>&nbsp;<%=chkVer%></td>
									<input type='hidden' name='HID' value='<%=Request("ID")%>'>
								</tr>
								<% If tmpHPID > 0 Then %>
								<tr>
									<td align='right'>Vendor Site ID:</td>
									<td class='confirm' width='300px'><%=tmpHPID%></td>
								</tr>	
								<% End If %>
								<tr>
									<td align='right'>Timestamp:</td>
									<td class='confirm' width='300px'><%=TS%></td>
								</tr>
								<tr>
									<td align='right' valign="top">Status:</td>
									<td class='confirm' width='300px'><%=Statko%><br><br>--- Current Status ---<br><%=chkPaid%><br><br>--- Trail ---<br><i><%=reqTrail%></i><br></td>
								</tr>
								<% If Request.Cookies("UID") = 8 Or Request.Cookies("UID") = 2 And HideFin = False Then %>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3">
											<table style="border: 1px solid;">
												<tr>
													<td>
														<input type='radio' name='radioFin' value='0' <%=chk_noChange%> onclick='financeonly();'>
													</td>
													<td>
														<b>No Change</b>
													</td>
												</tr>
												<tr>
													<td>
														<input type='radio' name='radioFin' value='1' <%=chk_billInst%> onclick='financeonly();'>
													</td>
													<td>
														<b>Medicaid Claim Denied</b>
													</td>
												</tr>
												<tr>
													<td>&nbsp;</td>
													<td>Date:<input class='main' size='11' maxlength='10' name='med_billInst' value='<%=med_billInst%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
												</tr>
												<tr>
													<td>
														<input type='radio' name='radioFin' value='2' <%=chk_billMed%> onclick='financeonly();'>
													</td>
													<td>
														<b>Institution Billing change to Medicaid</b>
													</td>
												</tr>
												<tr>
													<td>&nbsp;</td>
													<td>Date:<input class='main' size='11' maxlength='10' name='med_billMed' value='<%=med_billMed%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
												</tr>
												<tr>
													<td>&nbsp;</td>
													<td>Reason:<input class='main' size='50' maxlength='50' name='billmedreas' value='<%=billmedreas%>' onkeyup='bawal(this);'>
												</tr>
												<tr><td>&nbsp;</td></tr>
												<tr>
													<td align="center" colspan="2">
														<input class='btn' type='button' name='btnsave' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="changebill(<%=Request("ID")%>);">
													</td>
												</tr>
											</table>
										</td>
									</tr>
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%> <%=tmpTrain%></td>
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
								<tr>
										<td align='right' width='15%'>Other Info:</td>
										<td class='confirm'><%=tmpOtherInfo%></td>
								</tr>
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
<%	If (cc_email <> "") Then %>
								<tr><td align='right'>CC:</td>
									<td class='confirm'><%=cc_email%></td>
								</tr>
<%	End If %>
								<tr>
									<td align='right'>Alternate Phone:</td>
									<td class='confirm'><%=aFon%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>DHHS ID:</td>
									<td class='confirm'><%=tmpDHHS%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information
									<% If Cint(Request.Cookies("LBUSERTYPE")) = 0 Or Cint(Request.Cookies("LBUSERTYPE")) = 1 _
										Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Then %>
										<input class='btnLnk' type='button' name='btnEditApp' value='EDIT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
												<%=disableMe%> onclick='EditApp();' title='Edit Appointment Information'>
									<% End If %>
									</td>
								</tr>
								<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<tr>
										<td align='right'>Client Name:</td>
										<td class='confirm'>
											<%=tmpName%>
											<% If tmpHPID <> 0  Then%>
												&nbsp;<%=tmpMinor%>
											<% End If%>
										</td>
									</tr>
								<% End If %>
								<tr>
									<td align='right'>Client Address:</td>
									<td class='confirm'><%=tmpAddr%></td>
								</tr>
								<tr>
									<td align='right'>Client Phone:</td>
									<td class='confirm'><%=tmpFon%></td>
								</tr>
								<tr>
									<td align='right'>Client Alter. Phone:</td>
									<td class='confirm'><%=tmpAFon%></td>
								</tr>
								<tr>
									<td align='right'>Gender:</td>
									<td class='confirm'><%=tmpSex%>&nbsp;<%=tmpMinor2%></td>
								</tr>
								<tr>
									<td align='right'>Patient MR #:</td>
									<td class='confirm'><%=mrrec%></td>
								</tr>
								<tr>
									<td align='right'>Directions / Landmarks:</td>
									<td class='confirm'><%=tmpdir%></td>
								</tr>
								<tr>
									<td align='right'>Special Circumstances:</td>
									<td class='confirm'><%=tmpSC%></td>
								</tr>
								<tr>
									<td align='right'>DOB:</td>
									<td class='confirm'><%=tmpDOB%></td>
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
									<td align='right'>Appointment Location:</td>
									<td class='confirm'><%=tmpAppLoc%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'><b>For Medicaid/MCO Billing:</b></td>
									<td align='left'><b>----</b></td>
								</tr>
								<tr>
									<td align='right'>Medicaid Number:</td>
									<td class='confirm'><%=tmpMed%></td>
								</tr>
								<tr>
									<td align='right'>Meridian Number:</td>
									<td class='confirm'><%=tmpMer%></td>
								</tr>
								<tr>
									<td align='right'>NH Health Number:</td>
									<td class='confirm'><%=tmpnh%></td>
								</tr>
								<tr>
									<td align='right'>Well Sense Number:</td>
									<td class='confirm'><%=tmpwell%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'><b>For legal appointments:</b></td>
									<td align='left'><b>----</b></td>
								</tr>
								<tr>
									<td align='right'>Judge:</td>
									<td class='confirm'><%=tmpJudge%></td>
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
									<td align='right'><b>For Public Defender:</b></td>
									<td align='left'><b>----</b></td>
								</tr>
								<tr>
									<td align='right'>Amount requested from court:</td>
									<td class='confirm'>$<%=Z_FormatNumber(tmpLBAmount, 2)%></td>
								</tr>
								<% If PubDef = 1 Then %>
									<tr>
										<td align='right'>&nbsp;</td>
										<td class='confirm'><%=uploadfileviewLB%></td>
									</tr>
								<% End If %>
								<% If tmpHPID <> 0  Then%>
									<tr><td>&nbsp;</td></tr>
									<tr><td align='right'>Hospital Pilot Information</td><td>---</td></tr>
									<tr>
										<td align='right' valign='top'>Requester's Name:</td>
										<td class='confirm'><%=tmpReqname%></td>
									</tr>
									<tr>
										<td align='right' valign='top'>Requester's Phone:</td>
										<td class='confirm'><%=tmpRPhone%></td>
									</tr>
									<tr>
										<td align='right' valign='top'>Reason:</td>
										<td class='confirm'><%=tmpReason%></td>
									</tr>
									<tr>
										<td align='right'>Clinician/Docket Number:</td>
										<td class='confirm'><%=tmpClin%></td>
									</tr>
									<tr>
										<td align='right'>Amount requested from court:</td>
										<td class='confirm'>$<%=Z_FormatNumber(tmpHPAmount, 2)%></td>
									</tr>
									<% If tmpParents <> "" Then%>
										<tr>
											<td align='right'>Parents:</td>
											<td class='confirm'><%=tmpParents%></td>
										</tr>
									<%End If%>
									<tr>
										<td align='right'>HospitalPilot Comment:</td>
										<td class='confirm'><%=tmpHPCom%></td>
									</tr>
									<% If tmpBlock <> "" Then%>
										<tr>
											<td align='right'>&nbsp;</td>
											<td class='confirm'><%=tmpBlock%></td>
										</tr>
									<%End If%>
									<tr>
										<td align='right'>&nbsp;</td>
										<td class='confirm'><%=uploadfileview%></td>
									</tr>
								<%End If%>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Appointment Comment:</td>
									<td class='confirm'><%=tmpCom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information
									<% If Cint(Request.Cookies("LBUSERTYPE")) = 0 Or Cint(Request.Cookies("LBUSERTYPE")) = 1 _
										Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Then %>
										<input class='btnLnk' type='button' name='btnEditIntr' value='EDIT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
												<%=disableMe%> onclick='AssignMe();' title='Edit Interpreter Information'>	
									<% End If %>	
								</td>
								</tr>
								<tr>
									<td align='right'>Interpreter:</td>
									<td class='confirm'>
										<%=tmpIntrName%>
										<% If tmpLate <> "" Then %>
											<%=tmpLate %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td align='right' width='15%'>E-Mail:</td>
									<td class='confirm'><%=tmpIntrEmail%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Home Phone:</td>
									<td class='confirm'><%=tmpIntrP1%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Mobile Phone:</td>
									<td class='confirm'><%=tmpIntrP2%></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Fax:</td>
									<td class='confirm'><%=tmpIntrFax%></td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpIntrAdd%></td>
								</tr>
								<tr>
									<td align='right'>HP Provider ID:</td>
									<td class='confirm'><%=intrPID%></td>
								</tr>
								<tr>
									<td align='right'>Xerox Provider ID:</td>
									<td class='confirm'><%=intrXID%></td>
								</tr>
								<tr>
									<td align='right'>Worker ID:</td>
									<td class='confirm'><%=intrWID%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Interpreter Comment:</td>
									<td class='confirm'><%=tmpComintr%></td>
								</tr>
								<tr>
									<td align='right' valign='top'>Interpreter Feedback/Evaluation:</td>
									<td class='confirm'><%=strFB%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
									<tr>
										<td colspan='10' class='header'><nobr>Billing Information

										<input class='btnLnk' type='button' name='btnEditOther' value='EDIT' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											<%=disableMe%> onclick='EditMe2();' title='Edit Billing Information'>		
										</td>
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
									<% End If %>
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Billing Comment:</td>
									<td class='confirm'><%=tmpCombil%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>System Message:</td>
									<td class='confirm'><%=tmpsyscom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								
								<tr>
									<td valign="top"><div id="output" style="display: none;"></div></td>
								</tr>
								<tr>
										<td valign="top"><div id="map" style="display: none;"></div></td>
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
<% If tmpIntr <> 0 Then %>
<script language='JavaScript'><!--

-->
</script>
<% End If %>
<%
If Session("MSG") <> "" Then
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>