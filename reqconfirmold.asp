<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
If Left(Request.ServerVariables("REMOTE_ADDR"), 10) = "192.168.1." Or _
	Request.ServerVariables("REMOTE_ADDR") = "127.0.0.1" Or _
	Left(Request.ServerVariables("REMOTE_ADDR"), 8) = "10.10.1." Then
	googlekey = "ABQIAAAAd5OxJhCEqwRNwElUvBNmZxR9PeFMte5gUE1Dq7em5JwYVo_dVhScQdsXHPRROmqe71rlFsfMGLuovg"
Else
	googlekey = "ABQIAAAAd5OxJhCEqwRNwElUvBNmZxSZl_t-SL-f-oE8q1L92qagyvYqqhSeaBa4qBIqCn9H6Ik6hSNkS-Lp6w"
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
Function GetPrime(xxx)
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
Function GetPrime2(xxx)
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 1 Or rsRP("prime") = 2 Then
			GetPrime2 = ""
		ElseIf rsRP("prime") = 3 Then
			GetPrime2 = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
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
tmpPage = "document.frmConfirm."
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("reqID") 
	tmpClient = ""
	tmpDeptaddr = ""
	tmpmed = ""
	If rsConfirm("outpatient") And rsConfirm("hasmed") Then
		tmpmed = rsConfirm("medicaid")
	End If
	If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	If rsConfirm("CliAdd") = True Then 
		tmpDeptaddrG = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
		tmpZipInst = rsConfirm("czip")
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
	tmpBilHrs = rsConfirm("Billable")
	tmpActTFrom = Z_FormatTime(rsConfirm("astarttime")) 
	tmpActTTo = Z_FormatTime(rsConfirm("aendtime"))
	tmpBilTInst = rsConfirm("TT_Inst")
	tmpBilTIntr = rsConfirm("TT_Intr")
	tmpBilMInst = rsConfirm("M_Inst")
	tmpBilMIntr = rsConfirm("M_Intr")
	tmpComintr = rsConfirm("intrcomment")
	tmpcombil = rsConfirm("bilcomment")
	tmpLBcom = rsConfirm("LBcomment")
	tmpHPID = Z_CZero(rsConfirm("HPID"))
	chkPaid = ""
	If Not IsNull(rsConfirm("Processed")) Or rsConfirm("Processed") <> "" Then chkPaid = "<i>- BILLED on " & rsConfirm("Processed") & "</i>"
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
	tmpLate = ""
	If Z_Czero(rsConfirm("late")) > 0 Then tmpLate = " --> " & rsConfirm("late") & " MINS. TARDY"
	
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
	tmpZipInst = ""
	tmpClass = rsDept("Class")
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpZipInst = rsDept("zip")
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
	tmpIntrAdd = rsIntr("address1") & ", " & rsIntr("IntrAdrI") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrAddG = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrZip = rsIntr("Zip Code")
	PconIntr = GetPrime2(tmpIntr)
	tmpZipIntr = ""
	If rsIntr("Zip Code") <> "" Then tmpZipIntr = rsIntr("Zip Code")
	TTM = ""
Else
	tmpIntrName = "<i>To be assigned.</i>"
	TTM = "disabled"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'get mileage cap for interpreters
set rsmile = server.createobject("adodb.recordset")
sqlmile = "select * from travel_t"
rsmile.open sqlmile, g_strconn, 3, 1
if not rsmile.eof then
	tmpmilecap = Z_czero(rsmile("milediff"))
end if
rsmile.close
set rsmile = nothing
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
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Request Confirmation - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
			<script src="http://maps.google.com/?file=api&amp;v=2.x&amp;key=<%=googlekey%>"
      type="text/javascript"></script>
		<script language='JavaScript'>
		<!--
		<% If tmpIntr <> 0 Then %>
			var map;
	    var gdir;
	    var geocoder = null;
	    var addressMarker;
	    var duree;
	    var dist;
			var dureeHrs;
			var distMile;
			
	    function initialize() {
	      if (GBrowserIsCompatible()) {      
	        map = new GMap2(document.getElementById("map_canvas"));
	        gdir = new GDirections(map, document.getElementById("directions"));
	        GEvent.addListener(gdir, "load", onGDirectionsLoad);
	        GEvent.addListener(gdir, "error", handleErrors);

					setDirections("<%=tmpIntrAddG%>", "<%=tmpDeptaddrG%>", "en_US");
	      }
	    }
	    function setDirections(fromAddress, toAddress, locale) {
	      gdir.load("from: " + fromAddress + " to: " + toAddress,
	                { "locale": locale });
	     
	    }
	
	    function handleErrors(){
		   if (gdir.getStatus().code == G_GEO_UNKNOWN_ADDRESS)
		     {
		   		var ans = window.confirm("No corresponding geographic location could be found for one of the specified addresses. This may be due to the fact that the address is relatively new, or it may be incorrect.\nError code: " + gdir.getStatus().code + "\n\nDo you want ZIP CODES to be used instead (Directions button will be disabled) ?");
		   		if (ans)
		   		{		
		   				document.frmConfirm.zipcalc.disabled = true;
		   				setDirections("<%=tmpIntrZip%>", "<%=tmpZipInst%>", "en_US");
		   		}
		   		else
	   			{
	   				document.frmConfirm.zipcalc.disabled = false;
	   			}
		   	}
		   else if (gdir.getStatus().code == G_GEO_SERVER_ERROR)
		     alert("A geocoding or directions request could not be successfully processed, yet the exact reason for the failure is not known.\n Error code: " + gdir.getStatus().code);
		   
		   else if (gdir.getStatus().code == G_GEO_MISSING_QUERY)
		     alert("The HTTP q parameter was either missing or had no value. For geocoder requests, this means that an empty address was specified as input. For directions requests, this means that no query was specified in the input.\n Error code: " + gdir.getStatus().code);
	
		     
		   else if (gdir.getStatus().code == G_GEO_BAD_KEY)
		     alert("The given key is either invalid or does not match the domain for which it was given. \n Error code: " + gdir.getStatus().code);
	
		   else if (gdir.getStatus().code == G_GEO_BAD_REQUEST)
		     alert("A directions request could not be successfully parsed.\n Error code: " + gdir.getStatus().code);
		    
		   else alert("An unknown error occurred.");
		   
		}
	
		function onGDirectionsLoad(){ 
	      // Use this function to access information about the latest load()
	      // results.
				duree = gdir.getDuration();
				dist = gdir.getDistance();
				dureeHrs = ((duree.seconds) / 60) / 60;
				distMile = dist.meters / 1609.344;
	   		decHrs = dureeHrs;
				decMile = distMile;
	   		tmpRate = decMile / decHrs;
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
					var ans = window.confirm("This action will send an email/fax to the requesting person.\nClick Cancel to stop.");
					if (ans)
					{
						document.frmConfirm.action = "email.asp?sino=0&emailadd='" + tmpemail + "' &HID=" + <%=Request("ID")%>;
						document.frmConfirm.submit();
					}
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
		function EditMe()
		{
			document.frmConfirm.action = "main.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
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
			<% Else %>
				var ans = window.confirm("Travel Time and Mileage has been saved already. Do you want to save it again?\nClick Cancel to stop.\n\n*Travel Time and Mileage may change from time to time which may cause a different value to be produced when saved again.");
				if (ans)
				{
					document.frmConfirm.action = "action.asp?ctrl=18&ReqID=" + xxx;
				document.frmConfirm.submit();
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
		-->
		</script>
		<% If tmpIntr <> 0 Then %>
			<body onload='PopMe(<%=Request("PID")%>);' onunload="GUnload();">
		<% Else %>
			<body onload='PopMe(<%=Request("PID")%>);'>
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
								<tr>
									<td align='right'>Timestamp:</td>
									<td class='confirm' width='300px'><%=TS%></td>
								</tr>
								<tr>
									<td align='right'>Status:</td>
									<td class='confirm' width='300px'><%=Statko%>&nbsp;&nbsp;<%=chkPaid%></td>
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
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>
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
									<% If tmpmed <> "" Then %>
										<tr>
											<td align='right'>Medicaid Number:</td>
											<td class='confirm'>
												<%=tmpmed%>
											</td>
										</tr>				
									<% End If %>
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
								<tr>
									<td align='right'>Docket Number:</td>
									<td class='confirm'><%=tmpDoc%></td>
								</tr>
								<tr>
									<td align='right'>Court Room No:</td>
									<td class='confirm'><%=tmpCRN%></td>
								</tr>
								<% If tmpHPID <> 0  Then%>
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
										<td align='right'>Clinician:</td>
										<td class='confirm'><%=tmpClin%></td>
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
								<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
									<tr>
										<td align='right'>Rate:</td>
										<td class='confirm'><%=tmpIntrRate%></td>
									</tr>
								<% End If %>
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
								<tr><td>&nbsp;</td></tr>
								
								<tr>
									<td valign="top"><div id="directions" style="display: none;"></div></td>
								</tr>
								<tr>
										<td valign="top"><div id="map_canvas" style="display: none;"></div></td>
								</tr>
								<% If Cint(Request.Cookies("LBUSERTYPE")) = 4 Then %>
										<tr><td>&nbsp;<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
											<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
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
											<input type='hidden' name='txtMile'>
											<input type='hidden' name='txtTravel'>
											<input type='hidden' name='txtMileInst'>
											<input type='hidden' name='txtTravelInst'>
										</td>
									</tr>
								<% Else %>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' style='width: 253px;' <%=TTM%> value='Save Travel Time and Mileage' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveTTM('<%=Request("ID")%>');">
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
											<input class='btn' type='button' style='width: 253px;' value='Send to Requesting Person' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail('<%=Pcon%>');">
											<input class='btn' type='button' style='width: 253px;' value='Send to Interpreter' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="chkEmail2('<%=PconIntr%>', document.frmConfirm.txtMile.value, document.frmConfirm.txtTravel.value);">
											<input type='hidden' name='txtInstZip' value='<%=tmpZipInst%>'>
											<input type='hidden' name='txtIntrZip' value='<%=tmpZipIntr%>'>
											<input type='hidden' name='txtMile'>
											<input type='hidden' name='txtTravel'>
											<input type='hidden' name='txtMileInst'>
											<input type='hidden' name='txtTravelInst'>
										</td>
									</tr>
									<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
										<tr>
											<td colspan='10' align='center' valign='bottom'>
												<input class='btn' type='button' style='width: 510px;' value='Delete Appointment' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe(<%=Request("ID")%>);'>
											</td>	
										</tr>
									<% End If %>	
									<tr><td>&nbsp;</td></tr>
								<% End If %>
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
	initialize();
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