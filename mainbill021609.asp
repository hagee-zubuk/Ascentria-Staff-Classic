<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LB-USERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE index = " & xxx
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
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE index = " & xxx
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
Function GetMyClass(xxx)
	GetMyClass = 1
	Set rsMyClass = Server.CreateObject("ADODB.RecordSet")
	sqlmyclass = "SELECT * FROM dept_T where index = " & xxx
	rsMyClass.Open sqlmyclass, g_strCONN, 3, 1
	If Not rsMyClass.EOF Then
		GetMyClass = rsMyClass("class")
	End If
	rsMyClass.Close
	Set rsMyClass = Nothing
End Function
tmpPage = "document.frmConfirm."
on error resume next
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE index = " & Request("ID")
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
	billedna = ""
	'If tmpStat = 1 or tmpStat = 4 Then billedna = "disabled"
	If rsConfirm("Processed") <> Empty Or rsConfirm("ProcessedPR") <> Empty Then billedna = "disabled"
	tmpMiss = rsConfirm("Missed")
	tmpCancel = rsConfirm("Cancel")
	tmpClient = ""
	If rsConfirm("client") = True Then tmpClient = " (LSS Client)"
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname") & tmpClient
	tmpAddr = rsConfirm("CliAdrI") & " " & rsConfirm("caddress") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	If rsConfirm("CliAdd") = True Then 
		tmpDeptaddr =  rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
		tmpDeptaddrG = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
	End If
	tmpFon = rsConfirm("Cphone")
	tmpAFon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = rsConfirm("appTimeFrom") 
	tmpAppTTo = rsConfirm("appTimeTo")
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpmyclass = GetMyClass(rsConfirm("DeptID"))
	If rsConfirm("emerfee") = false Then
		tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	Else
		'EMERGENCY RATE
		Set rsRate = Server.CreateObject("ADODB.RecordSet")
		sqlRate = "SELECT * FROM EmergencyFee_T"
		rsRate.Open sqlRate, g_strCONN, 3,1
		If Not rsRate.EOF Then
			tmpFeeL = rsRate("FeeLegal")
			tmpFeeO = rsRate("FeeOther")
		End If
		rsRate.Close
		Set rsRate = Nothing
		If tmpmyclass = 3 or tmpmyclass = 5 Then
			tmpInstRate = tmpFeeL
		Else
			tmpInstRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
		EnD If
	End if
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
	tmpEmer = ""
	If rsConfirm("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpCom = rsConfirm("Comment")
	chkVer = ""
	If rsConfirm("Verified") = True Then chkVer = "checked"
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
	bInst = ""
	If rsConfirm("BillInst") = true Then bInst = "Checked"
	tmpTTrates = rsConfirm("TTRate")
	tmpMrates = rsConfirm("MRate")
	OverTTIntr = ""
	If rsConfirm("overTTIntr") = true Then OverTTIntr = "checked"
	OverTTInst = ""
	If rsConfirm("overTTInst") = true Then OverTTInst = "checked"
	OverMIntr = ""
	If rsConfirm("overMIntr") = true Then OverMIntr = "checked"
	OverMInst = ""
	If rsConfirm("overMInst") = true Then OverMInst = "checked"
	tmpCombil = rsConfirm("bilcomment")
	tmpcomintr = rsConfirm("intrcomment")
End If
rsConfirm.Close
Set rsConfirm = Nothing
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE index = " & RP
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
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE index = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
	'tmpIaddr = rsInst("address") & ", " & rsInst("City") & ", " &  rsInst("state") & ", " & rsInst("zip")
	'tmpBaddr = rsInst("Baddress") & ", " & rsInst("BCity") & ", " &  rsInst("Bstate") & ", " & rsInst("Bzip")
	'tmpBContact = rsInst("Blname") & ", " & rsInst("Bfname")
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE index = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	'tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpClass = rsDept("Class")
	If tmpDeptaddr = "" Then 
		tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE index = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
End If
rsLang.Close
Set rsLang = Nothing 
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE index = " & tmpIntr
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
	tmpIntrAdd = rsIntr("IntrAdrI") & " " & rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrAddG = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	PconIntr = GetPrime2(tmpIntr)
Else
	tmpIntrName = "<i>To be assigned.</i>"
	tmpIntr = 0
End If
rsIntr.Close
Set rsIntr = Nothing
'GET CANCELLATION REASON
Set rsCancel = Server.CreateObject("ADODB.RecordSet")
sqlCancel = "SELECT * FROM cancel_T"
rsCancel.Open sqlCancel, g_strCONN, 3, 1
Do Until rsCancel.EOF
	CancelMe = ""
	If tmpCancel = rsCancel("index") Then CancelMe = "selected"
	strCancel = strCancel & "<option value='" & rsCancel("index") & "' " & CancelMe & ">" & rsCancel("Reason") & "</option>" & vbCrLf
	rsCancel.MoveNext
Loop
rsCancel.Close
Set rsCancel = Nothing
'GET MISSED  REASON
Set rsMiss = Server.CreateObject("ADODB.RecordSet")
sqlMiss = "SELECT * FROM missed_T"
rsMiss.Open sqlMiss, g_strCONN, 3, 1
Do Until rsMiss.EOF
	MissMe = ""
	If tmpMiss = rsMiss("index") Then MissMe = "selected"
	strMiss = strMiss & "<option value='" & rsMiss("index") & "' " & MissMe & ">" & rsMiss("Reason") & "</option>" & vbCrLf
	rsMiss.MoveNext
Loop
rsMiss.Close
Set rsMiss = Nothing
'get mileage cap for interpreters
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travel_t"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecap = Z_czero(rsmile("milediff"))
	end if
	rsmile.close
	set rsmile = nothing
'get mileage rate
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from mileageRate_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilerate = Z_czero(rsmile("mileagerate"))
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
%>
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Service Verification - <%=Request("ID")%></title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script src=" http://maps.google.com/?file=api&amp;v=2.x&amp;key=ABQIAAAAF7oGqS_zHpf3iVxXTN7mvhTQ9IXshnULlpH6TVd_OTV0RdxjLRTjX3phYNjFFfzq0G34G5cpCNDJ2g"
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
			var decHrs;
			var decMile;
			var distMile;
			var bilMile;
			var bilTravel;
			var tmpRate;
	    function initialize() {
	      if (GBrowserIsCompatible()) {      
	        map = new GMap2(document.getElementById("map_canvas"));
	        //gdir = new GDirections(map, document.getElementById("directions"));
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
		     alert("No corresponding geographic location could be found for one of the specified addresses. This may be due to the fact that the address is relatively new, or it may be incorrect.\nError code: " + gdir.getStatus().code);
		   else if (gdir.getStatus().code == G_GEO_SERVER_ERROR)
		     alert("A geocoding or directions request could not be successfully processed, yet the exact reason for the failure is not known.\n Error code: " + gdir.getStatus().code);
		   
		   else if (gdir.getStatus().code == G_GEO_MISSING_QUERY)
		     alert("The HTTP q parameter was either missing or had no value. For geocoder requests, this means that an empty address was specified as input. For directions requests, this means that no query was specified in the input.\n Error code: " + gdir.getStatus().code);
	
		//   else if (gdir.getStatus().code == G_UNAVAILABLE_ADDRESS)  <--- Doc bug... this is either not defined, or Doc is wrong
		//     alert("The geocode for the given address or the route for the given directions query cannot be returned due to legal or contractual reasons.\n Error code: " + gdir.getStatus().code);
		     
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
				document.frmConfirm.txtActTT.value = Math.round((dureeHrs * 2) * 100)/100;
		  	document.frmConfirm.txtActMile.value = Math.round((distMile * 2) * 100)/100;
		  	//get data after deductions
		  	tmpRate = decMile / decHrs;
		  	if (decMile > <%=tmpmilecap%>) 
		  	{
					bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2); //billable mileage (2 way)
					bilTravel = bilMile / tmpRate; //billable travel time (2 way)
					document.frmConfirm.txtDecTT.value = Math.round(bilTravel * 100)/100;
		  		document.frmConfirm.txtDecMile.value = Math.round(bilMile * 100)/100;
		  			<% If tmpClass = 3 Or tmpClass = 5 Then %> //bill institution
		  			bilMileInst = (decMile * 2) - (<%=tmpmilecapcourts%> * 2); //billable mileage (2 way)
		  		<% Else %>
		  			bilMileInst = (decMile * 2) - (<%=tmpmilecapinst%> * 2); //billable mileage (2 way)
		  		<% End If %>
		  		bilTravelInst = bilMileInst / tmpRate; //billable travel time (2 way)
					document.frmConfirm.txtDecTTInst.value = Math.round(bilTravelInst * 100)/100;
		  		document.frmConfirm.txtDecMileInst.value = Math.round(bilMileInst * 100)/100;
		  		ComputeIntrTTM();
		  		ComputeInstTTM();
		  	}
				else
				{
					document.frmConfirm.txtActTT.value = 0;
					document.frmConfirm.txtActMile.value = 0;
					document.frmConfirm.txtDecTT.value = 0;
		  		document.frmConfirm.txtDecMile.value = 0;
		  		document.frmConfirm.txtDecTTInst.value = 0;
		  		document.frmConfirm.txtDecMileInst.value = 0;
				}
		 	}
		<% Else %>
		function noMap()
			{
				document.frmConfirm.txtActTT.value = 0;
			  document.frmConfirm.txtActMile.value = 0;
		  }
	  <% End If %>
		function CancelMe()
		{
			document.frmConfirm.selCancel.value = 0;
			document.frmConfirm.selCancel.disabled = true;
			if (document.frmConfirm.radioStat[3].checked == true || document.frmConfirm.radioStat[4].checked == true)
			{
				document.frmConfirm.selCancel.disabled = false;
			}
		}
		function CancelReason(xxx)
		{
			if (xxx !== 0)
			{
				document.frmConfirm.selCancel.value = xxx;
			}
		}
		function MissedMe()
		{
			document.frmConfirm.selMissed.value = 0;
			document.frmConfirm.selMissed.disabled = true;
			if (document.frmConfirm.radioStat[2].checked == true)
			{
				document.frmConfirm.selMissed.disabled = false;
			}
		}
		function MissedReason(xxx)
		{
			if (xxx !== 0)
			{
				document.frmConfirm.selMissed.value = xxx;
			}
		}
		function CompleteMe()
		{
			if (document.frmConfirm.radioStat[1].checked == true || document.frmConfirm.radioStat[4].checked == true)
			{
				document.frmConfirm.radioStat[0].disabled = true;
				document.frmConfirm.radioStat[2].disabled = true;
				document.frmConfirm.radioStat[3].disabled = true;
			}
		}
		
		function EditMe()
		{
			document.frmConfirm.action = "main.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
		}
		function EditMe2()
		{
			document.frmConfirm.action = "mainbill.asp?ID=" + <%=Request("ID")%>;
			document.frmConfirm.submit();
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
		function ChkComplete(IntrRate, InstRate)
		{
			var rate1 = new Boolean(true), rate2 = new Boolean(true)
				if (document.frmConfirm.txtActTFrom.value != "" && document.frmConfirm.txtActTTo.value != ""  && document.frmConfirm.txtBilHrs.value != "")
				{
					if (InstRate == "" || InstRate == 0)
					{
						rate1 = false;
					}
					if (IntrRate == "" || IntrRate == 0)
					{
						rate2 = false;
					}
					if(rate1 == false)
					{
						if (IntrRate != 0)
						{
							rate1 = true;
						}
					}	
					if(rate2 == false)
					{
						if (IntrRate != 0)
						{
							rate2 = true;
						}
					}	
					if (rate1 == false  && rate2 == false)
					{ 
						alert("ERROR: Please fill up required fields for billing."); 
						document.frmConfirm.radioStat[<%=tmpstat%>].checked = true;
						return;
					}
				}
				else
				{
					alert("ERROR: Please fill up required fields for billing."); 
					document.frmConfirm.radioStat[<%=tmpstat%>].checked = true;
					return;
				}
		}
		function KillMe(xxx)
		{
			var ans = window.confirm("Delete Request? Click Cancel to stop.");
			if (ans){
				document.frmConfirm.action = "action.asp?ctrl=9&ReqID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		function CalendarView(strDate)
		{
			document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmConfirm.submit();
		}
		function SaveBill(xxx)
		{
			document.frmConfirm.action = "action.asp?ctrl=10&ReqID=" + xxx;
			document.frmConfirm.submit();
		}
		function ComputeIntrTTM()
		{
			<% If OverTTIntr = "" Then %>
				document.frmConfirm.txtBilTIntr.value = Math.round((document.frmConfirm.txtDecTT.value * <%=tmpIntrRate%>) * 100)/100;
			<% Else %>
				document.frmConfirm.txtBilTIntr.value = <%=tmpBilTIntr%>;
			<% End If %>
			<% If OverMIntr = "" Then %>
				document.frmConfirm.txtBilMIntr.value = Math.round((document.frmConfirm.txtDecMile.value * <%=tmpmilerate%>) * 100)/100;
			<% Else %>
				document.frmConfirm.txtBilMIntr.value = <%=tmpBilMIntr%>;
			<% End If %>	
		}
		function InstRates()
		{
			if (document.frmConfirm.chkBillInst.checked == false)
			{
				document.frmConfirm.txtTTRate.value = 0;
				document.frmConfirm.txtMRate.value = 0;
				document.frmConfirm.txtTTRate.disabled = true;
				document.frmConfirm.txtMRate.disabled = true;
			}
			else
			{
				document.frmConfirm.txtTTRate.disabled = false;
				document.frmConfirm.txtMRate.disabled = false;
			}
		}
		function ComputeInstTTM()
		{
			<% If OverTTInst = "" Then %>
				document.frmConfirm.txtBilTInst.value = Math.round((document.frmConfirm.txtDecTTInst.value * document.frmConfirm.txtTTRate.value) * 100)/100;
			<% Else %>
				document.frmConfirm.txtBilTInst.value = <%=tmpBilTInst%>;
			<% End If %>
			<% If OverMInst = "" Then %>
				document.frmConfirm.txtBilMInst.value = Math.round((document.frmConfirm.txtDecMileInst.value * document.frmConfirm.txtMRate.value) * 100)/100;
			<% Else %>
				document.frmConfirm.txtBilMInst.value = <%=tmpBilMInst%>;
			<% End If %>	
		}
		function overwriteMe(xxx, yyy)
		{
			if (xxx.checked == true)
			{
				yyy.readOnly = false;
			}
			else
			{
				yyy.readOnly = true;
			}
		}
		-->
		</script>
		<body onload='CancelMe(); CancelReason(<%=tmpCancel%>); CompleteMe();MissedMe(); MissedReason(<%=tmpMiss%>);InstRates();
			<% If tmpIntr <> 0 Then %>
				initialize();' 
			<% Else %>
				noMap();'
			<% End If %>
			 onunload="GUnload();">
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - Service Verification</td>
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
								<tr>
									<td align='right' width='15%'>Rate:</td>
									<td class='confirm'><%=tmpInstRate%></td>
								</tr>
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
									<td class='confirm'><%=tmpIntrName%></td>
								</tr>
								<tr>
									<td align='right'>Rate:</td>
									<td class='confirm'><%=tmpIntrRate%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Interpreter Comment:</td>
									<td class='confirm'><%=tmpComIntr%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan='10' class='header'><nobr>Billing Information</td>
										</tr>
										<tr>
											<td align='right'><b>Status:</b></td>
											<td>
												<input type='radio' name='radioStat' value='0' <%=stat%> onclick='CancelMe(); MissedMe();'>&nbsp;<b>Pending</b>
												&nbsp;&nbsp;
												<input type='radio' <%=comp%>  name='radioStat' value='1' onclick='CancelMe(); MissedMe() ;ChkComplete(<%=tmpIntrRate%>, <%=tmpInstRate%>);'>&nbsp;<b>Completed</b>
												&nbsp;&nbsp;
												<input type='radio' name='radioStat' value='2' <%=misd%> onclick='CancelMe(); MissedMe();'>&nbsp;<b>Missed</b>
												&nbsp;&nbsp;
												<input type='radio' name='radioStat' value='3' <%=canc%> onclick='CancelMe(); MissedMe();'>&nbsp;<b>Canceled</b>
											</td>
											<td align='right'>
												Cancel Reason:
											</td>
											<td>
												<select name='selCancel' class='seltxt' style='width:150px;'>
													<option value='0'>-- Select a reason --</option>
													<%=strCancel%>
												</select>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td align='left'>
												<input type='radio' <%=canc2%>  name='radioStat' value='4' onclick='CancelMe(); MissedMe(); ChkComplete();'>&nbsp;<b>Canceled (billable)</b>
											</td>
											<td align='right'>
												Missed Reason:
											</td>
											<td>
												<select name='selMissed' class='seltxt' style='width:150px;'>
													<option value='0'>-- Select a reason --</option>
													<%=strMiss%>
												</select>
											</td>
										</tr>
										
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td align='right'>Billed:</td>
											<td><input type='checkbox' name='chkPaid' value='1' <%=chkPaid%> disabled ></td>
											
										</tr>
										<tr>
											<td align='right'>Billable Hours:</td>
											<td>
												<input class='main' size='5' maxlength='5' name='txtBilHrs' value='<%=tmpBilHrs%>'>
											</td>
											
										</tr>
										<tr>
											<td align='right'>Actual Time:</td>
											<td>
												&nbsp;From:<input class='main' size='5' maxlength='5' name='txtActTFrom' value='<%=tmpActTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
												&nbsp;To:<input class='main' size='5' maxlength='5' name='txtActTTo' value='<%=tmpActTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
												<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour</span>
											</td>
											
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td align='right'>Travel Time:</td>
											<td  align='left'>
													<input class='main' size='8' maxlength='8' readonly name='txtActTT' value='<%=tmpActTT%>' >
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">HRS.</span>
											</td>
										</tr>
										<tr>
											<td align='right'>Mileage:</td>
											<td  align='left'>
												<input class='main' size='8' maxlength='8' readonly name='txtActMile' value='<%=tmpActMile%>' >
												<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Miles</span>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td>
												<input type='checkbox' name='chkBillIntr' value='1' checked disabled >
												Pay Interpreter
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td align='left'>
												<table cellpadding='2' cellspacing='2' border='0' width='100%'>
													<tr>
														<td align='right'>
															Travel Time:
														</td>
														<td  align='center'>
															<input class='main' size='8' maxlength='8' readonly name='txtDecTT' value='<%=tmpDecTT%>' >
															<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">HRS.</span>
														</td>
														<td  align='center'>
															$<input class='main' size='8' maxlength='8' readonly name='txtBilTIntr' value='<%=tmpBilTIntr%>' >
															<input type='checkbox' name='chkBilTIntr' <%=OverTTIntr%> value='1' onclick='overwriteMe(this, document.frmConfirm.txtBilTIntr);'>
														</td>
													</tr>
													<tr>
														<td align='right'>
															Mileage:
														</td>
														<td  align='center'>
															<input class='main' size='8' maxlength='8' readonly name='txtDecMile' value='<%=tmpDecMile%>' >
															<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Miles</span>
														</td>
														<td  align='center'>
															$<input class='main' size='8' maxlength='8' readonly name='txtBilMIntr' value='<%=tmpBilMIntr%>' >
															<input type='checkbox' name='chkBilMIntr' <%=OverMIntr%> value='1' <%=bInst%> onclick='overwriteMe(this, document.frmConfirm.txtBilMIntr);'>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td>
												<input type='checkbox' name='chkBillInst' value='1' <%=bInst%> onclick='InstRates();ComputeInstTTM();'>
												Bill Institution
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td colspan='2'>
												<table cellpadding='2' cellspacing='2' border='0'>
													<tr>
														<td>Rates:</td>
														<td align='right'>Travel Time:</td>
														<td>
															$<input class='main' size='8' maxlength='8' name='txtTTRate' value='<%=tmpTTrates%>' onblur='ComputeInstTTM();'>
														</td>
													</tr>
													<tr>
														<td>&nbsp;</td>
														<td align='right'>Mileage:</td>
														<td>
															$<input class='main' size='8' maxlength='8' name='txtMRate' value='<%=tmpMrates%>' onblur='ComputeInstTTM();'>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td>&nbsp;</td>
											<td align='left'>
												<table cellpadding='2' cellspacing='2' border='0' width='100%'>
													<tr>
														<td align='right'>
															Travel Time:
														</td>
														<td  align='center'>
															<input class='main' size='8' maxlength='8' readonly name='txtDecTTInst' value='<%=tmpDecTTInst%>' >
															<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">HRS.</span>
														</td>
														<td  align='center'>
															$<input class='main' size='8' maxlength='8' readonly name='txtBilTInst' value='<%=tmpBilTInst%>' >
															<input type='checkbox' name='chkBilTInst' <%=OverTTInst%> value='1' onclick='overwriteMe(this, document.frmConfirm.txtBilTInst);'>
														</td>
													</tr>
													<tr>
														<td align='right'>
															Mileage:
														</td>
														<td  align='center'>
															<input class='main' size='8' maxlength='8' readonly name='txtDecMileInst' value='<%=tmpDecMileInst%>' >
															<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Miles</span>
														</td>
														<td  align='center'>
															$<input class='main' size='8' maxlength='8' readonly name='txtBilMInst' value='<%=tmpBilMInst%>' >
															<input type='checkbox' name='chkBilMIns' <%=OverMInst%> value='1' <%=bInst%> onclick='overwriteMe(this, document.frmConfirm.txtBilMInst);'>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										
										<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Billing Comment:</td>
										<td colspan='3' >
											<textarea name='txtcombil' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCombil%></textarea>
										</td>
									</tr>
										
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>&nbsp;</td>
											<td align='left' colspan='2'>
												* To make the request billable, please set actual time, billable hours, and rates.
											</td>
										</tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
												<input type='hidden' name="HID" value='<%=Request("ID")%>'>
												<input type='hidden' name="hidInstRate" value='<%=tmpInstRate%>'>
												<input type='hidden' name="hidIntrRate" value='<%=tmpIntrRate%>'>
												<input class='btn' type='button' value='Save' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveBill(<%=Request("ID")%>);'>
												<input class='btn' type='button' value='Back' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>';">
												<input class='btn' type='button' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe(<%=Request("ID")%>);'>
											</td>
									</tr>
									<tr>
										<td valign="top"><div id="directions" style="display: none;"></div></td>
									</tr>
									<tr>
											<td valign="top"><div id="map_canvas" style="display: none;"></div></td>
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