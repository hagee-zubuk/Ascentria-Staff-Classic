<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
frmdte = Z_DateNull(Request("frmdte"))
todte = Z_DateNull(Request("todte"))
myLang = Z_Czero(Request("selLang"))
'get info on request
tmpDept = Z_GetInfoFROMAppID(Request("appID"), "DeptID")
tmpAppDate = Z_GetInfoFROMAppID(Request("appID"),"appdate")
If Z_GetInfoFROMAppID(Request("appID"), "CliAdd") = True Then 
	tmpDeptaddrG = Z_GetInfoFROMAppID(Request("appID"),"CAddress") &", " & Z_GetInfoFROMAppID(Request("appID"),"CCity") & ", " & Z_GetInfoFROMAppID(Request("appID"),"CState") & ", " & Z_GetInfoFROMAppID(Request("appID"),"CZip")
	tmpZipInstg = Z_GetInfoFROMAppID(Request("appID"),"czip")
End If
FreeMe = 0
Workme = 0
'check if assign or on vaca but current table not refreshed
If AppAssigned(Request("appID")) Then FreeMe = 1
If OnVacation(Request("IntrID"), tmpAppDate) Then Workme = 1
If FreeMe = 1 Then
	Session("MSG") = "This appointment (ID: " & Request("appID") & ") has already been assigned."
	Response.Redirect "openappts.asp?reload=1&frmdte=" & frmdte & "&todte=" & todte & "&selLang=" & myLang
End If
If Workme = 1 Then
	Session("MSG") = "Interpreter being assigned for this appointment (ID: " & Request("appID") & ") is on vacation."
	Response.Redirect "openappts.asp?reload=1&frmdte=" & frmdte & "&todte=" & todte & "&selLang=" & myLang
End If
'save inter
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
rsIntr.Open "UPDATE request_T SET intrID = " & Request("IntrID") & " WHERE [index] = " & Request("appID"), g_strCONN, 1, 3
Set rsIntr = Nothing
'save assigner
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
rsIntr.Open "UPDATE request_T SET assignedby = " & Request.Cookies("UID") & " WHERE [index] = " & Request("appID"), g_strCONN, 1, 3
Set rsIntr = Nothing
'save default rate of interpreter
intrRate = Z_GetDefRate(Request("IntrID"))
If Z_EligibleHigherPay(Z_GetInfoFROMAppID(Request("appID"), "LangID")) Then IntrRate = Z_GetHigherPay(intrRate, Request("IntrID"))
Set rsIntrrate = Server.CreateObject("ADODB.RecordSet")
rsIntrrate.Open "UPDATE request_T SET IntrRate = " & IntrRate & " WHERE [index] = " & Request("appID"), g_strCONN, 1, 3
Set rsIntrrate = Nothing
'save inter in HP if applicable
HPID = GetReqHPID2(Request("appID"))
If HPID > 0 Then
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "UPDATE appointment_T SET intrID = " & Request("IntrID") & " WHERE [index] = " & HPID, g_strCONNHP, 1, 3
	Set rsIntr = Nothing
End If
'savehist
'SAVE HISTORY
TimeNow = Now
Set rsHist = Server.CreateObject("ADODB.RecordSet")
sqlHist = "SELECT * FROM History_T WHERE ReqID = " & Request("appID")
rsHist.Open sqlHist, g_strCONNHist, 1,3 
If Not rsHist.EOF Then
	If Z_CZero(rsHist("interID")) <> Z_CZero(Request("IntrID")) Then
		rsHist("interID") = Z_CZero(Request("IntrID"))
		rsHist("interTS") = TimeNow
		rsHist("interU") = Request.Cookies("LBUsrName")
	End If
Else
	rsHist.AddNew
	rsHist("interID") = Z_CZero(Request("IntrID"))
	rsHist("interTS") = TimeNow
	rsHist("interU") = Request.Cookies("LBUsrName")
End If
rsHist.Update
rsHist.Close
Set rsHist = Nothing
'get interpreter address and email
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT address1, City, state, [zip code], [E-mail] FROM interpreter_T WHERE [index] = " & Request("IntrID")
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpEmail = rsIntr("E-mail")
	tmpIntrAddG = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrZip = rsIntr("Zip Code")
End If
rsIntr.Close
Set rsIntr = Nothing
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpClass = rsDept("Class")
	If tmpDeptaddrG = "" Then 
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpZipInstg = rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'get other values
'get mileage cap for interpreters
set rsmile = server.createobject("adodb.recordset")
sqlmile = "select * from travel_t"
rsmile.open sqlmile, g_strconn, 3, 1
if not rsmile.eof then
	tmpmilecap = Z_czero(rsmile("milediff"))
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
'get mileage cap for institutions except courts
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travelInst_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecapinst = Z_czero(rsmile("milediffinst"))
	end if
	rsmile.close
	set rsmile = nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
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
				        document.frmAssignApp.submit();
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
	   		document.frmAssignApp.txtRTravel.value = Math.round((dureeHrs * 2) * 100)/100;
	   		document.frmAssignApp.txtRMile.value = Math.round((distMile * 2) * 100)/100; 
		  	if (decMile > <%=tmpmilecap%>) //interpreter
		  	{
					bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2); //billable mileage (2 way)
					bilTravel = bilMile / tmpRate; //billable travel time (2 way)
				
		  		document.frmAssignApp.txtTravel.value = Math.round(bilTravel * 100)/100;
		  		document.frmAssignApp.txtMile.value = Math.round(bilMile * 100)/100; 
		 		}
		 		else
		 		{
		 			document.frmAssignApp.txtTravel.value = 0
		  		document.frmAssignApp.txtMile.value = 0
		 		}
		  	//institution
				<% If tmpClass = 3 Or tmpClass = 5 Then %> //bill institution
					if (decMile > <%=tmpmilecapcourts%>)
					{
	  				bilMileInst = (decMile * 2) - (<%=tmpmilecapcourts%> * 2); //billable mileage (2 way)
	  				document.frmAssignApp.txtMileInst.value = Math.round(bilMileInst * 100)/100;
	  			}
	  			else
	  			{
	  				document.frmAssignApp.txtMileInst.value = 0;
	  			}
	  		<% Else %>
	  			if (decMile > <%=tmpmilecapinst%>)
					{
	  				bilMileInst = (decMile * 2) - (<%=tmpmilecapinst%> * 2); //billable mileage (2 way)
	  				document.frmAssignApp.txtMileInst.value = Math.round(bilMileInst * 100)/100;
	  			}
	  			else
	  			{
	  				document.frmAssignApp.txtMileInst.value = 0;
	  			}
	  		<% End If %>
	  		if (document.frmAssignApp.txtMileInst.value > 0)
	  		{
	  			bilTravelInst = bilMileInst / tmpRate; //billable travel time (2 way)
	  			document.frmAssignApp.txtTravelInst.value = Math.round(bilTravelInst * 100)/100;
				}
				else
				{
					document.frmAssignApp.txtMileInst.value = 0;
				}
	   	}
	   	function chkvalue() {
	   		
	   		//if (document.frmAssignApp.txtRTravel.value != "") {
					alert("TT: " + document.frmAssignApp.txtRTravel.value);
				//} 
				//else {
				//	chkvalue();
				//}
	   	}
	   
    //-->
    </script>
    <script src="https://maps.googleapis.com/maps/api/js?key=<%=googlemapskey%>" type="text/javascript"></script>
	</head>
	<body onload="initMap();">
		<form name="frmAssignApp" method="POST" action="saveTTM.asp">
			<div id="output" style="display: none;"></div>
			<div id="map" style="display: none;"></div>
			<div style="position: fixed; top: 50%; left: 50%; margin-top: -50px; margin-left: -75px;" border=1>
				<img src="images/loading.gif" border='0' />
				<br>
				<p align="center"><b>GETTING DATA FROM GOOGLE MAPS...</b></P>
			</div>
			<input type="hidden" name="frmdte" value='<%=frmdte%>'>
			<input type="hidden" name="todte" value='<%=todte%>'>
			<input type="hidden" name="selLang" value='<%=myLang%>'>
			<input type="hidden" name="txtRTravel">
			<input type="hidden" name="txtRMile">
			<input type="hidden" name="txtTravel">
			<input type="hidden" name="txtMile">
			<input type="hidden" name="txtMileInst">
			<input type="hidden" name="txtTravelInst">
			<input type="hidden" name="appID" value="<%=Request("appID")%>">
		</form>
	</body>
</html>
<script>
	
</script>
