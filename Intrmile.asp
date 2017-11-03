<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
	Function GetDeptZip(xxx)
		Set rsDept = Server.CreateObject("ADODB.RecordSet")
		sqlDept = "SELECT * FROM dept_T WHERE [index] = " & xxx
		rsDept.Open sqlDept, g_strCONN, 3, 1
		If Not rsDept.EOF Then
			GetDeptZip = rsDept("zip")
		End If
		rsDept.Close
		Set rsDept = Nothing 
	End Function
	'get mileage cap for interpreters
	set rsmile = server.createobject("adodb.recordset")
	sqlmile = "select * from travel_t"
	rsmile.open sqlmile, g_strconn, 3, 1
	if not rsmile.eof then
		tmpmilecap = Z_czero(rsmile("milediff"))
	end if
	rsmile.close
	set rsmile = nothing
	'GET ADDRESS AND ZIP of Intrpreter
	Set rsIntr = Server.CreateObject("ADODB.REcordSet")
	sqlIntr = "SELECT * FROM Interpreter_T WHERE [index] = " & Request("selIntr")
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	If Not rsIntr.EOF Then
		tmpIntrAddG = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
		tmpIntrZip = rsIntr("Zip Code")
		tmpAvail = rsIntr("Availability")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	'GET ADDRESS AND ZIP of DEPARTMENT/CLIENT
	Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
	sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("RID")
	rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
	If Not rsConfirm.EOF Then
		If rsConfirm("CliAdd") = True Then 
			tmpDeptaddrG = rsConfirm("CAddress") &", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
			tmpZipInstg = rsConfirm("czip")
		Else
			tmpDeptaddrG = GetDeptAdr(rsConfirm("DeptID"))
			tmpZipInstg = GetDeptZip(rsConfirm("DeptID"))
		End If
	End If
	rsConfirm.CLose
	Set rsConfirm = Nothing
%>
<html>
	<head>
		<title>Email Interpreter</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
    <script language='JavaScript'>
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
				        document.getElementById("btnOK").disabled = false;
				      }
				    }
				  }
				  else {
				  	alert('Error: One of the addresses is invalid. System used ZIP CODES to calculate Travel Time and Mileage');
				  	var service = new google.maps.DistanceMatrixService();
				  	calculateDistances(service, originzip, desinationzip);
				  }
				}
		function getDistanceValues(dista, dura) {
	      // Use this function to access information about the latest load()
	      // results.
				duree = dura;
				dist = dista;
				dureeHrs = ((duree) / 60) / 60;
				distMile = dista / 1609.344;
				//document.getElementById("ttM").innerHTML = (Math.round(dureeHrs * 100)/100) + " Hrs. - " + (Math.round(distMile*100)/100) + " Miles"; 
	   		decHrs = dureeHrs;
				decMile = distMile;
	   		tmpRate = decMile / decHrs;
	   		//alert(decHrs + "         " + decMile);
	   		if (decMile > <%=tmpmilecap%>) //interpreter
		  	{
		  		bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2); //billable mileage (2 way)
					bilTravel = bilMile / tmpRate; //billable travel time (2 way)
		   		document.frmMile.txtTravel.value = Math.round(bilTravel * 100)/100;
			  	document.frmMile.txtMile.value = Math.round(bilMile * 100)/100;
		  	}
		  	else
		  	{
		  		document.frmMile.txtTravel.value = 0;
		  		document.frmMile.txtMile.value = 0;
		  	}
		  	//alert(document.frmMile.txtTravel.value + "         " + document.frmMile.txtMile.value);
		 	}
		 	function SubmitMe()
		 	{
		 			//alert(document.frmMile.txtTravel.value + "    " + document.frmMile.txtMile.value);
		 			document.frmMile.action = "emailIntr.asp"
		 			document.frmMile.submit();	
		 	}
    </script>
    <script src="https://maps.googleapis.com/maps/api/js?key=<%=googlemapskey%>" type="text/javascript"></script>
	</head>
	<body onload='document.getElementById("btnOK").disabled = true;initMap();'>
		<form name='frmMile' method='post'>
			<center>
			<table>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>
						Intepreter:
					</td>
					<td>
						<b><%=GetIntr(Request("selIntr"))%></b>
					</td>
				</tr>
				<tr>
					<td align='right' valign='top'>
						Availablity:
					</td>
					<td>
						<textarea readonly><%=tmpAvail%></textarea>
					</td>
				</tr>
				<tr>
					<td align='right'>
						Mileage:
					</td>
					<td>
						<input class='main' size='5' readonly name='txtMile'>&nbsp;miles
					</td>
				</tr>
				<tr>
					<td align='right'>
						Travel Time:
					</td>
					<td>
						<input class='main' size='5' readonly name='txtTravel'>&nbsp;hrs
					</td>
				</tr>	
				<tr>
					<td colspan='2' align='center'>
							<input class='btn' type='button' name="btnOK" id="btnOK" value='OK' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SubmitMe();'>
								<input class='btn' type='button' value='Back' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='document.location="emailIntr.asp?ID=<%=Request("RID")%>";'>
					</td>
				</tr>
			</table>
			
			<input type='hidden' name='selIntr' value='<%=Request("selIntr")%>'>
			<input type='hidden' name='adr1'  value='<%=tmpIntrAdd%>'>
			<input type='hidden' name='adr2'  value='<%=tmpDeptaddr%>'>
			<input type='hidden' name='zip1'  value='<%=tmpIntrZip%>'>
			<input type='hidden' name='zip2'  value='<%=tmpZipInst%>'>
			<input type='hidden' name='ID'  value='<%=Request("RID")%>'>
			<tr>
									<td valign="top"><div id="output" style="display: none;"></div></td>
								</tr>
								<tr>
										<td valign="top"><div id="map" style="display: none;"></div></td>
								</tr>
		</form>
	</body>
</html>
<!-- #include file="_closeSQL.asp" -->
<script language='JavaScript'>
	//initialize();
	
</script>