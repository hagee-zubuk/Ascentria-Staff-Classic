<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_googleDMA.asp" -->
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<title>Override Travel Time or Mileage</title>
	<meta name="description" content="Override Travel Time or Mileage">
	<meta name="author" content="Hagee@zubuk">
 	<link rel="stylesheet" type="text/css" href="css/normalize.css" />
 	<link rel="stylesheet" type="text/css" href="css/skeleton.css" />
 	<link rel="stylesheet" type="text/css" href="css/jquery-ui.min.css" />
	<link rel="stylesheet" type="text/css" href="style.css" />
	<script langauge="javascript" type="text/javascript" src="js/jquery-3.3.1.min.js"></script>
	<script langauge="javascript" type="text/javascript" src="js/jquery-ui.min.js"></script>
	<script src="https://maps.googleapis.com/maps/api/js?key=<%=googlemapskey%>" type="text/javascript"></script>
  <!--[if lt IE 9]>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/html5shiv.js"></script>
  <![endif]-->
	<style>
div.formatsmall { display: inline-block; }
.ui-autocomplete-loading { background: white url("images/ui-anim_basic_16x16.gif") right center no-repeat; }
table.summary { width: 100%; }
.summary th { text-align: center; }
.summary td { text-align: center; padding-top: 2px; padding-bottom: 3px; vertical-align: text-top; line-height: 1.0em;}
.summary input { height: 20px; padding: 2px 4px; border-radius: 2px; margin: 1px;}
.summary td input[type=checkbox] { height: 13px; padding: 1px; border-radius: 0px; }
.summary td:first-child	{ text-align: right; padding-right: 2px;}
.makered { color: red; }
.ico_view { display: inline-block; float: left; }
.onvacation { background-color: pink; }
.address { display: inline-block; float: left; width: 300px; }
	</style>
</head>
<body>
	<div class="container">
<%
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function

Function Z_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
	strTmp = DatePart("yyyy", dtDate)
	Z_YMDDate = strTmp & "-" & Z_YMDDate
End Function

Function OnVacation(IntrID, appDate)
	OnVacation = False
	Set rsVac = Server.CreateObject("ADODB.RecordSet")
	sqlVac = "SELECT vacto, vacfrom, vacto2, vacfrom2 FROM interpreter_T WHERE [index] = " & intrID
	rsVac.Open sqlVac, g_strCONN, 3, 1
	If Not rsVac.EOF Then
		If Not IsNull(rsVac("vacfrom")) Then
			If appDate >= rsVac("vacfrom") And appDate <= rsVac("vacto") Then 
				OnVacation = True
			End If
		End If
		If onVacation = False Then
			If Not IsNull(rsVac("vacfrom2")) Then
				If appDate >= rsVac("vacfrom2") And appDate <= rsVac("vacto2") Then 
					OnVacation = True
				End If
			End If
		End If
	End If
	rsVac.Close
	Set rsVac = Nothing
End Function



'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If

IntrLang = "nothing"
tmpIntr = -1
'intrmileage.asp?id=<request id>
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT req.[index], req.[intrID], req.[appDate], req.[appTimeFrom], lan.[language]" & _
		", CASE WHEN req.[CliAdd] = 1 THEN " & _
		"req.[CAddress] + ', ' + req.[CCity] + ', ' + req.[CState] + ' ' + req.[CZip] " & _
		"ELSE " & _
		"dep.[Address] + ', ' + dep.[City] + ', ' + dep.[State] + ' ' + dep.[Zip] " & _
		"END AS [dest_Address] " & _
		", CASE WHEN req.[CliAdd] = 1 THEN " & _
		"req.[CZip] " & _
		"ELSE " & _
		"dep.[Zip] " & _
		"END AS [dest_ZIP] " & _
		"FROM [request_T] AS req " & _
		"INNER JOIN [dept_T] AS dep ON req.[deptid] = dep.[index] " & _
		"INNER JOIN [language_T] AS lan ON req.[langid]=lan.[index] WHERE req.[index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not (rsConfirm.EOF) Then
	reqID 		= Z_CLng( rsConfirm("index") )
	IntrLang 	= Z_FixNull( rsConfirm("language") )
	tmpIntr 	= Z_CLng( rsConfirm("intrID") )
	appDate 	= Z_YMDDate( rsConfirm("appDate") )
	tmpAvail 	= Weekday(appDate) & "," & Hour( rsConfirm("appTimeFrom") )
	strDest 	= Z_FixNull( rsConfirm("dest_Address"))
	zipDest		= Z_CLng( rsConfirm("dest_ZIP"))
End If
rsConfirm.Close
Set rsConfirm = Nothing	
strApptAddr = strDest
sqlIntrLang = "SELECT itr.[index], itr.[last name], itr.[first name]" & _
		", itr.[address1] + ', ' + itr.[city] + ', ' + itr.[state] + ' ' + itr.[zip code] AS [address]" & _
		", itr.[zip code] AS [zip]" & _
		", CASE WHEN (NOT( itr.[vacfrom]  IS NOT NULL AND itr.[vacto]  IS NOT NULL AND itr.[vacfrom]  <= '" & appDate & "' AND itr.[vacto]  >= '" & appDate & "') " & _
		"AND NOT( itr.[vacfrom2] IS NOT NULL AND itr.[vacto2] IS NOT NULL AND itr.[vacfrom2] <= '" & appDate & "' AND itr.[vacto2] >= '" & appDate & "') ) THEN 0 " & _
		"ELSE 1 END AS [onVacation] " & _
		"FROM interpreter_T AS itr " & _
		"LEFT JOIN [Avail_T] AS ava ON itr.[index]=ava.[intrID] " & _
		"LEFT JOIN [tmpGoogleDist] AS goo ON itr.[index]=goo.[intrid] AND goo.[reqid]=" & reqID & " " & _
		"WHERE (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & _
		"' OR Upper(Language3) = '" & IntrLang & "' OR Upper(Language4) = '" & IntrLang & _
		"' OR Upper(Language5) = '" & IntrLang & "' OR Upper(Language6) = '" & IntrLang & "') " & _
		"AND [Active] = 1 " & _
		"AND (ava.[avail] IS NULL OR ava.[avail] = '" & tmpAvail & "' ) " & _
		"ORDER BY [Last Name], [First Name]" 
Set rsIntrLang = Server.CreateObject("ADODB.RecordSet")
rsIntrLang.Open sqlIntrLang, g_strCONN, 3, 1
'Response.Write sqlIntrLang
'rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
'tmpIntrZip = rsIntr("Zip Code")
%>
	<div class="row" style="margin-top: 10px;">
		<div class="twelve columns">
			<h4 id="lblResults">Interpreter Travel Times &amp; Mileage</h4>
			<div>Destination: <address><%= strDest %></address></div>
			<table class="summary">
				<thead><tr><th>Interpreter</th><th>Mileage</th><th>Travel Time</th></thead>
				<tbody>
<%
strZipArray = ""
strOrigs = ""
Do Until rsIntrLang.EOF
	tmpIntrName = CleanMe(rsIntrLang("last name")) & ", " & CleanMe(rsIntrLang("first name"))
	' Or (Avail(rsIntrLang("index"), tmpAvail) And NotRestrict(rsIntrLang("index"), tmpInst, tmpDept)) Then
	'If NotRestrict(rsIntrLang("index"), tmpInst, tmpDept) = false Then rest = " (restricted)"
	Response.Write "<tr><td "
	If rsIntrLang("onVacation") = 1 Then
		Response.Write "class=""onvacation"""
		tmpIntrName = " (on vacation) " & tmpIntrName
	End If
	Response.Write ">" & tmpIntrName & "</td><td><div id=""miles" & rsIntrLang("index") & """>"
	tmpItrZIP = Z_CLng( rsIntrLang("zip") )
	blnOK = CBool(Abs(zipDest-tmpItrZIP) > 3)
	If (blnOK) Then
		strItrZIP = Z_FixNull(rsIntrLang("zip"))
		strItrAdr = Z_FixNull(rsIntrLang("address"))
		If Len(strItrAdr) > 6 Then
			strZipArray		= strZipArray & "zips.set('" & rsIntrLang("index") & "','" & strItrZIP & ");" & vbCrLf
			strOrigs 		= strOrigs & strItrAdr
			strDest			= strDest & strApptAddr
			Response.Write "<img src=""images/ajax-loader-small.gif"" alt="".."" title=""loading"" />" & _
					"ZIP: " & strItrZIP & "</div></td>"
			Response.Write "<td><div id=""trav" & rsIntrLang("index") & """>"
			Response.Write "<img src=""images/ajax-loader-small.gif"" alt="".."" title=""loading"" />"
		Else
			blnOK = False
			Response.Write "(bad ZIP)</div></td><td><div id=""trav" & rsIntrLang("index") & """> -- "
		End If
	Else
		Response.Write "(close)</div></td><td><div id=""trav" & rsIntrLang("index") & """> -- "
	End If
	Response.Write "</div></td></tr>" & vbCrLf

	rsIntrLang.MoveNext
	If Not rsIntrLang.EOF And blnOK Then
		strOrigs = strOrigs & "|"
		strDest = strDest & "|"
	End If
Loop
rsIntrLang.Close
Set rsIntrLang = Nothing

%>
	</div><!-- container -->
</body>
</html>
<script language="javascript" type="text/javascript"><!--
var desinationdef	= "<%= strDest %>"	;
var desinationzip	= "<%= zipDest %>"	;
var originzips		= "<%= strOrigs %>" ;

function getDMAData() {
	var service = new google.maps.DistanceMatrixService();
	calculateDistances(service, originzips, desinationdef);	
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
				// getDistanceValues(distance, duration);
			}
		}
	} else {
			//alert('one of the addresses and/or zip codes are invalid.\n\nUnable to calculate distance.');
		$('#txtRTravel').html('n/a');
		$('#txtRMile').html('n/a');
		$('#txtMile').html('&mdash;');
		$('#txtTravel').html('&mdash;');
		$('#txtMileInst').html('&mdash;');
		$('#txtTravelInst').html('&mdash;');
		$('#txtRate').html('n/a');
	}
}
function getDistanceValues(dista, dura){ 
// Use this function to access information about the latest load() results.
	duree = dura;
	dist = dista;
	dureeHrs = ((duree) / 3600);
	distMile = dist / 1609.344;
	decHrs = dureeHrs;
	decMile = distMile;
	tmpRate = decMile / decHrs;
	/*
	$('#txtRTravel').html( Math.round((dureeHrs * 2) * 100)/100 );
	$('#txtRMile').html( Math.round((distMile * 2) * 100)/100 );  
	// interpreter section
	var bilMile = (decMile * 2) - (<%=tmpmilecap%> * 2);	//billable mileage (2 way)
	$('#txtMile').html( Math.round(bilMile * 100)/100 ); 
	if (bilMile < 0) {
		$('#txtMile').addClass("makered");
		//$('#txtMile2').val('0.00');
	}
	var bilTravel = bilMile / tmpRate;						//billable travel time (2 way)	
	$('#txtTravel').html( Math.round(bilTravel * 100)/100 );
	if (bilTravel < 0) {
		$('#txtTravel').addClass("makered");
		//$('#txtTravel2').val('0.00');
	}

	//institution
	var bilMileInst = (decMile * 2) - (<%=tmpMileCapInst%> * 2); //billable mileage (2 way)
	$('#txtMileInst').html( Math.round(bilMileInst * 100)/100 );
	if (bilMileInst < 0) {
		$('#txtMileInst').addClass("makered");
		//$('#txtMileInst2').val('0.00');
	}
	var bilTravelInst = bilMileInst / tmpRate; //billable travel time (2 way)
	$('#txtTravelInst').html( Math.round(bilTravelInst * 100)/100 );
	if (bilTravelInst < 0) {
		$('#txtTravelInst').addClass("makered");
		//$('#txtTravelInst2').val('0.00');
	}

	$('#txtRate').html( Math.round(tmpRate * 10)/10 );
	*/
}
$( document ).ready(function() {
	getDMAData();
});
// --></script>
<%
'Set oGDM = New acaDistanceMatrix
'oGDM.DBCONN = g_strCONN
%>