<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_googleDMA.asp" -->
<%
'ovrd_ttm.asp?ReqID=

Function Z_FormatTime(xxx, zzz)
	Z_FormatTime = ""
	If Not Z_Blank(xxx) Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, zzz)
	End If
End Function

strPostBack = Z_FixNull(Request("postback"))
If strPostBack = "" Then
	strPostBack = Request.ServerVariables("HTTP_REFERER")
Else
	strPostBack = LCase(strPostBack)
End If

'get mileage rate
Set rsmile = Server.CreateObject("ADODB.Recordset")
sqlMile = "select * from mileageRate_T"
rsmile.open sqlMile, g_strconn, 3, 1
If Not rsmile.eof Then
	tmpmilerate = Z_czero(rsmile("mileagerate"))
End If
rsmile.Close
Set rsmile = Nothing
'get mileage cap for interpreters
Set rsmile = Server.CreateObject("ADODB.Recordset")
sqlMile = "select * from travel_t"
rsmile.open sqlMile, g_strconn, 3, 1
If Not rsmile.EOF Then
	tmpmilecap = Z_czero(rsmile("milediff"))
End If
rsmile.Close
Set rsmile = Nothing


strReqID = Z_FixNull(Request("ReqID"))
If strReqID = "" Then
	Session("MSG") = "Appointment id missing"
	Response.Redirect strPostBack '' & "?id=" & strReqID
End If

Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT req.*, xrs.[statusname], ins.[Facility], dep.[Dept], lan.[language] " & _
		", dep.[City] + ', ' + dep.[State] AS [dept_addr]" & _
		", CASE " & _
		"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CAddress] + ', ' + req.[CCity] + ', ' + req.[CState] + ' ' + req.[CZip] " & _
		"ELSE " & _
				"dep.[Address] + ', ' + dep.[City] + ', ' + dep.[State] + ' ' + dep.[Zip] " & _
		"END AS [dest_Address] " & _
		", CASE " & _
		"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CZip] " & _
		"ELSE " & _
				"dep.[Zip] " & _
		"END AS [dest_ZIP] " & _
		"FROM [Request_T] AS req " & _
		"INNER JOIN [xrStatus] AS xrs ON req.[status] = xrs.[index] " & _
		"INNER JOIN [institution_T] AS ins ON req.[instid] = ins.[index] " & _
		"INNER JOIN [dept_T] AS dep ON req.[deptid] = dep.[index] " & _
		"INNER JOIN [language_T] AS lan ON req.[langid] = lan.[index] " & _
		"WHERE req.[index] = " & strReqID '& _
		'" AND ([Processed] IS NULL OR [Processed] = '') " & _
		'" AND ([Processedmedicaid] IS NULL OR [ProcessedMedicaid] = '') " 
'Response.Write sqlConfirm
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If rsConfirm.EOF Then
	rsConfirm.close
	Set rsConfirm = Nothing
	Session("MSG") = "Appointment not found, or you cannot view/edit this Appointment."
	Response.Redirect strPostBack' & "?id=" & strReqID
	'Response.End
End If

TS = rsConfirm("timestamp")
RP = rsConfirm("reqID") 

tmpDeptaddr = rsConfirm("dept_addr")
'tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & ", " & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
tmpDeptaddrG = rsConfirm("dest_Address")
tmpZipInstG = rsConfirm("dest_Zip")
If rsConfirm("CliAdd") = True Then 
	tmpAddr = "(client address)"
Else
	tmpAddr = tmpDeptaddr ' rsConfirm("dest_Address")
End If
tmpDOB = rsConfirm("DOB")
tmpSalita = rsConfirm("language")
tmpAppDate = rsConfirm("appDate")
tmpAppTFrom = CTime(rsConfirm("appTimeFrom"))
tmpAppTTo = CTime(rsConfirm("appTimeTo"))
tmpAppLoc = rsConfirm("appLoc")
tmpIname = rsConfirm("facility")
tmpDname = rsConfirm("dept")

chkPayIntr = ""
If rsConfirm("PayIntr") = TRUE Then
	chkPayIntr = "checked"
End If
chkBillInst = ""
If rsConfirm("BillInst") = TRUE Then
	chkBillInst = "checked"
End If
tmpToll = Z_FixNull( rsConfirm("toll") )
If tmpToll <> "" And Z_CDbl(tmpToll) > 0 Then tmpToll =	Z_FormatNumber(tmpToll, 2)
tmpChkLBconfirmToll = ""
If rsConfirm("LBconfirmToll") = True Then
	tmpChkLBconfirmToll = "checked"
End If
'var desinationdef	= "tmpDeptaddrG"	;
'var desinationzip	= "tmpZipInstg"	;

If rsConfirm("training") Then
	tmpTrain = " (Training Appointment)"
Else
	tmpTrain = ""
End If
tmpIntrRate = Z_FormatNumber(rsConfirm("IntrRate"), 2)
tmpCom = rsConfirm("Comment")
Statko = rsConfirm("statusname")
tmpstat = rsConfirm("Status")
tmpBilHrs = rsConfirm("Billable")
tmpActTFrom = Z_FormatTime(rsConfirm("astarttime"), 3) 
tmpActTTo = Z_FormatTime(rsConfirm("aendtime"), 3)

tmpMrate = "0.50" 'Z_FormatNumber(rsConfirm("MRate"), 2)
tmpRealTT = rsConfirm("RealTT")
tmpRealM = rsConfirm("RealM")

tmpIntrTT = Z_FormatNumber(rsConfirm("actTT"), 2)
tmpBilTIntr = Z_FormatNumber( Z_CZero(rsConfirm("actTT")) * Z_CZero(rsConfirm("intrrate")) , 2)
tmpIntrMI = Z_FormatNumber(rsConfirm("actMil"), 2)
tmpBilMIntr = Z_FormatNumber( rsConfirm("actMil") * tmpmilerate, 2)

tmpInstActMil = Z_FormatNumber( rsConfirm("InstActMil"), 2)
tmpBilMInst = Z_FormatNumber( rsConfirm("M_Inst"), 2)
tmpInstActTT = Z_FormatNumber( rsConfirm("InstActTT"), 2)
tmpBilTInst = Z_FormatNumber( rsConfirm("TT_Inst"), 2)
ttrate = rsConfirm("TTrate")

tmpIntr = rsConfirm("intrID")
tmpComintr = rsConfirm("intrcomment")
tmpcombil = rsConfirm("bilcomment")
tmpLBcom = rsConfirm("LBcomment")
tmpHPID = Z_CZero(rsConfirm("HPID"))
mrrec = rsConfirm("mrrec")
cc_email = Z_FixNull( rsConfirm("cc_addr") )

rsConfirm.close
Set rsConfirm = Nothing

Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If rsIntr.EOF Then
		rsIntr.close
	Set rsIntr = Nothing
	Session("MSG") = "Interpreter is unassigned: unable to set/manage travel time"
	Response.Redirect strPostBack' & "?id=" & strReqID

End If
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
'PconIntr = GetPrime2(tmpIntr)
tmpZipIntr = ""
If rsIntr("Zip Code") <> "" Then tmpZipIntr = rsIntr("Zip Code")
intrPID = rsIntr("PID")
intrWID = rsIntr("WID")	
intrXID = rsIntr("XID")	
If rsIntr("certIntr") Then 
	ttrate = "38"
Else
	If (ttrate <= 0) Then ttrate = 28
	'ttrate = "28.00"
End If
ttrate = Z_FormatNumber( ttrate, 2)
TTM = ""

rsIntr.Close
Set rsIntr = Nothing

'get mileage cap for institutions / courts
Set rsmile = Server.CreateObject("ADODB.Recordset")
If tmpClass = 3 Or tmpClass = 5 Then ' courts'
	sqlmile = "select * from travelInstCourt_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	If Not rsmile.EOF Then
		tmpMileCapInst = Z_czero(rsmile("milediffcourt"))
	end if
Else
	sqlmile = "select * from travelInst_T"
	rsmile.open sqlmile, g_strconn, 3, 1
	If Not rsmile.EOF Then
		tmpMileCapInst = Z_czero(rsmile("milediffinst"))
	End If
End If
rsmile.Close
Set rsmile = Nothing


Set oGDM = New acaDistanceMatrix
oGDM.DBCONN = g_strCONN
Call oGDM.FetchMileageV2(strReqID, tmpIntr, tmpIntrAddG, tmpIntrZip, FALSE)
'// Call oGDM.FetchMileageFromReqID(rsReq("index"), TRUE)
fltRealTT	= oGDM.fltRealTT
fltRealM	= oGDM.fltRealM
fltActTT	= oGDM.fltActTT
fltActMil	= oGDM.fltActMil
tmpDeptAddr = oGDM.ApptAddr 	'= strDstAdr
tmpZipInst	= oGDM.ApptZIP		'= strDstZIP
tmpAvgSpd	= Z_FormatNumber(fltRealM / fltRealTT, 1)

%>
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
.summary th { text-align: center; }
.summary td { padding-top: 2px; padding-bottom: 3px; vertical-align: text-top; line-height: 1.0em;}
.summary input { height: 20px; padding: 2px 4px; border-radius: 2px; margin: 1px;}
.summary td input[type=checkbox] { height: 13px; padding: 1px; border-radius: 0px; }
.summary td:first-child	{ text-align: right; padding-right: 2px;}
.makered { color: red; }
.ico_view { display: inline-block; float: left; }
	</style>
</head>
<body>
<!-- #include file="_header.asp" -->
<div class="container">
	
	<div class="row" style="margin-top: 40px;">
		<div class="twelve columns">
			<h4 id="lblResults">Travel Time &amp; Mileage</h4>
<%
If Session("MSG") <> "" Then
	Response.Write "<div class=""err"">" & Session("MSG") & "</div>"
	Session("MSG") = ""
End If
%>
		</div><!-- columns -->
	</div><!-- row-->
<form action="ovrd_ttm_do.asp" method="post" name="frmA" id="frmA">
	<div class="row" style="margin-top: 10px;">
		<div class="one column">&nbsp;</div>
		<div class="five columns">
			<input type="hidden" name="ReqID" id="ReqID" value="<%= strReqID %>" />
			<input type="hidden" name="Postback" id="Postback" value="<%=strPostBack%>" />
			<table class="summary"><thead></thead>
				<tbody>
					<tr><td>Request ID:</td>
						<td><%= strReqID%>&nbsp;<%= tmpEmer%>&nbsp;<%= chkVer%></td>
						</tr>
<% If tmpHPID > 0 Then %>
					<tr><td>Vendor Site ID:</td>
						<td><%=tmpHPID%></td>
						</tr>	
<% End If %>
					<tr><td>Timestamp:</td>
						<td><%=TS%></td>
						</tr>
					<tr><td>Institution:</td>
						<td><%=tmpIname%> <%=tmpTrain%></td></tr>
					<tr><td>Department:</td>
						<td><%=tmpDname%></td></tr>
					<tr><td>Status:</td>
						<td><%=Statko%><br />
							<br />--- Current Status ---<br />
							<%=chkPaid%><br />
							<br />--- Trail ---<br />
							<i><%=reqTrail%></i><br /></td>
						</tr>
					<!-- tr><td>Mileage Rate:</td>
						<td>$ <%= tmpMrate %></td></tr>
					<tr><td>Travel Time Rate:</td>
						<td>$ <%= ttrate %></td></tr -->
					</tbody></table>
		</div>
		<div class="five columns">
			<table class="summary"><thead><tr><th colspan="2">&nbsp;</th></tr></thead>
				<tbody>
					<tr><td>Appointment Date:</td>
						<td><%=tmpAppDate%></td></tr>
					<tr><td>Appointment Time:</td>
						<td><%=tmpAppTFrom%> - <%=tmpAppTTo%></td></tr>
					<tr><td>Interpreter:</td>
						<td><%= tmpIntrName %></td></tr>
					<tr><td>Language:</td>
						<td><%=tmpSalita%></td></tr>
					<!-- tr><td>Interpreter Rate:</td>
						<td>$ <%= tmpIntrRate %></td></tr>
					<tr><td>Mileage Rate:</td>
						<td>$ <%= tmpmilerate %></td></tr -->
					<tr><td colspan="2" style="text-align: center; font-weight: bold;">&nbsp;Trip Information</td></tr>
					<tr><td>From:</td>
						<td>(Interpreter Residence)
							<!-- br /><%= tmpIntrZip%> --></td>
						</tr>
					<tr><td>Destination:</td>
						<td><%= tmpAddr %>&nbsp;</td>
						</tr>
					<!--tr><td>Interpreter Mileage Cap:</td>
						<td><%= tmpmilecap %></td></tr>
					<tr><td>Institution Mileage Cap:</td>
						<td><%= tmpMileCapInst %></td></tr -->
			</tbody></table>
		</div></div>
	<div class="row" style="margin-top: 10px;">
		<div class="one column">&nbsp;</div>
		<div class="seven columns">
			<table class="u-full-width summary"><thead>
					<tr><th colspan="3"	style="text-align: center;"><b style="font-size: 125%;">Transportation Time & Mileage Information</b></th></tr>
				</thead>
				<tbody>
					
					<tr><td>&nbsp;</td><td>&nbsp;&nbsp;&nbsp;Setting</td><td>Google Maps Values</td></tr>
					<tr><td>Driving Duration:</td>
							<td><input type="text" name="RealTT" id="RealTT" value="<%= tmpRealTT %>" /></td>
							<td><div name="txtRTravel" id="txtRTravel"><%= fltRealTT %></div></td>
							<td>hr</td>
						</tr>
					<tr><td>Distance:</td>
							<td><input type="text" name="RealM" id="RealM" value="<%= tmpRealM%>" /></td>
							<td><div name="txtRMile" id="txtRMile"><%= fltRealM %></div></td>
							<td>mile</td>
						</tr>
					<tr><td>Average Speed:</td><td>&nbsp;</td><td>
							<div name="txtRate" id="txtRate"><%= tmpAvgSpd %></div></td>
							<td>mph</td>
						</tr>
					<tr><td colspan="4" style="text-align: center; background-color: pink; font-size: 110%;">Interpreter (cap: <%= tmpmilecap %>)</td></tr>
					<tr><td>Billable Distance:</td>
							<td><input type="text" name="actMil" id="actMil" value="<%= tmpIntrMI %>" /></td>
							<td><div name="txtMile" id="txtMile"><%= fltActMil %></div></td>
							<td>mile</td>
						</tr>	
					<tr><td>Billable Travel Time:</td>
							<td><input type="text" name="actTT" id="actTT" value="<%= tmpIntrTT %>" /></td>
							<td><div name="txtTravel" id="txtTravel"><%= fltActTT %></div></td>
							<td>hr
							<input type="hidden" name="TT_Intr" id="TT_Intr" value="<%= tmpBilTIntr %>" />
							<input type="hidden" name="M_Intr" id="M_Intr" value="<%= tmpBilMIntr %>" />
								</td></tr>
					<tr><td>Tolls &amp; Parking:</td>
							<td colspan='2'>$<input type="text" maxlength="8" name="toll" id="toll" value='<%=tmpToll%>' style="width: 100px;" />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<!-- input type='checkbox' name='LBconfirmToll' id='LBconfirmToll' value='1' <%=tmpChkLBconfirmToll%> /> Approve -->
							</td></tr>
								
					<tr><td>&nbsp;</td>
							<td colspan="3"><input type="checkbox" name="payintr" id="payintr" value="1" <%= chkPayIntr %> />
								<div style="display: inline-block;">Pay to Interpreter</div>
								</td></tr>
					<!-- tr><td>Billed Distance Amt:</td>
							<td>$</td>
							<td colspan="2">&nbsp;</td></tr>
					<tr><td>Travel Time Bill Amt:</td>
							<td>$</td>
							<td colspan="2">&nbsp;</td></tr -->
					<tr><td colspan="4" style="text-align: center; background-color: pink; font-size: 110%;">Institution (cap: <%= tmpMileCapInst %>)</td></tr>
					<tr><td>Billable Mileage:</td>
							<td><input type="text" name="InstActMil" id="InstActMil" value="<%= tmpInstActMil %>"/></td>
							<td><div name="txtMileInst" id="txtMileInst"><%= fltActMil %></div></td>
							<td>mile</td></tr>
					<tr><td>Billable Travel Time:</td>
							<td><input type="text" name="InstActTT" id="InstActTT" value="<%= tmpInstActTT %>" /></td>
							<td><div name="txtTravelInst" id="txtTravelInst"><%= fltActTT %></div></td>
							<td>hr
								<input type="hidden" name="M_Inst" id="M_Inst" value="<%= tmpBilMInst %>" />
								<input type="hidden" name="TT_Inst" id="TT_Inst" value="<%= tmpBilTInst %>" />
							</td></tr>
					<tr><td>&nbsp;</td>
							<td colspan="3"><input type="checkbox" name="billinst" id="billinst" value="1" <%= chkBillInst %> />
								<div style="display: inline-block;">Bill Institution [ <%= tmpIname %> ]</div>
								</td></tr>
					<!-- tr><td>Billed Distance Amt:</td>
							<td>$</td>
							<td colspan="2">&nbsp;</td></tr>
					<tr><td>Travel Time Bill Amt:</td>
							<td>$</td>
							<td colspan="2">&nbsp;</td></tr -->
			</tbody></table>
		</div>
		<div class="four columns">
			<div style="height: 100px;">&nbsp;</div>
			<code>Billable Miles = Distance - (2 x Cap )</code>
			<br /><br />
			<code>Billable Travel Time = Billable Miles / Avg.Speed</code>
			<br />
		</div>
	</div>
	<div class="row">
		<div class="two columns">&nbsp;</div><div class="eight columns">
			<button id="btnSave" name="btnSave" class="button button-primary">SAVE</button>
			<!-- div style="display: inline-block"><b style="font-size: 150%;"><%= strPostBack %>
			</b></div -->
		</div>
	</div>
	</form>
</div><!-- container -->
</body>
</html>