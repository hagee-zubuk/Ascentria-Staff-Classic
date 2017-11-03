<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Please sign-in as user with admin rights."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmAdmin."
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [index], [Language] FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpLang = Request("LangID")
	tmpL = ""
	If tmpLang = "" Then tmpLang = "0"
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	'LANGUAGE CHECKER
	strChkLang = strChkLang & "if (Ucase(document.frmAdmin.txtLang.value) == """ & UCase(rsLang("Language")) & """ && document.frmAdmin.selLang.value == 0) " & vbCrLf & _
		"{alert(""ERROR: Language already exists.""); return;}" & vbCrLf
	strChkLang2 = strChkLang2 & "if (Ucase(document.frmAdmin.txtLang.value) == """ & UCase(rsLang("Language")) & """ && document.frmAdmin.selLang.value != 0) " & vbCrLf & _
		"{alert(""ERROR: Language already exists.""); return;}" & vbCrLf
	strLangI = strLangI & "if (lng == " & rsLang("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtLang.value = """ & rsLang("Language") & """; " & vbCrLf & _
		"document.frmAdmin.LangName.value = """ & rsLang("Language") & """; }" & vbCrLf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET LANGUAGE INFO
'Set rsLangI = Server.CreateObject("ADODB.RecordSet")
'sqlLangI = "SELECT * FROM Language_T ORDER BY [Language]"
'rsLangI.Open sqlLangI, g_strCONN, 3, 1
'Do Until rsLangI.EOF
'	strLangI = strLangI & "if (lng == " & rsLangI("Index") & ") " & vbCrLf & _
'		"{document.frmAdmin.txtLang.value = """ & rsLangI("Language") & """; " & vbCrLf & _
'		"document.frmAdmin.LangName.value = """ & rsLangI("Language") & """; }" & vbCrLf
'	rsLangI.MoveNext
'Loop
'rsLangI.Close
'Set rsLangI = Nothing
'GET INSTITUTION RATES 
Set rsRates = Server.CreateObject("ADODB.RecordSet")
sqlRates = "SELECT Rate FROM rate_T ORDER BY Rate"
rsRates.Open sqlRates, g_strCONN, 3, 1
Do Until rsRates.EOF
	RateKo = ""
	If Request("RateID") = rsRates("Rate") Then RateKo = "selected"
	strRate1 = strRate1 & "<option " & RateKo & " value='" & rsRates("Rate") & "'>$" & Z_FormatNumber(rsRates("Rate"), 2) & "</option>" & vbCrLf
	strRate2 = strRate2 & "if (xxx == " & rsRates("Rate") & ") " & vbCrLf & _
		"{document.frmAdmin.txtRate.value = """ & rsRates("Rate") & """; " & vbCrLf & _
		"document.frmAdmin.RateReas.value = """ & rsRates("Rate") & """; }" & vbCrLf 
	rsRates.MoveNext
Loop
rsRates.Close
Set rsRates = Nothing
'GET INTERPRETER RATES
Set rsRates = Server.CreateObject("ADODB.RecordSet")
sqlRates = "SELECT Rate2 FROM rate2_T ORDER BY Rate2"
rsRates.Open sqlRates, g_strCONN, 3, 1
Do Until rsRates.EOF
	RateKo2 = ""
	If Request("RateID2") = rsRates("Rate2") Then RateKo2= "selected"
	strRateIntr = strRateIntr & "<option " & RateKo2 & " value='" & rsRates("Rate2") & "'>$" & Z_FormatNumber(rsRates("Rate2"), 2) & "</option>" & vbCrLf
	strRateIntr2 = strRateIntr2 & "if (xxx == " & rsRates("Rate2") & ") " & vbCrLf & _
		"{document.frmAdmin.txtRate2.value = """ & rsRates("Rate2") & """; " & vbCrLf & _
		"document.frmAdmin.RateReas2.value = """ & rsRates("Rate2") & """; }" & vbCrLf 
	rsRates.MoveNext
Loop
rsRates.Close
Set rsRates = Nothing
'MILEAGE CAP interpreter
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT milediff FROM travel_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpMileDiff = rsRate("milediff")
End If
rsRate.Close
Set rsRate = Nothing
'MILEAGE RATE 
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT mileagerate FROM mileagerate_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpMR = rsRate("mileagerate")
End If
rsRate.Close
Set rsRate = Nothing
'EMERGENCY RATE
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT FeeLegal, FeeOther FROM EmergencyFee_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpFeeL = rsRate("FeeLegal")
	tmpFeeO = rsRate("FeeOther")
End If
rsRate.Close
Set rsRate = Nothing
'MILEAGE RATE court only
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT milediffcourt FROM travelInstCourt_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpMileDiffcourt = rsRate("milediffcourt")
End If
rsRate.Close
Set rsRate = Nothing
'MILEAGE RATE court only
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT milediffinst FROM travelInst_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpMileDiffinst = rsRate("milediffinst")
End If
rsRate.Close
Set rsRate = Nothing
'GET CANCELLATION REASON
Set rsCancel = Server.CreateObject("ADODB.RecordSet")
sqlCancel = "SELECT * FROM cancel_T"
rsCancel.Open sqlCancel, g_strCONN, 3, 1
Do Until rsCancel.EOF
	CancelMe = ""
	tmpCancel = Request("ReasonID")
	If tmpCancel = rsCancel("index") Then CancelMe = "selected"
	strCancel = strCancel & "<option value='" & rsCancel("index") & "' " & CancelMe & ">" & rsCancel("Reason") & "</option>" & vbCrLf
	strCancelled = strCancelled & "if (xxx == " & rsCancel("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtCancel.value = """ & rsCancel("Reason") & """; " & vbCrLf & _
		"document.frmAdmin.CancelReas.value = """ & rsCancel("Reason") & """; }" & vbCrLf
	rsCancel.MoveNext
Loop
rsCancel.Close
Set rsCancel = Nothing
'GET MISSED REASON
Set rsMissed = Server.CreateObject("ADODB.RecordSet")
sqlMissed = "SELECT * FROM Missed_T"
rsMissed.Open sqlMissed, g_strCONN, 3, 1
Do Until rsMissed.EOF
	MissedMe = ""
	tmpMissed = Request("ReasonID1")
	If tmpMissed = rsMissed("index") Then MissedMe = "selected"
	strMissed = strMissed & "<option value='" & rsMissed("index") & "' " & MissedMe & ">" & rsMissed("Reason") & "</option>" & vbCrLf
	strMisses = strMisses & "if (xxx == " & rsMissed("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtMissed.value = """ & rsMissed("Reason") & """; " & vbCrLf & _
		"document.frmAdmin.MissedReas.value = """ & rsMissed("Reason") & """; }" & vbCrLf
	rsMissed.MoveNext
Loop
rsMissed.Close
Set rsMissed = Nothing
%>
<html>
	<head>
		<title>Language Bank - Administrator</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script type='text/javascript' language='JavaScript'>
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
		function LangInfo(lng)
		{
			if (lng == 0)
			{document.frmAdmin.txtLang.value = "";}
			<%=strLangI%>
		}
		function CancelInfo(xxx)
		{
			if (xxx == 0)
			{document.frmAdmin.txtCancel.value = "";}
			<%=strCancelled%>
		}
		function MissedInfo(xxx)
		{
			if (xxx == 0)
			{document.frmAdmin.txtMissed.value = "";}
			<%=strMisses%>
		}
		function RateInfo(xxx)
		{
			if (xxx == 0)
			{document.frmAdmin.txtRate.value = "";}
			<%=strRate2%>
		}
		function RateInfo2(xxx)
		{
			if (xxx == 0)
			{document.frmAdmin.txtRate2.value = "";}
			<%=strRateIntr2%>
		}
		function chkLang()
		{
			<%=strChkLang2%>
		}
		function Ucase(xxx)
		{
			var xxx;
			return xxx.toUpperCase();
			
		}
		function CalendarView(strDate)
		{
			document.frmAdmin.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAdmin.submit();
		}
		function KillMe()
		{
			var ans = window.confirm("DELETE?");
			if (ans)
			{
				document.frmAdmin.action = "adminothersaction.asp?ctrl=2";
				document.frmAdmin.submit();
			}
		}
		function SaveMe()
		{
			document.frmAdmin.action = "adminothersaction.asp?ctrl=1";
			document.frmAdmin.submit();
		}
		-->
		</script>
	</head>
	<body onload='LangInfo(document.frmAdmin.selLang.value);CancelInfo(document.frmAdmin.selCancel.value);
		MissedInfo(document.frmAdmin.selMissed.value); RateInfo(document.frmAdmin.selRate.value);
		 RateInfo2(document.frmAdmin.selRate2.value);'>
		<form method='post' name='frmAdmin'>
			<table cellSpacing='0' cellPadding='0'  width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top' colspan='2'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<!-- #include file="_greetme.asp" -->
				<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
				<tr>
					<td class='header' colspan='10'><nobr>Language</td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selLang'  style='width:150px;' onchange='LangInfo(document.frmAdmin.selLang.value);'>
							<option value='0'>&nbsp;</option>
							<%=strLang%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right' style='width:100px;'>Dialect:</td>
					<td>
						<input style='width:150px;' class='main' maxlength='50' name='txtLang' value='<%=tmpNewLang%>' onkeyup='bawal(this);'>
						<input type='hidden' name='LangName'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Institution Rates:</td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selRate'  style='width:75px;' onchange='RateInfo(document.frmAdmin.selRate.value);'>
							<option value='0'>&nbsp;</option>
							<%=strRate1%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right'>Rate:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='5' name='txtRate' value='<%=tmpNewRate%>' onkeyup='bawal(this);'>
						<input type='hidden' name='RateReas'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Interpreter Rates:</td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selRate2'  style='width:75px;' onchange='RateInfo2(document.frmAdmin.selRate2.value);'>
							<option value='0'>&nbsp;</option>
							<%=strRateIntr%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right'>Rate:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='5' name='txtRate2' value='<%=tmpNewRate2%>' onkeyup='bawal(this);'>
						<input type='hidden' name='RateReas2'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>Mileage Rate:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='5' name='txtMR' value='<%=tmpMR%>' onkeyup='bawal(this);'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Cancellation Reason</td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selCancel'  style='width:150px;' onchange='CancelInfo(document.frmAdmin.selCancel.value);'>
							<option value='0'>&nbsp;</option>
							<%=strCancel%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right'>Reason:</td>
					<td>
						<input style='width:150px;' class='main' maxlength='50' name='txtCancel' value='<%=tmpNewCancel%>' onkeyup='bawal(this);'>
						<input type='hidden' name='CancelReas'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Missed Reason</td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selMissed'  style='width:150px;' onchange='MissedInfo(document.frmAdmin.selMissed.value);' >
							<option value='0'>&nbsp;</option>
							<%=strMissed%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right'>Reason:</td>
					<td>
						<input style='width:150px;' class='main' maxlength='50' name='txtMissed' value='<%=tmpNewMissed%>' onkeyup='bawal(this);'>
						<input type='hidden' name='MissedReas'>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Interpreter's Mileage Cap</td>
				</tr>	
				<tr>
					<td align='right'>Cap:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='8' name='txtMile' value='<%=tmpMileDiff%>' onkeyup='bawal(this);'>
					</td>	
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Institution's Mileage Cap (Court and Legal only)</td>
				</tr>	
				<tr>
					<td align='right'>Cap:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='8' name='txtMileCourt' value='<%=tmpMileDiffcourt%>' onkeyup='bawal(this);'>
					</td>	
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Institution's Mileage Cap (Others)</td>
				</tr>	
				<tr>
					<td align='right'>Cap:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='8' name='txtMileInst' value='<%=tmpMileDiffinst%>' onkeyup='bawal(this);'>
					</td>	
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td class='header' colspan='10'><nobr>Emergency Fee</td>
				</tr>	
				<tr>
					<td align='right'>Legal & Court:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='8' name='txtFeel' value='<%=tmpFeeL%>' onkeyup='bawal(this);'>
					</td>	
				</tr>
				<tr>
					<td align='right'>Social Services, Medical & Private:</td>
					<td>
						<input style='width:75px;' class='main' maxlength='8' name='txtFeeO' value='<%=tmpFeeO%>' onkeyup='bawal(this);'>
					</td>	
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan='10' align='center' height='100px' valign='bottom'>
						<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe();">
						<input class='btn' type='button' style='width: 125px;' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe();'>
					</td>
				</tr>
				<tr>
					<td height='50px' valign='bottom' colspan='2'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
		</form>
	</body>
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