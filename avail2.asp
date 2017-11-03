<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
tmpPage = "document.frmTS."
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Set rsavail = Server.Createobject("ADODB.RecordSet")
	sqlavail = "SELECT * FROM avail_T WHERE intrID = " & Request("selIntr")
	rsavail.Open sqlAvail, g_strCONN, 1, 3
	If not rsavail.EOF Then
		x = 0
		Do Until rsavail.EOF
			NewInput = 0
			tmpDT = Split(rsavail("avail"), ",")
			
			tmpElement = tmpDT(0) & tmpDT(1) 
			
			strCHK = strCHK & "document.getElementsByName(" & tmpElement & ")[0].checked = true;" & vbCrLf
			
			x = x + 1
			rsavail.MoveNext
		Loop
	Else
		NewInput = 1	
	End If
	rsavail.Close
	set rsavail = Nothing
End If
'INTERPRETER LIST 
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T WHERE Active = 1 ORDER BY [Last Name], [First Name]"
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	Do Until rsIntr.EOF
		IntrName = rsIntr("Last Name") & ", " & rsIntr("First Name")
		myIntr = ""
		If cint(Request("selIntr")) = rsIntr("index") Then myIntr = "SELECTED"
		strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & myIntr & ">" & IntrName & "</option>" & vbCrlf
	rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Availability</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
			function checkALL()
			{
				<% If NewInput = 1 Then %>
					var tmpElem;
					var z;
					for(z = 0; z <= 23; z ++)
						{
							tmpElem = "1" + z;
							//alert(tmpElem);
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "2" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "3" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "4" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "5" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "6" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
							tmpElem = "7" + z;
							document.getElementsByName(tmpElem)[0].checked = true;
						}
				<% Else %> 
				<%=strChk %>		
				<% End If %>	
			}
			function CalendarView(strDate)
		{
			document.frmTS.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTS.submit();
		}
		
		//-->
		</script>
		<style>
			.myTime{
				width: 30px;
				text-align: center;
				font-weight: bold;
			}	
		</style>
	</head>
	<body onload="checkALL();">
		<form name='frmTS' method='POST'>
		<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
	<tr bgcolor='#336601'>
		<td class='head'><nobr>" Understand And Be Understood "</td>
		<td colspan='15' align='right'><a href='http://www.zubuk.com'><img src='images/zubuk-gear.gif' border='0' height='20px' width='20px' alt='Zubuk Inc.' title='Zubuk Inc.'></a></td>
	</tr>
	<tr>
		<td valign='top' align='left' style="width: 290px; height: 85px; background: url('images/LBISLOGO.jpg') no-repeat;"
			title="Language Bank">&nbsp;
		</td>
		<td colspan='15' valign='top' align='left' style="height: 85px; background: url('images/LBISLOGOside3.jpg') no-repeat;"
			title="Language Bank">&nbsp;
		</td>
	</tr>
	<tr bgcolor='#336601'>
		<td class='info'><nobr><a class='link' href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a> | <a class='link' href='http://www.thelanguagebank.org'>www.thelanguagebank.org</a></td>
		<td>&nbsp;</td>
	</tr>
</table>
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Availability</td>
								</tr>
								<tr>
									<td  align='center' colspan='12'>
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
									<td align='right' width='150px'>Name:</td>
									<td class='confirm'><select class='seltxt' name='selIntr' style='width: 150px;' onchange='document.frmTS.submit();' onblur='document.frmTS.submit();'>
							<option value='0'>&nbsp;</option>
							<%=strIntr%>
						</select></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										
											<table border='0' cellpadding='0' cellspacing='2'>
												<tr>
													<td>&nbsp;</td>
													<td class='myTime'>0:00</td>
													<td class='myTime'>1:00</td>
													<td class='myTime'>2:00</td>
													<td class='myTime'>3:00</td>
													<td class='myTime'>4:00</td>
													<td class='myTime'>5:00</td>
													<td class='myTime'>6:00</td>
													<td class='myTime'>7:00</td>
													<td class='myTime'>8:00</td>
													<td class='myTime'>9:00</td>
													<td class='myTime'>10:00</td>
													<td class='myTime'>11:00</td>
													<td class='myTime'>12:00</td>
													<td class='myTime'>13:00</td>
													<td class='myTime'>14:00</td>
													<td class='myTime'>15:00</td>
													<td class='myTime'>16:00</td>
													<td class='myTime'>17:00</td>
													<td class='myTime'>18:00</td>
													<td class='myTime'>19:00</td>
													<td class='myTime'>20:00</td>
													<td class='myTime'>21:00</td>
													<td class='myTime'>22:00</td>
													<td class='myTime'>23:00</td>
												</tr>
												<tr bgcolor='#FFFFCE'>
													<td align='right'>Sunday</td>
													<td align='center'><input type='checkbox' disabled name='10' value='1,0'></td>
													<td align='center'><input type='checkbox' disabled name='11' value='1,1'></td>
													<td align='center'><input type='checkbox' disabled name='12' value='1,2'></td>
													<td align='center'><input type='checkbox' disabled name='13' value='1,3'></td>
													<td align='center'><input type='checkbox' disabled name='14' value='1,4'></td>
													<td align='center'><input type='checkbox' disabled name='15' value='1,5'></td>
													<td align='center'><input type='checkbox' disabled name='16' value='1,6'></td>
													<td align='center'><input type='checkbox' disabled name='17' value='1,7'></td>
													<td align='center'><input type='checkbox' disabled name='18' value='1,8'></td>
													<td align='center'><input type='checkbox' disabled name='19' value='1,9'></td>
													<td align='center'><input type='checkbox' disabled name='110' value='1,10'></td>
													<td align='center'><input type='checkbox' disabled name='111' value='1,12'></td>
													<td align='center'><input type='checkbox' disabled name='112' value='1,13'></td>
													<td align='center'><input type='checkbox' disabled name='113' value='1,14'></td>
													<td align='center'><input type='checkbox' disabled name='114' value='1,15'></td>
													<td align='center'><input type='checkbox' disabled name='115' value='1,16'></td>
													<td align='center'><input type='checkbox' disabled name='116' value='1,17'></td>
													<td align='center'><input type='checkbox' disabled name='117' value='1,18'></td>
													<td align='center'><input type='checkbox' disabled name='118' value='1,20'></td>
													<td align='center'><input type='checkbox' disabled name='119' value='1,19'></td>
													<td align='center'><input type='checkbox' disabled name='120' value='1,20'></td>
													<td align='center'><input type='checkbox' disabled name='121' value='1,21'></td>
													<td align='center'><input type='checkbox' disabled name='122' value='1,22'></td>
													<td align='center'><input type='checkbox' disabled name='123' value='1,23'></td>
												</tr>
												<tr>
													<td align='right'>Monday</td>
													<td align='center'><input type='checkbox' disabled name='20' value='2,0'></td>
													<td align='center'><input type='checkbox' disabled name='21' value='2,1'></td>
													<td align='center'><input type='checkbox' disabled name='22' value='2,2'></td>
													<td align='center'><input type='checkbox' disabled name='23' value='2,3'></td>
													<td align='center'><input type='checkbox' disabled name='24' value='2,4'></td>
													<td align='center'><input type='checkbox' disabled name='25' value='2,5'></td>
													<td align='center'><input type='checkbox' disabled name='26' value='2,6'></td>
													<td align='center'><input type='checkbox' disabled name='27' value='2,7'></td>
													<td align='center'><input type='checkbox' disabled name='28' value='2,8'></td>
													<td align='center'><input type='checkbox' disabled name='29' value='2,9'></td>
													<td align='center'><input type='checkbox' disabled name='210' value='2,10'></td>
													<td align='center'><input type='checkbox' disabled name='211' value='2,11'></td>
													<td align='center'><input type='checkbox' disabled name='212' value='2,12'></td>
													<td align='center'><input type='checkbox' disabled name='213' value='2,13'></td>
													<td align='center'><input type='checkbox' disabled name='214' value='2,14'></td>
													<td align='center'><input type='checkbox' disabled name='215' value='2,15'></td>
													<td align='center'><input type='checkbox' disabled name='216' value='2,16'></td>
													<td align='center'><input type='checkbox' disabled name='217' value='2,17'></td>
													<td align='center'><input type='checkbox' disabled name='218' value='2,18'></td>
													<td align='center'><input type='checkbox' disabled name='219' value='2,19'></td>
													<td align='center'><input type='checkbox' disabled name='220' value='2,20'></td>
													<td align='center'><input type='checkbox' disabled name='221' value='2,21'></td>
													<td align='center'><input type='checkbox' disabled name='222' value='2,22'></td>
													<td align='center'><input type='checkbox' disabled name='223' value='2,23'></td>
												</tr>
												<tr bgcolor='#FFFFCE'>
													<td align='right'>Tuesday</td>
													<td align='center'><input type='checkbox' disabled name='30' value='3,0'></td>
													<td align='center'><input type='checkbox' disabled name='31' value='3,1'></td>
													<td align='center'><input type='checkbox' disabled name='32' value='3,2'></td>
													<td align='center'><input type='checkbox' disabled name='33' value='3,3'></td>
													<td align='center'><input type='checkbox' disabled name='34' value='3,4'></td>
													<td align='center'><input type='checkbox' disabled name='35' value='3,5'></td>
													<td align='center'><input type='checkbox' disabled name='36' value='3,6'></td>
													<td align='center'><input type='checkbox' disabled name='37' value='3,7'></td>
													<td align='center'><input type='checkbox' disabled name='38' value='3,8'></td>
													<td align='center'><input type='checkbox' disabled name='39' value='3,9'></td>
													<td align='center'><input type='checkbox' disabled name='310' value='3,10'></td>
													<td align='center'><input type='checkbox' disabled name='311' value='3,11'></td>
													<td align='center'><input type='checkbox' disabled name='312' value='3,12'></td>
													<td align='center'><input type='checkbox' disabled name='313' value='3,13'></td>
													<td align='center'><input type='checkbox' disabled name='314' value='3,14'></td>
													<td align='center'><input type='checkbox' disabled name='315' value='3,15'></td>
													<td align='center'><input type='checkbox' disabled name='316' value='3,16'></td>
													<td align='center'><input type='checkbox' disabled name='317' value='3,17'></td>
													<td align='center'><input type='checkbox' disabled name='318' value='3,18'></td>
													<td align='center'><input type='checkbox' disabled name='319' value='3,19'></td>
													<td align='center'><input type='checkbox' disabled name='320' value='3,20'></td>
													<td align='center'><input type='checkbox' disabled name='321' value='3,21'></td>
													<td align='center'><input type='checkbox' disabled name='322' value='3,22'></td>
													<td align='center'><input type='checkbox' disabled name='323' value='3,23'></td>
												</tr>
												<tr>
													<td align='right'>Wednesday</td>
													<td align='center'><input type='checkbox' disabled name='40' value='4,0'></td>
													<td align='center'><input type='checkbox' disabled name='41' value='4,1'></td>
													<td align='center'><input type='checkbox' disabled name='42' value='4,2'></td>
													<td align='center'><input type='checkbox' disabled name='43' value='4,3'></td>
													<td align='center'><input type='checkbox' disabled name='44' value='4,4'></td>
													<td align='center'><input type='checkbox' disabled name='45' value='4,5'></td>
													<td align='center'><input type='checkbox' disabled name='46' value='4,6'></td>
													<td align='center'><input type='checkbox' disabled name='47' value='4,7'></td>
													<td align='center'><input type='checkbox' disabled name='48' value='4,8'></td>
													<td align='center'><input type='checkbox' disabled name='49' value='4,9'></td>
													<td align='center'><input type='checkbox' disabled name='410' value='4,10'></td>
													<td align='center'><input type='checkbox' disabled name='411' value='4,11'></td>
													<td align='center'><input type='checkbox' disabled name='412' value='4,12'></td>
													<td align='center'><input type='checkbox' disabled name='413' value='4,13'></td>
													<td align='center'><input type='checkbox' disabled name='414' value='4,14'></td>
													<td align='center'><input type='checkbox' disabled name='415' value='4,15'></td>
													<td align='center'><input type='checkbox' disabled name='416' value='4,16'></td>
													<td align='center'><input type='checkbox' disabled name='417' value='4,17'></td>
													<td align='center'><input type='checkbox' disabled name='418' value='4,18'></td>
													<td align='center'><input type='checkbox' disabled name='419' value='4,19'></td>
													<td align='center'><input type='checkbox' disabled name='420' value='4,20'></td>
													<td align='center'><input type='checkbox' disabled name='421' value='4,21'></td>
													<td align='center'><input type='checkbox' disabled name='422' value='4,22'></td>
													<td align='center'><input type='checkbox' disabled name='423' value='4,23'></td>
												</tr>
												<tr bgcolor='#FFFFCE'>
													<td align='right'>Thursday</td>
													<td align='center'><input type='checkbox' disabled name='50' value='5,0'></td>
													<td align='center'><input type='checkbox' disabled name='51' value='5,1'></td>
													<td align='center'><input type='checkbox' disabled name='52' value='5,2'></td>
													<td align='center'><input type='checkbox' disabled name='53' value='5,3'></td>
													<td align='center'><input type='checkbox' disabled name='54' value='5,4'></td>
													<td align='center'><input type='checkbox' disabled name='55' value='5,5'></td>
													<td align='center'><input type='checkbox' disabled name='56' value='5,6'></td>
													<td align='center'><input type='checkbox' disabled name='57' value='5,7'></td>
													<td align='center'><input type='checkbox' disabled name='58' value='5,8'></td>
													<td align='center'><input type='checkbox' disabled name='59' value='5,9'></td>
													<td align='center'><input type='checkbox' disabled name='510' value='5,10'></td>
													<td align='center'><input type='checkbox' disabled name='511' value='5,11'></td>
													<td align='center'><input type='checkbox' disabled name='512' value='5,12'></td>
													<td align='center'><input type='checkbox' disabled name='513' value='5,13'></td>
													<td align='center'><input type='checkbox' disabled name='514' value='5,14'></td>
													<td align='center'><input type='checkbox' disabled name='515' value='5,15'></td>
													<td align='center'><input type='checkbox' disabled name='516' value='5,16'></td>
													<td align='center'><input type='checkbox' disabled name='517' value='5,17'></td>
													<td align='center'><input type='checkbox' disabled name='518' value='5,18'></td>
													<td align='center'><input type='checkbox' disabled name='519' value='5,19'></td>
													<td align='center'><input type='checkbox' disabled name='520' value='5,20'></td>
													<td align='center'><input type='checkbox' disabled name='521' value='5,21'></td>
													<td align='center'><input type='checkbox' disabled name='522' value='5,22'></td>
													<td align='center'><input type='checkbox' disabled name='523' value='5,23'></td>
												</tr>
												<tr>
													<td align='right'>Friday</td>
													<td align='center'><input type='checkbox' disabled name='60' value='6,0'></td>
													<td align='center'><input type='checkbox' disabled name='61' value='6,1'></td>
													<td align='center'><input type='checkbox' disabled name='62' value='6,2'></td>
													<td align='center'><input type='checkbox' disabled name='63' value='6,3'></td>
													<td align='center'><input type='checkbox' disabled name='64' value='6,4'></td>
													<td align='center'><input type='checkbox' disabled name='65' value='6,5'></td>
													<td align='center'><input type='checkbox' disabled name='66' value='6,6'></td>
													<td align='center'><input type='checkbox' disabled name='67' value='6,7'></td>
													<td align='center'><input type='checkbox' disabled name='68' value='6,8'></td>
													<td align='center'><input type='checkbox' disabled name='69' value='6,9'></td>
													<td align='center'><input type='checkbox' disabled name='610' value='6,10'></td>
													<td align='center'><input type='checkbox' disabled name='611' value='6,11'></td>
													<td align='center'><input type='checkbox' disabled name='612' value='6,12'></td>
													<td align='center'><input type='checkbox' disabled name='613' value='6,13'></td>
													<td align='center'><input type='checkbox' disabled name='614' value='6,14'></td>
													<td align='center'><input type='checkbox' disabled name='615' value='6,15'></td>
													<td align='center'><input type='checkbox' disabled name='616' value='6,16'></td>
													<td align='center'><input type='checkbox' disabled name='617' value='6,17'></td>
													<td align='center'><input type='checkbox' disabled name='618' value='6,18'></td>
													<td align='center'><input type='checkbox' disabled name='619' value='6,19'></td>
													<td align='center'><input type='checkbox' disabled name='620' value='6,20'></td>
													<td align='center'><input type='checkbox' disabled name='621' value='6,21'></td>
													<td align='center'><input type='checkbox' disabled name='622' value='6,22'></td>
													<td align='center'><input type='checkbox' disabled name='623' value='6,23'></td>
												</tr>
												<tr bgcolor='#FFFFCE'>
													<td align='right'>Saturday</td>
													<td align='center'><input type='checkbox' disabled name='70' value='7,0'></td>
													<td align='center'><input type='checkbox' disabled name='71' value='7,1'></td>
													<td align='center'><input type='checkbox' disabled name='72' value='7,2'></td>
													<td align='center'><input type='checkbox' disabled name='73' value='7,3'></td>
													<td align='center'><input type='checkbox' disabled name='74' value='7,4'></td>
													<td align='center'><input type='checkbox' disabled name='75' value='7,5'></td>
													<td align='center'><input type='checkbox' disabled name='76' value='7,6'></td>
													<td align='center'><input type='checkbox' disabled name='77' value='7,7'></td>
													<td align='center'><input type='checkbox' disabled name='78' value='7,8'></td>
													<td align='center'><input type='checkbox' disabled name='79' value='7,9'></td>
													<td align='center'><input type='checkbox' disabled name='710' value='7,10'></td>
													<td align='center'><input type='checkbox' disabled name='711' value='7,11'></td>
													<td align='center'><input type='checkbox' disabled name='712' value='7,12'></td>
													<td align='center'><input type='checkbox' disabled name='713' value='7,13'></td>
													<td align='center'><input type='checkbox' disabled name='714' value='7,14'></td>
													<td align='center'><input type='checkbox' disabled name='715' value='7,15'></td>
													<td align='center'><input type='checkbox' disabled name='716' value='7,16'></td>
													<td align='center'><input type='checkbox' disabled name='717' value='7,17'></td>
													<td align='center'><input type='checkbox' disabled name='718' value='7,18'></td>
													<td align='center'><input type='checkbox' disabled name='719' value='7,19'></td>
													<td align='center'><input type='checkbox' disabled name='720' value='7,20'></td>
													<td align='center'><input type='checkbox' disabled name='721' value='7,21'></td>
													<td align='center'><input type='checkbox' disabled name='722' value='7,22'></td>
													<td align='center'><input type='checkbox' disabled name='723' value='7,23'></td>
												</tr>
											</table>
										
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
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
<!-- #include file="_closeSQL.asp" -->
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