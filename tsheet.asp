<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
	If Session("UIntr") = "" Then 
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
	Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
	tmpPage = "document.frmTS."
	'Get week range
	If Request("action") = 1 Then
		sundate = DateAdd("d", -7, GetSun(Request("tmpDate")))
		satdate = DateAdd("d", -7, GetSat(Request("tmpDate")))
	ElseIf Request("action") = 2 Then
		sundate = DateAdd("d", 7, GetSun(Request("tmpDate")))
		satdate = DateAdd("d", 7, GetSat(Request("tmpDate")))
	Else 
		sundate = GetSun(Request("tmpDate"))
		satdate = GetSat(Request("tmpDate"))
	End If
	Set rsTS = Server.CreateObject("ADODB.RecordSet")
	sqlTS = "SELECT overpayhrs, payhrs, LBconfirm, [index], totalhrs, Status, confirmed, InstID, Cfname, Clname, AStarttime, AEndtime, totalhrs, actTT, appDate FROM request_T WHERE appDate >= '" & sundate & "' AND appDate <= '" & satDate & "' AND IntrID = " & Session("UIntr") & " " & _
		"AND showintr = 1 AND status <> 2 AND Status <> 3 ORDER BY appDate ,appTimeFrom"
	rsTS.Open sqlTS, g_strCONN, 1, 3
	ctr = 0
	ctrCon = 0
	If Not rsTS.EOF Then
		Do Until rsTS.EOF
			tmpTot = rsTS("totalhrs")
			myStat = ""
			'If rsTS("Status") = 2 Or rsTS("Status") = 3 Or rsTS("Status") = 4 Or rsTS("confirmed") <> "" Then myStat = "DISABLED"
			'tmpAMT = "$" & Z_FormatNumber(AmtRate(rsTS("m_intr")), 2)
			TT = Z_FormatNumber(rsTS("actTT"), 2)
			If rsTS("overpayhrs") Then 
				PHrs = Z_FormatNumber(rsTS("payhrs"), 2)
			Else
				PHrs = Z_FormatNumber(IntrBillHrs(rsTS("AStarttime"), rsTS("AEndtime")), 2)
			End If
			FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
			TotFPHrs = TotFPHrs + FPHrs
			tmpCon = ""
			LBcon = ""
			If rsTS("LBconfirm") = True Then 
				tmpCon = "checked" '"<b>*</b>"
				LBcon = "readonly"
				ctrCon = ctrCon + 1
			End If
			IntrCon = ""
			If rsTS("confirmed") <> "" Then IntrCon = "disabled checked"
			myAct = rsTS("index") & " - " & GetInst(rsTS("InstID")) & " - " & left(rsTS("Cfname"), 1) & ". " & left(rsTS("Clname"), 1) & "."
			AStime = Z_FormatTime(Ctime(rsTS("AStarttime")))
			AEtime = Z_FormatTime(Ctime(rsTS("AEndtime")))

			If rsTS("appDate") = cdate(sundate) Then
				sunTS = sunTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
					"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					sunTS = sunTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				sunTS = sunTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(sundate) + 1 Then
				monTS = monTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					monTS = monTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				monTS = monTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(sundate) + 2 Then
				tueTS = tueTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					tueTS = tueTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				tueTS = tueTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(sundate) + 3 Then
				wedTS = wedTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					wedTS = wedTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				wedTS = wedTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(sundate) + 4 Then
				thuTS = thuTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main'size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					thuTS = thuTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				thuTS = thuTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(sundate) + 5 Then
				friTS = friTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					friTS = friTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				friTS = friTS &	"</tr>"
			End If
			If rsTS("appDate") = cdate(satdate) Then
				satTS = satTS & "<tr bgcolor='#F5F5F5'><td><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'></td><td align='center'><nobr>" & myAct & "</td>" & _
				 	"<td align='center'>" & TT & "</td>" & _
				 	"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunstart" & ctr & "' value='" & AStime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' " & myStat & " maxlength='5' name='sunend" & ctr & "' value='" & AEtime & "' onKeyUp=""javascript:return maskMe(this.value,this,'2',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"" " & LBcon & "></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly name='totalhrs" & ctr & "' value='" & rsTS("totalhrs") & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsx" & ctr & "' value='" & PHrs & "'><input type='hidden' name='hidpayhrs" & ctr & "' value='" & PHrs & "'></td>" & _
					"<td align='center'><input class='main' size='6' maxlength='11' readonly  name='totalhrsxx" & ctr & "' value='" & z_formatnumber(FPHrs, 2) & "'></td>" & _
					"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
				If tmpCon = "checked" Then
					satTS = satTS & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				End If
				satTS = satTS &	"</tr>"
			End If
			ctr = ctr + 1
			rsTS.MoveNext
		Loop
	End If
	rsTS.Close
	Set rsClose = Nothing
	
%>
<html>
	<head>
		<title>Language Bank - Interpreter Timesheet</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
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
		function SaveTS()
		{
			document.frmTS.action = "tsheetaction.asp?action=1";
			document.frmTS.submit();
	
		}
		function SendTS()
		{
			if (document.frmTS.myCTR.value != document.frmTS.myCTR2.value)
			{
				alert("All entries must be approved first before you can confirm.")
				return;
			}
			var ans = window.confirm("Confirm CHECKED Appointments? \nCancel to stop.");
			if (ans)
			{
				document.frmTS.action = "tsheetaction.asp?action=1&confirm=1";
				document.frmTS.submit();
			}
		}
		function CalendarView(strDate)
		{
			document.frmTS.action = 'tsheet.asp?tmpdate=' + strDate;
			document.frmTS.submit();
		}
		function PrevMonth()
		{
			document.frmTS.action = "tsheet.asp?action=1&tmpDate=" + '<%=sundate%>';
			document.frmTS.submit();	
		}
		function NextMonth()
		{
			document.frmTS.action = "tsheet.asp?action=2&tmpDate=" + '<%=sundate%>';
			document.frmTS.submit();	
		}
		//-->
		</script>
	</head>
	<body>
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter Timesheet</td>
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
									<td class='confirm'><%=GetIntr(Session("UIntr"))%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Date:</td>
									<td class='confirm'><%=sundate%> - <%=satdate%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<form name='frmTS' method='POST'>
											<table border='0' cellpadding='1' cellspacing='2' width='75%'>
												<tr>
													<td align='center' class='tblgrn'>Date</td>
													<td align='center' class='tblgrn'>Activity</td>
													<td align='center' class='tblgrn'>Travel Time</td>
													<td align='center' class='tblgrn'>Appt. Start Time</td>
													<td align='center' class='tblgrn'>Appt. End Time</td>
													<td align='center' class='tblgrn'>Total Hours</td>
													<td align='center' class='tblgrn'>Payable Hours</td>
													<td align='center' class='tblgrn'>Final Payable Hours</td>
													<td align='center' class='tblgrn'>Approved</td>
													<td align='center' class='tblgrn'>Confirmed</td>
												</tr>	
												<tr><td align='center' class='confirm'>SUN</td></tr>
												<%=sunTS%>
												<tr><td align='center' class='confirm'>MON</td></tr>
												<%=monTS%>
												<tr><td align='center' class='confirm'>TUE</td></tr>
												<%=tueTS%>
												<tr><td align='center' class='confirm'>WED</td></tr>
												<%=wedTS%>
												<tr><td align='center' class='confirm'>THU</td></tr>
												<%=thuTS%>
												<tr><td align='center' class='confirm'>FRI</td></tr>
												<%=friTS%>
												<tr><td align='center' class='confirm'>SAT</td></tr>
												<%=satTS%>
												<tr>
													<td colspan='6'>&nbsp;</td>
													<td align='center' class='main' bgcolor='#FFFFCE'>Total</td>
													<td align='center' class='main' bgcolor='#FFFFCE'><%=TotFPHrs%></td>
												</tr>
											</table>
										
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='12' align='center'>
										*To <b>SAVE</b> your Appointment Time, input Appt. Start time and App. End time then click on 'Save' button.<br>
										PLEASE USE MILITARY TIME (24-Hour FORMAT). Do not use 00:00 / midnight on both fields. If you need to enter it, use 00:01.<br>
										Once approved, you can no longer edit Appt. Start Time and Appt. End Time
										<br><br>
										*To <b>CONFIRM</b> your appointment Time, check the correspoding checkbox then click on 'Confirm' button.<br>
										Appointment time needs to be approved by Languagebank before you can confirm it.
									</td>
								</tr>
								<tr>
									<td colspan='12' align='center' height='100px' valign='bottom'>
										<input type='hidden' name='tmpDate' value="<%=sundate%>">
										<input type='hidden' name='myCTR' value="<%=ctr%>">
										<input type='hidden' name='myCTR2' value="<%=ctrCon%>">
										<input class='btn' type='button' value='<Prev Week'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrevMonth();'>
										<input class='btn' type='button' value='Save' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveTS();'>
										<input class='btn' type='button' value='Confirm' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SendTS();'>
										<input class='btn' type='button' value='Next Week>'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='NextMonth();'>
										
									</td>
								</tr>
								</form>
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