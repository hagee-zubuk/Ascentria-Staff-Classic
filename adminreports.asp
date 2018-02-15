<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
tmpPage = "document.frmReport."
'GET INTERPRETER
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = 1 ORDER BY [Last Name], [First Name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	tmpSel = ""
	If tmpIntr = rsIntr("index") Then tmpSel = "selected"
	strIntr = strIntr & "<option " & tmpSel & " value='" & rsIntr("index") & "'>" & rsIntr("Last Name") & ", " & rsIntr("First Name") & "</option>" & vbCrLf 
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsInt = Nothing

strUsers = "<option value=""-1"">-- all --</option>" & vbCrLf
strSQL = "SELECT [index], [lname], [fname], [username] FROM [user_T] WHERE [type]<>2 ORDER BY [fname] ASC"
Set rsUser = Server.CreateObject("ADODB.RecordSet")
rsUser.Open strSQL, g_strCONN, 3, 1
Do Until rsUser.EOF
	strUsers = strUsers & "<option value=""" & rsUser("index") & """>" & rsUser("fname") & " " & rsUser("lname") & "</value>" & vbCrLf
	rsUser.MoveNext
Loop
rsUser.Close
Set rsUser = Nothing
%>
<html>
	<head>
		<title>Language Bank - Admin Reports</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
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
		function RepGen(xxx,yyy, tmpfrom, tmpto, intrID)
		{
			if (xxx == 7 || xxx == 12 || xxx == 13 || xxx == 15 || xxx == 16) {
				if (tmpfrom == "") {
					alert("Date (from:) needed for this report.");
					return;
				}
				if (tmpto == "") {
					alert("Date (to:) needed for this report.");
					return;
				}
			}
			if (xxx == 10) {
				if (intrID == 0) {
					alert("Interpreter needed for this report.");
					return;
				}
			}
			if (xxx == 14) {
				if (tmpto == "") {
					alert("Date (to:) needed for this report.");
					return;
				}
			}
			//if (xxx == 15) {
			//	if (intrID == 0) {
			//		alert("Interpreter needed for this report.");
			//		return;
			//	}
			//}
			if (yyy == "")
			{
				yyy = 0;
			}
			newwindow = window.open('Intrreports.asp?ctrl=2' + '&selRep=' + xxx + 
					'&txtyear=' + yyy + '&txtRepFrom=' + tmpfrom + 
					'&selUser=' + document.frmReport.selUser.value +
					'&txtRepTo=' + tmpto + '&selIntr=' + intrID,
					'','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			//document.frmReport.action = "Intrreports.asp?ctrl=2"
			//document.frmReport.submit();
		}
		function CriSel(xxx)
		{
			document.frmReport.txtyear.disabled = true;
			document.frmReport.txtyear.value = "";
			document.frmReport.cal1.disabled = true;
			document.frmReport.cal2.disabled = true;
			document.frmReport.txtRepFrom.disabled = true;
			document.frmReport.txtRepTo.disabled = true;
			document.frmReport.txtRepFrom.value = "";
			document.frmReport.txtRepTo.value = "";
			document.frmReport.selUser.disabled = true;
			document.frmReport.selIntr.disabled = true;
			document.frmReport.selIntr.value = 0;
			if (xxx == 1)
			{
				document.frmReport.txtyear.disabled = false;
			}
			if (xxx == 2)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selIntr.disabled = false;
			}
			if (xxx == 4 || xxx == 6)
			{
				document.frmReport.txtyear.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
			} 
			if (xxx == 7 || xxx == 16)
			{
				document.frmReport.txtyear.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.selUser.disabled = false;
			} 
			if (xxx == 10 || xxx == 11 || xxx == 12 || xxx == 13 || xxx == 15)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				if (xxx == 10 || xxx == 15) {
					document.frmReport.selIntr.disabled = false;
				}
			} 
			if (xxx == 14) 
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selIntr.disabled = true;
			}
		}
		function CalendarView(strDate)
			{
				document.frmReport.action = 'calendarview2.asp?appDate=' + strDate;
				document.frmReport.submit();
			}
		</script>
	</head>
	<body onload='CriSel(0);'>
		<form method='post' name='frmReport' action='Intreports.asp?ctrl=2'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" class='bgstyle2' border='0'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<!-- #include file="_greetme.asp" -->
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='4' cellPadding='0' align='center' border='0' bgcolor='#FBEEB7'>
								<tr>
									<td colspan='2' align='center'>
										<b>Admin Report Query</b>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Type:
									</td>
									<td>
										<select class='seltxt' name='selRep'  style='width:200px;' onchange='CriSel(document.frmReport.selRep.value);'>
											<option value='0'>&nbsp;</option>
											<option value='9' <%=TypeSel9%>>Client Import Report</option>
											<option value='12' <%=TypeSel12%>>DHHS Survey</option>
											<option value='10' <%=TypeSel10%>>Interpreter Answers</option>
											<option value='11' <%=TypeSel11%>>Interpreter Answers per Appointment</option>
											<option value='14' <%=TypeSel14%>>Interpreter Average Hours (3 month period)</option>
											<option value='15' <%=TypeSel15%>>Interpreter Comment</option>
											<option value='4' <%=TypeSel4%>>Interpreter Date of Hire</option>
											<option value='6' <%=TypeSel6%>>Interpreter Date of Termination</option>
											<option value='3' <%=TypeSel3%>>Interpreter Documents</option>
											<option value='5' <%=TypeSel5%>>Interpreter Driving and Criminal Record</option>
											<option value='2' <%=TypeSel2%>>Interpreter Evaluation/Feedback</option>
											<option value='13' <%=TypeSel3%>>Interpreter Feedback 2</option>
											<option value='1' <%=TypeSel1%>>Interpreter Training</option>
											<option value='8' <%=TypeSel8%>>Public Defender Report</option>
											<option value='16' <%=TypeSel16%>>User Assign</option>
											<option value='7' <%=TypeSel7%>>User Record</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Year:
									</td>
									<td>
										<input class='main' size='5' maxlength='4' name='txtyear' value='' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">yyyy</span>
									</td>
								</tr>
								<tr>
									<td align='right'>Timeframe:</td>
									<td>
										&nbsp;From:<input class='main' size='10' maxlength='10' name='txtRepFrom' readonly value='<%=tmpRepFrom%>'>
										<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepFrom);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
										&nbsp;To:<input class='main' size='10' maxlength='10' name='txtRepTo' readonly value='<%=tmpRepTo%>'>
										<input type="button" value="..." title='Calendar' name="cal2" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepTo);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
									</td>
								</tr>
								<tr>
									<td align='right'>
										Assigners:
									</td>
									<td>
										<select class='seltxt' name='selUser' id="selUser"  style='width:200px;' onchange=''>
											<!-- option value='0'>&nbsp;</option -->
											<%=strUsers%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Interpreter:
									</td>
									<td>
										<select class='seltxt' name='selIntr'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strIntr%>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input class='btn' type='button' style='width: 200px;' value='Generate' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='RepGen(document.frmReport.selRep.value, document.frmReport.txtyear.value, document.frmReport.txtRepFrom.value, document.frmReport.txtRepTo.value, document.frmReport.selIntr.value);'>
										<input type='hidden' name='hideID'>
									</td>
								</tr>
								<tr>
									<td colspan='2' align='center'>
										<span class='error'><%=Session("MSG")%></span>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
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