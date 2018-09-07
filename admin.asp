<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmConfirm."
If Request("edits") = 1 Then
	Set dload = Server.CreateObject("SCUpload.Upload")
	dload.Download "C:\work\LSS-LBIS\log\editlogs.txt"
Set dload = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Admin page</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			<!--
		function CalendarView(strDate)
			{
				document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
				document.frmConfirm.submit();
			}
			function MyEdits()
		{
			document.frmConfirm.action = 'admin.asp?edits=1';
			document.frmConfirm.submit();
		}
		-->
			-->
			</script>
	</head>
	<body>
		<form method='post' name='frmConfirm' action='admin.asp'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0' align='center' >
							<!-- #include file="_greetme.asp" -->
							<tr><td>&nbsp;</td></tr>
							<!--
							<tr>
								<td align='center'>
									<a href='<%=tmphref%>' class='admin'>[Interpreter]</a>&nbsp;&nbsp;
									<input type='radio' name='RadioType' value='0' <%=myType0%> onclick='document.frmConfirm.submit();'>Active 
									<input type='radio' name='RadioType' value='1' <%=myType1%> onclick='document.frmConfirm.submit();'>Inactive
									<input type='radio' name='RadioType' value='2' <%=myType2%> onclick='document.frmConfirm.submit();'>All
									<%If tmpexpire Then%>
										&nbsp;<%=tmpWarn%>
									<%End If %> 	
									<br>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='admintools.asp' class='admin'>[Institutions/Departments/Requesting Persons]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='adminOthers.asp' class='admin'>[Rates/Mileage/Language/Reasons]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='adminusers.asp' class='admin'>[Users]</a>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							-->
							<tr>
								<td align='center'>
									<a href='nohours.asp' class='admin'>[Interpreter - No Appointment Hours Appointment]</a>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='center'>
									<a href='client.asp' class='admin'>[Client-Interpreter Preferred]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='cid.asp' class='admin'>[Department Billing Data]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='wid.asp' class='admin'>[Interpreter Billing Data]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='holiday.asp' class='admin'>[Holiday Dates]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='reqtable2.asp?ctrlX=1' class='admin'>[Timesheet]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='reqtable2.asp?ctrlX=2' class='admin'>[Mileage]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='reqtable4.asp' class='admin'>[Approve/Disaaprove Medicaid/MCO]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='upload271.asp' class='admin'>[Upload 271]</a>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<a href='reqtable3.asp?ctrlX=1' class='admin'>[Institution Billable Hours]</a>
								</td>
							</tr>
							<!--<tr>
								<td align='center'>
									<a href='reqtable3.asp?ctrlX=2' class='admin'>[Medicaid Billable Hours]</a>
								</td>
							</tr>//-->
							<tr>
								<td align='center'>
									<a href='adminreports.asp' class='admin'>[Admin Reports]</a>
								</td>
							</tr>
							<tr><td align="center"><a href="admin_foi.asp" class="admin" target="_BLANK">[Consent Forms]</a></td></tr>
							<tr><td align="center"><a href="survey.list.asp" class="admin" target="_BLANK">[Interpreter survey2018]</a></td></tr>
							<tr style="height: 30px;"><td>&nbsp;</td></tr>
							<tr><td align="center"><a href="rep.novform.asp" class="admin" >[Unsubmitted Verification Form Report]</a></td></tr>
							<tr style="height: 30px;"><td>&nbsp;</td></tr>
							<tr>
								<td align='center'>
									<a href="JavaScript: MyEdits();" class='admin'>[Download Edit Appointment log]</a>
								</td>
							</tr>
							<tr style="height: 200px;"><td>&nbsp;</td></tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign='bottom'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
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