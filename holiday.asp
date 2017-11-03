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
tmpPage = "document.frmHoliday."

Set rsHol = Server.CreateObject("ADODB.RecordSet")
sqlHol = "SELECT * FROM holiday_T ORDER BY holdate"
rsHol.Open sqlHol, g_strCONN, 3, 1
y = 0
Do Until rsHol.EOF
	kulay = "#FFFFFF"
	If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	strHol = strHol & "<tr bgcolor='" & kulay & "'><td class='tblgrn2' align='center'><input type='checkbox' name='chk" & y & "' value=" & rsHol("index") & "></td>" & vbCrLf & _
		"<td class='tblgrn2' align='left'>" & rsHol("holdate") & "</td></tr>" & vbCrLf
	y = y + 1
	rsHol.MoveNext
Loop
%>
<html>
	<head>
		<title>Language Bank - Admin page - Holiday Dates</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CalendarView(strDate)
		{
			document.frmHoliday.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmHoliday.submit();
		}
		function DelMe()
		{
			var ans = window.confirm("Delete Holiday?");
			if (ans)
			{
				document.frmHoliday.action = "holidayaction.asp?ctrl=2";
				document.frmHoliday.submit();
			}
		}
		function SaveMe()
		{
			document.frmHoliday.action = "holidayaction.asp?ctrl=1";
			document.frmHoliday.submit();
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmHoliday' action='holiday.asp'>
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

							<tr>
								<table cellSpacing='2' cellPadding='0' border='0' align='center'>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='4' align='center'><span class='error'><%=Session("MSG")%></span></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='4' align='center'><u><b>HOLIDAY DATES</b></u></td></tr>
									<%=strHol%>
									<tr><td class='foot' colspan='4'>&nbsp;</td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<select name='selMonth1' class='seltxt' style='width: 50px;'>
												<option value='0'>&nbsp;</option>
												<option value='1'>Jan</option>
												<option value='2'>Feb</option>
												<option value='3'>Mar</option>
												<option value='4'>Apl</option>
												<option value='5'>May</option>
												<option value='6'>Jun</option>
												<option value='7'>Jul</option>
												<option value='8'>Aug</option>
												<option value='9'>Sep</option>
												<option value='10'>Oct</option>
												<option value='11'>Nov</option>
												<option value='12'>Dec</option>
											</select>
											&nbsp;/&nbsp;
											<select name='selDay1' class='seltxt' style='width: 50px;'>
												<option value='0'>&nbsp;</option>
												<option value='1'>1</option>
												<option value='2'>2</option>
												<option value='3'>3</option>
												<option value='4'>4</option>
												<option value='5'>5</option>
												<option value='6'>6</option>
												<option value='7'>7</option>
												<option value='8'>8</option>
												<option value='9'>9</option>
												<option value='10'>10</option>
												<option value='11'>11</option>
												<option value='12'>12</option>
												<option value='13'>13</option>
												<option value='14'>14</option>
												<option value='15'>15</option>
												<option value='16'>16</option>
												<option value='17'>17</option>
												<option value='18'>18</option>
												<option value='19'>19</option>
												<option value='20'>20</option>
												<option value='21'>21</option>
												<option value='22'>22</option>
												<option value='23'>23</option>
												<option value='24'>24</option>
												<option value='25'>25</option>
												<option value='26'>26</option>
												<option value='27'>27</option>
												<option value='28'>28</option>
												<option value='29'>29</option>
												<option value='30'>30</option>
												<option value='31'>31</option>
											</select>
											&nbsp;/&nbsp;<input type='text' name='txtYear' class='main' style='width: 50px;' maxlength='4'>
										</td>
									</tr>
								</table>
							</tr>
							<tr>
								<td align='center'>
									<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe()">
									<input class='btn' type='button' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="DelMe()">
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
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
			<input type='hidden' name='ctr' value='<%=y%>'>
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