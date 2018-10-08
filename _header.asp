<%
'header of web page - includes menu
%>
<table cellSpacing='0' cellPadding='0' width="100%" border='0' align="center">
	<tr>
		<td valign='top' align="left" rowspan="2" width="75%" height="65px" colspan="18">
			<img src='images/LBISLOGO.jpg' border="0">
		</td>
		<td align="center" width="25%" class="tollnum">
		Toll-Free 844.579.0610
		</td>
	</tr>
	<tr><td>&nbsp;</td></tr>	
	<tr bgcolor='#f68328'>
		<td class="motto" align="center">
			<nobr>Understand and Be Understood.</nobr>
		</td>
		<% If Request.Cookies("LBUSERTYPE") <> 2 Then %>
			<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 And Cint(Request.Cookies("LBUSERTYPE")) <> 5 Then %>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
				<td align='center' width='90px'><a href='wMain1.asp' class='link2'>New Request</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>
				<td align='center' width='90px'><a href='openappts.asp' class='link2'><nobr>Open Appointments</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>  
				<td align='center' width='90px'><a href='reqtable.asp' class='link2'>List</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
				<td align='center' width='90px'><a href='client.asp' class='link2'>Intr. Pref'd</a></td>
			<% End If %>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td>
			<% If Cint(Request.Cookies("LBUSERTYPE")) <> 5 Then %>
				<% If CalendarPage = False  Then %>
					<td align='center' width='90px'>
						<input name='CalCal' style='height: 0px; width: 0px; border: none; background-color: #71C4EE;' onfocus='CalendarView(this.value);'>
						<a name='calLink' href='JavaScript:showCalendarControl(<%=tmpPage%>CalCal);' class='link2'>Calendar</a>
					</td>
				<% Else %>
					<td align='center' width='90px'><a href='calendarview2.asp' class='link2'>Calendar</a></td>
				<% End If %>
			<% End If %>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<% If Cint(Request.Cookies("LBUSERTYPE")) = 1 Or Cint(Request.Cookies("LBUSERTYPE")) = 5 Then %>
				<td align='center' width='90px'><a href='reports.asp' class='link2'>Reports</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<% End If %>
			<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
				<td align='center' width='90px'><a href='admin.asp' class='link2'>Admin</a></td>
				<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<% End If %>
		<% Else %>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='90px'><a href='calendarview2.asp' class='link2'>Calendar</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
		<% End If %>
		<% If Cint(Request.Cookies("LBUSERTYPE")) <> 5 Then %>
			<td align='center' width='90px'><a href='avail2.asp' class='link2' target="_BLANK">Intr. Availability</a></td>
			<td align='center' class='head' width='10px'>&nbsp;|&nbsp;</td> 
		<% End If %>
		<td align='right'><a href='default.asp?chk=1' class='link2'>Sign Out</a>&nbsp;</td>
	</tr>
</table>