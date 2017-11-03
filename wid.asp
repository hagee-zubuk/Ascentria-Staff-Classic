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

Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	myIntr = ""
	If request("IntrID") <> "" Then
		If Z_CZero(request("IntrID")) = rsInst("index") Then myInst = "SELECTED"
	End If
	strIntr = strIntr & "<option value='" & rsInst("index") & "' " & myInst & ">" & rsInst("last name") & ", " & rsInst("first name") & "</option>" & vbCrLf
	
	strIntr2 = strIntr2 & "if (document.frmHoliday.selIntr.value == " & rsInst("index") & ")" & vbCrLf & _
			"{document.frmHoliday.txtpid.value = '" & rsInst("PID") & "';" & vbCrLf & _
			"document.frmHoliday.txtxid.value = '" & rsInst("XID") & "';" & vbCrLf & _
			"document.frmHoliday.txtwid.value = '" & rsInst("WID") & "';" & vbCrLf & _
			"}" & vbCrLf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'LIST NO CID
Set rsCID = Server.CreateObject("ADODB.RecordSet")
sqlCID = "SELECT [index] as myIntrID, [last name], [first name] FROM interpreter_T WHERE Active = 1 AND ((PID IS NULL OR PID = '') " & _
	"OR (WID IS NULL OR WID = '') OR (XID IS NULL OR XID = '')) ORDER BY [last name], [first name]"
rsCID.Open sqlCID, g_strCONN, 3, 1
y = 0
If Not rsCID.EOF Then
	Do Until rsCID.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		strCID2 = strCID2 & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsCID("myIntrID") & "</td><td class='tblgrn2'><nobr>" & rsCID("last name") & ", " & rsCID("first name") & "</td>" & vbCrLf
		'strCSV = strCSV & rsCID("myIntrID") & "," & rsCID("last name") & ", " & rsCID("first name") & "<br>"
		rsCID.MoveNext
		y = y + 1
	Loop
Else
	strCID2 = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
End If
rsCID.CLose
SEt rsCID = Nothing

%>
<html>
	<head>
		<title>Language Bank - Admin page - Interpreter Billing Data</title>
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
		
		function SaveMe()
		{
			document.frmHoliday.action = "widaction.asp?";
			document.frmHoliday.submit();
		}
		function getData(xxx)
		{
			<%=strIntr2%>
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmHoliday' action='wid.asp'>
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
									<tr><td colspan='4' align='left'><u><b>Interpreter</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<select name='selIntr' class='seltxt' onchange="getData(this.value);" onblur="getData(this.value);">
												<option value='0'>&nbsp;</option>
											<%=strIntr%>
											</select>
										</td>
									</tr>	
									<tr><td colspan='4' align='left'><u><b>HP Provider ID</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtpid'>
										</td>
									</tr>	
									<tr><td colspan='4' align='left'><u><b>Xerox Provider ID</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtxid'>
										</td>
									</tr>	
									<tr><td colspan='4' align='left'><u><b>Worker ID</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtwid'>
										</td>
									</tr>	
								</table>
							</tr>
							<tr>
								<td align='center'>
									<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe()">
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='center'>
									<table cellSpacing='2' cellPadding='0' border='0' align='center'>
										<tr><td colspan='2' align='center'><b><u>Interpreters with Missing Data</u></b></td></tr>
										<tr>
											<td class='tblgrn'>ID</td>
											<td class='tblgrn'>Interpreter</td>
										</tr>
										<%=strCID2%>
									</table>
								</td>
							</tr>
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