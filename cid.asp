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
sqlInst = "SELECT * FROM institution_T ORDER BY facility"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	myInst = ""
	If request("InstID") <> "" Then
		If Z_CZero(request("InstID")) = rsInst("index") Then myInst = "SELECTED"
	End If
	strInst = strInst & "<option value='" & rsInst("index") & "' " & myInst & ">" & rsInst("facility") & "</option>" & vbCrLf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
If request("InstID") <> "" Then
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_T WHERE InstID = " & request("InstID") & " ORDER BY dept"
	rsInst.Open sqlInst, g_strCONN, 3, 1
	Do Until rsInst.EOF
		myDept = ""
		If request("DeptID") <> "" Then
			If Z_CZero(request("DeptID")) = rsInst("index") Then myDept = "SELECTED"
		End If
		tmpDeptlbl = rsInst("dept") & " (" & GetClass(rsInst("class")) & ")"
		strDept = strDept & "<option value='" & rsInst("index") & "' " & myDept & ">" & tmpDeptlbl & "</option>" & vbCrLf
		strCID = strCID & "if (document.frmHoliday.selDept.value == " & rsInst("index") & ")" & vbCrLf & _
			"{document.frmHoliday.txtcid.value = '" & rsInst("custID") & "';" & vbCrLf & _
			"document.frmHoliday.txtbg.value = '" & rsInst("billgroup") & "';" & vbCrLf & _
			"document.frmHoliday.txtdc.value = '" & rsInst("distcode") & "';" & vbCrLf & _
			"}" & vbCrLf
		rsInst.MoveNext
	Loop
	rsInst.Close
	Set rsInst = Nothing
End If
'LIST NO CID
Set rsCID = Server.CreateObject("ADODB.RecordSet")
sqlCID = "SELECT Facility, dept FROM Institution_T, dept_T WHERE InstID = Institution_T.[index] AND ((CustID IS NULL OR CustID = '') " & _
	"OR (billgroup IS NULL OR billgroup = '') OR (distcode IS NULL OR distcode = '')) AND SageActive = 1 ORDER BY Facility, dept"
rsCID.Open sqlCID, g_strCONN, 3, 1
y = 0
If Not rsCID.EOF Then
	Do Until rsCID.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		strCID2 = strCID2 & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsCID("facility") & "</td><td class='tblgrn2'><nobr>" & rsCID("dept") & "</td>" & vbCrLf
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
		<title>Language Bank - Admin page - Department Billing Data</title>
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
			document.frmHoliday.action = "cidaction.asp?";
			document.frmHoliday.submit();
		}
		function getDept(xxx)
		{
			document.frmHoliday.action = "cid.asp?InstID=" + xxx;
			document.frmHoliday.submit();
		}
		function getCID(xxx)
		{
			<%=strCID%>
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmHoliday' action='cid.asp'>
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
									<tr><td colspan='4' align='left'><u><b>Institution</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<select name='selInst' class='seltxt' onchange="getDept(this.value);" onblur="getDept(this.value);">
												<option value='0'>&nbsp;</option>
											<%=strInst%>
											</select>
										</td>
									</tr>	
									<tr><td colspan='4' align='left'><u><b>Department</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<select name='selDept' class='seltxt' onchange="getCID(this.value);" onblur="getCID(this.value);" onclick="getCID(this.value);">
												<option value='0'>&nbsp;</option>
											<%=strDept%>
											</select>
										</td>
									</tr>	
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='4' align='left'><u><b>Customer ID</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtcid'>
										</td>
									</tr>	
									<tr><td colspan='4' align='left'><u><b>Billing Group</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtbg'>
										</td>
									</tr>	
										<tr><td colspan='4' align='left'><u><b>Distribution Code</b></u></td></tr>
									<tr>
										<td>&nbsp</td>
										<td>
											<input type='textbox' class='main' size='50' maxlength='50' name='txtdc'>
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
										<tr><td colspan='2' align='center'><b><u>Departments with Missing Data</u></b></td></tr>
										<tr>
											<td class='tblgrn'>Instiution</td>
											<td class='tblgrn'>Department</td>
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