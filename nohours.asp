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
tmpPage = "document.frmnohours."
'get interpreter list
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT [index] as myIntrID, [last name], [first name] FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	strIntr = strIntr & "<option value='" & rsIntr("myIntrID") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbcrlf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
'GET DEPT INFO
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT [index], [dept] FROM dept_T WHERE Active = 1 AND InstID = 479 ORDER BY dept"
rsDept.Open sqlDept, g_strCONN, 3, 1
Do Until rsDept.EOF
		DeptName = rsDept("Dept")
		strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept("Index") & "'>" &  DeptName & "</option>" & vbCrlf
		strRP = strRP & "if (deptID == " & rsDept("index") & ") {" & vbCrlf
		Set rsRP = Server.CreateObject("ADODB.RecordSet")
		sqlRP = "SELECT requester_T.[index] as RPID, lname, fname FROM requester_T, reqdept_T WHERE deptID = " & _
				rsDept("index") & " AND reqID = requester_T.[index] ORDER BY lname, fname"
		rsRP.Open sqlRP, g_strCONN, 3, 1
		Do Until rsRP.EOF
			tmpRPname = rsRP("lname") & ", " & rsRP("fname")
			strRP = strRP & "{var ChoiceRP = document.createElement('option');" & vbCrLf & _
					"ChoiceRP.value = " & rsRP("RPID") & ";" & vbCrLf & _
					"ChoiceRP.appendChild(document.createTextNode(""" & tmpRPname & """));" & vbCrLf & _
					"document.frmnohours.selRP.appendChild(ChoiceRP);}" & vbCrLf
			rsRP.MoveNext
		Loop
		rsRP.Close
		Set rsRP = Nothing
		strRP = strRP & "}" & vbCrLf
	rsDept.MoveNext
Loop
rsDept.Close
Set rsDept = Nothing
%>
<html>
	<head>
		<title>Language Bank - Admin page - No Appointment Hours Appointment</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
<script language='JavaScript'><!--
	function CalendarView(strDate) {
		document.frmnohours.action = 'calendarview2.asp?appDate=' + strDate;
		document.frmnohours.submit();
	}
	function RPList(deptID) {
		document.frmnohours.selRP.options.length = 0;
		<%=strRP%>
	}
	function maskMe(str,textbox,loc,delim) {
		var locs = loc.split(',');
		for (var i = 0; i <= locs.length; i++) {
			for (var k = 0; k <= str.length; k++) {
				if (k == locs[i]) {
					if (str.substring(k, k+1) != delim) {
					 	str = str.substring(0,k) + delim + str.substring(k,str.length);
		     		}
				}
			}
		}
		textbox.value = str
	}
	function RTrim(str) {
		var whitespace = new String(" \t\n\r");
		var s = new String(str);
		if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
        	var i = s.length - 1;       
            while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
					i--;
			s = s.substring(0, i+1);
		}
		return s;
    }
    function LTrim(str) {
		var whitespace = new String(" \t\n\r");
		var s = new String(str);
		if (whitespace.indexOf(s.charAt(0)) != -1) {
			var j=0, i = s.length;
			while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
					j++;
			s = s.substring(j, i);
		}
		return s;
    }
    function Trim(str) {
		return RTrim(LTrim(str));
    }
	function SaveNoHours() {
		if (document.frmnohours.selDept.value == 0) {
			alert("ERROR: Department is Required.");
			return;
		}
		if (document.frmnohours.selRP.value == 0) {
			alert("ERROR: Requesting Person is Required.");
			return;
		}
		if (document.frmnohours.txtAppDate.value == '') {
			alert("ERROR: Date is Required.");
			return;
		}
		if (Trim(document.frmnohours.txtAppTFrom.value) == "") {
			alert("ERROR: Appointment Time (From:) is Required."); 
			return;
		}
		if (document.frmnohours.txtAppTFrom.value == "24:00") {
			alert("ERROR: Appointment Time (From:) is invalid (24:00 not accepted)."); 
			return;
		}
		if (Trim(document.frmnohours.txtAppTTo.value) == "") {
			alert("ERROR: Appointment Time (To:) is Required."); 
			return;
		}
		if (document.frmnohours.txtAppTTo.value == "24:00") {
			alert("ERROR: Appointment Time (To:) is invalid (24:00 not accepted)."); 
			return;
		}
		var ans = window.confirm("Submit Appointment to Database?");
		if (ans) {
			//alert(document.frmnohours.selIntr.value);
			document.frmnohours.action = "nohours_proc.asp"; // action.asp?ctrl=25";
			document.frmnohours.submit();
		}
	}
// 
--></script>
	</head>
	<body>
		<form method='post' name='frmnohours'>
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
<tr><td class='title' colspan='10' align='center'><nobr> No Appointment Hours</td></tr>
<tr><td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td></tr>
				<tr><td>&nbsp;</td></tr>
<tr><td align='right'>Type:</td><td><select name='selTrain' class='seltxt' style='width:150px;'>
							<option value='0'>Regular</option>
							<option value='1'>Training</option>
							<option value='2'>In house Training</option>
							<option value='3'>Interpreter Training Hours</option><!-- added 2017-12-04 on req from Alen -->
						</select>
					</td>
				</tr>
<tr><td>&nbsp;</td></tr>
<tr><td align='right'>*Institution:</td>
		<td class='confirm'>No Appointments Hours (other hours)</td>
		<td>&nbsp;</td></tr>
<tr><td align='right'>*Department:</td>
					<td><select class='seltxt' name='selDept'  style='width:250px;' onfocus='RPList(this.value); '  onchange='RPList(this.value);'>
							<option value='0'>&nbsp;</option>
							<%=strDept2%>
						</select>
					</td>
				</tr>
				<tr>
					<td align='right'>*Requesting Person:</td>
					<td>	
						<select class='seltxt' name='selRP'>
							<option value='0'>&nbsp;</option>
							</select>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>*Appointment Date:</td>
					<td>
						<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
						<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
						onclick="showCalendarControl(document.frmnohours.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
						<input type='hidden' name='mydate' value='<%=tmpAppDate%>'>
					</td>
				</tr>
				<tr>
					<td align='right'>*Appointment Time:</td>
					<td>
							From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpAppTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
							&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpAppTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
							<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
							<input type='hidden' name='mystime' value='<%=tmpAppTFrom%>'>
						
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right' valign='top'>*Interpeter:</td>
								<td>
									<select  name="selIntr" class='seltxt' multiple  style="height: 200px; width:250px;">
										<%=strIntr%>
									</select>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan="2" align="center">
									<input class='btn' type='button' value='Submit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveNoHours();'>
								</td>
							</tr>
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
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>