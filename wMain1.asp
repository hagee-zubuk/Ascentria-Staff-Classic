<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
tmpPage = "document.frmMain."
tmpInst = "-1"
tmpIntr = "-1"
'default
selRPEmail = ""
selRPPhone = ""
selRPFax = "checked"
'default
selIntrFax = "checked"
selIntrP2 = ""
selIntrP1 = ""
selIntrEmail = ""
tmpTS = Now
tmpDept = 0
tmpReqP = "-1"
tmpHPID = 0
If Request("tmpID") <> "" Then
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T WHERE index = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsw1.EOF Then
		myInst = rsW1("InstID")
		tmpEmer = ""
		If rsW1("Emergency") = True Then tmpEmer = "checked"
		tmpEmerFee = ""
		If rsW1("EmerFee") = True Then tmpEmerFee = "checked"
		tmpInstRate = rsW1("InstRate")
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE Active = 1 ORDER BY Facility"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpDO = ""
	If Cint(myInst) = rsInst("index") Then tmpDO = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option " & tmpDO & " value='" & rsInst("Index") & "'>" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET INSTITUTION RATES
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM rate_T ORDER BY Rate"
rsRate.Open sqlRate, g_strCONN, 3, 1
Do Until rsRate.EOF
	RateKo = ""
	If tmpInstRate = rsRate("Rate") Then Rateko = "selected"
	strRate1 = strRate1 & "<option " & Rateko & " value='" & rsRate("Rate") & "'>$" & Z_FormatNumber(rsRate("Rate"), 2) & "</option>" & vbCrLf
	rsRate.MoveNext
Loop
rsRate.Close
Set rsRate = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Contact Information</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
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
		function CalendarView(strDate)
		{
			document.frmMain.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmMain.submit();
		}
		function myfee()
		{
			if (document.frmMain.chkEmer.checked == true)
			{
				document.frmMain.chkEmerFee.disabled = false;
			}
			else
			{
				document.frmMain.chkEmerFee.disabled = true;
				document.frmMain.chkEmerFee.checked = false;
			}
		}
		function PopMe()
		{
			newwindow = window.open('find.asp','name','height=150,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			if (window.focus) {newwindow.focus()}
		}
		//function textboxchangeInst() 
		//{
		//	if (document.frmMain.btnNew.value == 'NEW')
		//	{
		//		alert("To save a new Institution, complete the form and click 'Next' button.");
		//		document.frmMain.btnNew.value = 'BACK';
		//		document.frmMain.selInst.disabled = true;
		//		document.frmMain.txtNewInst.style.visibility = 'visible';
		//		document.frmMain.HnewInt.value = 'BACK';
		//	}
		//	else
		//	{
		//		document.frmMain.btnNew.value = 'NEW';
		//		document.frmMain.selInst.disabled = false;
		//		document.frmMain.txtNewInst.value = "";
		//		document.frmMain.txtNewInst.style.visibility = 'hidden';
		//		document.frmMain.HnewInt.value = 'NEW';
		//	}
		//}
		//function hideNewInts() 
		//{
		//	if (document.frmMain.txtNewInst.value == "")
		//	{	
		//		document.frmMain.txtNewInst.style.visibility = 'hidden';
		//		document.frmMain.btnNew.value = 'NEW';
		//		document.frmMain.txtNewInst.value = "";
		//		document.frmMain.HnewInt.value = 'NEW';
		//	}
		//	else
		//	{
		//		document.frmMain.txtNewInst.style.visibility = 'visible';
		//		document.frmMain.btnNew.value = 'BACK';
		//		document.frmMain.selInst.disabled = true;
		//		document.frmMain.HnewInt.value = 'BACK';
		//	}
		//}
		function WNext()
		{
			if (document.frmMain.selInst.value == 0 && document.frmMain.HnewInt.value == 'NEW')
			{
				alert("ERROR: Institution is Required."); 
				return;
			}
			if (document.frmMain.HnewInt.value == 'BACK' && document.frmMain.txtNewInst.value == "")
			{
				alert("ERROR: Institution is Required."); 
				return;
			}
			document.frmMain.action = "waction.asp?ctrl=1";
			document.frmMain.submit();
		}
		//-->
		</script>
		</head>
		<body onload='myfee();'> <!--<body onload='hideNewInts(); myfee();'>//-->
			<form method='post' name='frmMain'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<form name='frmService' method='post' action=''>
								<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
									<tr>
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 1 / 4</td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td  align='left'>
											<div name="dErr" style="width:100%; height:55px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr>
										<td align='right'>Number of Appointments:</td>
										<td>
											<select class='seltxt' name='selNum'  style='width:50px;'>
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
											</select>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td class='header' colspan='10'><nobr>Contact Information</td>
									</tr>
									<tr>
										<td align='right'>Emergency:</td>
										<td width='300px'><input type='checkbox' name='chkEmer' value='1' <%=tmpEmer%> onclick='myfee();'></td>
									</tr>
									<tr>
										<td align='right'>Apply Emergency Fee:</td>
										<td><input type='checkbox' name='chkEmerFee' value='1' <%=tmpEmerFee%>></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>*Institution:</td>
										<td width='350px'>
											<select class='seltxt' name='selInst'  style='width:250px;'>
												<option value='0'>&nbsp;</option>
												<%=strInst%>
											</select>
											<input type='button'  value="FIND" <%=HpLock%>  name="findReq" onclick='PopMe();' title='Search instiution' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<!--<input class='btnLnk' type='button' name='btnNew' value='NEW'  <%=HpLock%> onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeInst();'>//-->
											<input type='hidden' name='hideInst' value='<%=tmpInst%>'>
										</td>
										<td>
											<input type='hidden' name='HnewInt'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<!--<td><input size='50' class='main' maxlength='50' name='txtNewInst' value='<%=tmpNewInstTxt%>' onkeyup='bawal(this);'></td>//-->
									</tr>
									<!--
										<tr>
										<td align='right' width='15%'>Rate:</td>
										<td>
											<select class='seltxt' style='width: 70px;' name='selInstRate'>
												<option value='0' >$0.00</option>
												<%=strRate1%>
											</select>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">
												*Rate varies per request
											</span>
										</td>
									</tr>
									-->
									<tr>
										<td align='right' width='15%'>&nbsp;</td>
										<td>
											&nbsp;
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" disabled>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.location='calendarview2.asp'">
											<input class='btn' type='button' value='>>' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WNext();'>
										</td>
									</tr>
									
								</table>
							</form>
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
