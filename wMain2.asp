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
	sqlW1 = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsw1.EOF Then
		myInst = rsW1("InstID")
		tmpEmer = ""
		If rsW1("Emergency") = True Then tmpEmer = "checked"
		tmpEmerFee = ""
		If rsW1("EmerFee") = True Then tmpEmerFee = "checked"
		tmpInstRate = rsW1("InstRate")
		tmpDept = rsW1("DeptID")
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET TEMP DATA
Set rsWdata = Server.CreateObject("ADODB.RecordSet")
sqlWdata = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
rsWdata.Open sqlWdata, g_strCONNW, 1, 3
If Not rsWdata.EOF Then
	tmpInst = rsWdata("instID")
	tmpEmer = ""
	If rsWdata("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)	
End If
rsWdata.Close
Set rsWdata = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
End If
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
'GET DEPT INFO
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE Active = 1 AND InstID = " & tmpInst & " ORDER BY dept"
rsDept.Open sqlDept, g_strCONN, 3, 1
Do Until rsDept.EOF
	tmpOLDAddr = rsDept("address") & "|" & rsDept("city") & "|" & rsDept("state") & "|" & rsDept("zip")
	strDept = strDept & "if (dept == " & rsDept("Index") & ") " & vbCrLf & _
		"{document.frmMain.txtInstAddr.value = """ & rsDept("address") &"""; " & vbCrLf & _
		"document.frmMain.selDept.value = " & rsDept("Index") & "; " & vbCrLf & _
		"document.frmMain.txtInstCity.value = """ & rsDept("city") &"""; " & vbCrLf & _
		"document.frmMain.txtInstState.value = """ & rsDept("state") &"""; " & vbCrLf & _
		"document.frmMain.txtInstZip.value = """ & rsDept("zip") &"""; " & vbCrLf & _
		"document.frmMain.txtInstAddrI.value = """ & rsDept("InstAdrI") &"""; " & vbCrLf & _
		"document.frmMain.txtBlname.value = """ & rsDept("BLname") &"""; " & vbCrLf & _
		"document.frmMain.txtBilAddr.value = """ & rsDept("Baddress") &"""; " & vbCrLf & _
		"document.frmMain.txtBilCity.value = """ & rsDept("Bcity") &"""; " & vbCrLf & _
		"document.frmMain.txtBilState.value = """ & rsDept("Bstate") &"""; " & vbCrLf & _
		"document.frmMain.txtBilZip.value = """ & rsDept("Bzip") &"""; " & vbCrLf & _
		"document.frmMain.selInstRate.value = """ & rsDept("defrate") &"""; " & vbCrLf & _
		"document.frmMain.OldAddr.value = """ & tmpOLDAddr &"""; " & vbCrLf & _
		"document.frmMain.selClass.value = """ & GetClass(rsDept("Class")) &"""; }" & vbCrLf 
		
		tmpDpt = ""
		If Cint(tmpDept) = rsDept("index") Then tmpDpt = "selected"
		DeptName = rsDept("Dept")
		strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept("Index") & "'>" &  DeptName & "</option>" & vbCrlf
	
	rsDept.MoveNext
Loop
rsDept.Close
Set rsDept = Nothing
'GET DEPARTMENTS
'Set rsDept2 = Server.CreateObject("ADODB.RecordSet")
'sqlDept2 = "SELECT * FROM dept_T WHERE InstID = " & tmpInst & " ORDER BY Dept"
'rsDept2.Open sqlDept2, g_strCONN, 3, 1
'Do Until rsDept2.EOF
'	tmpDpt = ""
'	If Cint(tmpDept) = rsDept2("index") Then tmpDpt = "selected"
'	DeptName = rsDept2("Dept")
'	strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept2("Index") & "'>" &  DeptName & "</option>" & vbCrlf
'	rsDept2.MoveNext
'Loop
'rsDept2.Close
'Set rsDept2 = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
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
		function DeptInfo(dept)
		{
			if (dept == 0 && document.frmMain.txtInstDept.value == "" )
			{
				document.frmMain.selDept.value =0;
				document.frmMain.txtInstDept.value = "";
				document.frmMain.txtInstAddr.value = "";
				document.frmMain.txtInstCity.value = "";
				document.frmMain.txtInstState.value = "";
				document.frmMain.txtInstZip.value = "";
				document.frmMain.txtInstAddrI.value = "";
				document.frmMain.txtBlname.value = "";
				document.frmMain.txtBilAddr.value = "";
				document.frmMain.txtBilCity.value = "";
				document.frmMain.txtBilState.value = "";
				document.frmMain.txtBilZip.value = "";
				document.frmMain.OldAddr.value = "";
			}
			else
			{
				hideNewDept();
			}
			<%=strDept%>
		}
		function hideNewDept() 
		{
			if (document.frmMain.txtInstDept.value == "")
			{	
				document.frmMain.txtInstDept.style.visibility = 'hidden';
				document.frmMain.txtInstDept.value = "";
				document.frmMain.HnewDept.value = 'NEW';
			}
			else
			{
				document.frmMain.txtInstDept.style.visibility = 'visible';
				document.frmMain.selDept.disabled = true;
				document.frmMain.selClass.value = '<%=tmpClass%>';
				document.frmMain.txtInstAddr.value = '<%=tmpNewInstAddr%>';
				document.frmMain.txtInstCity.value = '<%=tmpNewInstCity%>';
				document.frmMain.txtInstState.value = '<%=tmpNewInstState%>';
				document.frmMain.txtInstZip.value = '<%=tmpNewInstZip%>';
				document.frmMain.txtInstAddrI.value = '<%=tmpNewInstAddrI%>';
				document.frmMain.txtBlname.value = '<%=tmpBLname%>';
				if (document.frmMain.chkBill.checked != true)
				{
					document.frmMain.chkBill.checked = false;
					document.frmMain.txtBilAddr.value = '<%=tmpBilInstAddr%>';
					document.frmMain.txtBilCity.value = '<%=tmpBilInstCity%>';
					document.frmMain.txtBilState.value = '<%=tmpBilInstState%>';
					document.frmMain.txtBilZip.value = '<%=tmpBilInstZip%>';
				}
				else
				{
					document.frmMain.chkBill.checked = true;
					document.frmMain.txtBilAddr.value = "";
					document.frmMain.txtBilCity.value = "";
					document.frmMain.txtBilState.value = "";
					document.frmMain.txtBilZip.value = "";
				}
			}
		}
		
		function WNext()
		{
			var strNewAddr = document.frmMain.txtInstAddr.value + "|" + document.frmMain.txtInstCity.value + "|" + document.frmMain.txtInstState.value + "|" + document.frmMain.txtInstZip.value;
			if (strNewAddr != document.frmMain.OldAddr.value)
			{
				var ans = window.confirm("WARNING: Changing of institution address will be effective for all instances of that institution. Click Cancel to stop.");
				if (!ans)
				{
					return;
				}
			}
			if (document.frmMain.txtInstAddr.value == "" || document.frmMain.txtInstCity.value == "" || document.frmMain.txtInstState.value == "" || document.frmMain.txtInstZip.value == "")
			{
				alert("ERROR: Department's full address is required."); 
				return;
			}
			document.frmMain.action = "waction.asp?ctrl=2";
			document.frmMain.submit();
		}
		function WBack(xxx)
		{
			var ans = window.confirm("Any changes made in this page will not be saved.");
			if (ans){
				document.frmMain.action = "wMain1.asp?tmpID=" + xxx;
				document.frmMain.submit();
			}
		}
		//-->
		</script>
		</head>
		<body onload='DeptInfo(<%=tmpDept%>); hideNewDept();'>
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
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 2 / 4</td>
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
										<td class='header' colspan='10'><nobr>Contact Information</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<% If tmpInst = 479 Then %>
	<tr>
		<td align='right'>Training:</td>
		<td><input type='checkbox' name='chkTrain' value='1'></td>
	</tr>
	<tr><td>&nbsp;</td></tr>
									<% End If %>
									<tr>
										<td align='right'>Institution:</td>
										<td class='confirm'><%=tmpIname%></td>
									</tr>
									<!--<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>
									<% End If %>//-->
									<tr>
										<td align='right' width='15%'>*Department:</td>
										<td>	
											<select class='seltxt' name='selDept'  style='width:250px;' onfocus='DeptInfo(document.frmMain.selDept.value); '  onchange='DeptInfo(document.frmMain.selDept.value);'>
												<option value='0'>&nbsp;</option>
												<%=strDept2%>
											</select>
											<!--<input class='btnLnk' type='button' name='btnNewDept' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeDept();'>//-->
											<input type='hidden' name='HnewDept'>
											<input type='hidden' name='hideDept' value='<%=tmpDept%>'>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><input class='main' size='50' maxlength='50' name='txtInstDept' value='<%=tmpNewInstDept%>' onkeyup='bawal(this);'></td>
									</tr>
									<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td>
											<select class='seltxt' style='width: 70px;' name='selInstRate'>
												<option value='0' >$0.00</option>
												<%=strRate1%>
											</select>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">
												*Rate may vary per request
											</span>
										</td>
									</tr>
									<% Else %>
										<input class='main' size='5' maxlength='5' type='hidden' name='selInstRate' value='<%=tmpInstRate%>'>
									<% End If %>
									<tr>
										<td align='right'>Classification:</td>
										<td>
											
											<input class='main' size='50' maxlength='50' name='selClass' value='<%=tmpClass%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddrI' value='<%=tmpNewInstAddrI%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddr' value='<%=tmpNewInstAddr%>' onkeyup='bawal(this);' Readonly>
											
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtInstCity' value='<%=tmpNewInstCity%>' onkeyup='bawal(this);' Readonly>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtInstState' value='<%=tmpNewInstState%>' onkeyup='bawal(this);' Readonly>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtInstZip' value='<%=tmpNewInstZip%>' onkeyup='bawal(this);' Readonly>
											<input type='hidden' name='OldAddr'>
										</td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBlname' value='<%=tmpBLname%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										<!--<td>
											<input type='checkbox' name='chkBill' <%=chkBillMe%>>
											(same as appointment address)
										</td>//-->
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBilAddr' value='<%=tmpBilInstAddr%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtBilCity' value='<%=tmpBilInstCity%>' onkeyup='bawal(this);' Readonly>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtBilState' value='<%=tmpBilInstState%>' onkeyup='bawal(this);' Readonly>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtBilZip' value='<%=tmpBilInstZip%>' onkeyup='bawal(this);' Readonly>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WBack(<%=Request("tmpID")%>);'>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.location='calendarview2.asp'">
											<input class='btn' type='button' value='>>' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WNext();'>
											<input type='hidden' name='tmpID' value='<%=Request("tmpID")%>'>
											<input type='hidden' name='tmpInst' value='<%=tmpInst%>'>
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
