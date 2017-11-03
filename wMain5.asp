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
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE index = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			'GetPrime = rsRP("Phone")
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
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
tmpReqP = 0
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
		tmpDept = rsW1("DeptID")
		tmpReqP = rsW1("ReqID")
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET TEMP DATA
Set rsWdata = Server.CreateObject("ADODB.RecordSet")
sqlWdata = "SELECT * FROM Wrequest_T WHERE index = " & Request("tmpID")
rsWdata.Open sqlWdata, g_strCONNW, 1, 3
If Not rsWdata.EOF Then
	tmpInst = rsWdata("instID")
	tmpEmer = ""
	If rsWdata("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)
	tmpClient = ""
	If rsWdata("client") = True Then tmpClient = " (LSS Client)"
	tmpName = rsWdata("clname") & ", " & rsWdata("cfname") & tmpClient
	tmpAddr = rsWdata("CliAdrI") & " " & rsWdata("caddress") & ", " & rsWdata("cCity") & ", " &  rsWdata("cstate") & ", " & rsWdata("czip")
	tmpFon = rsWdata("Cphone")
	tmpAFon = rsWdata("CAphone")
	tmpDir = rsWdata("directions")
	tmpSC = rsWdata("spec_cir")
	tmpDOB = rsWdata("DOB")
	tmpLang = rsWdata("langID")
	tmpAppDate = rsWdata("appDate")
	tmpAppTFrom = rsWdata("appTimeFrom") 
	tmpAppTTo = rsWdata("appTimeTo")
	tmpAppLoc = rsWdata("appLoc")
	tmpInst = rsWdata("instID")
	tmpDept = rsWdata("DeptID")
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)
	tmpDoc = rsWdata("docNum")
	tmpCRN = rsWdata("CrtRumNum")	
End If
rsWdata.Close
Set rsWdata = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE index = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE index = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		'tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
End If
rsDept.Close
Set rsDept = Nothing 
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE index = " & tmpReqP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(tmpReqP)
End If
rsReq.Close
Set rsReq = Nothing
'GET INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = True ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	IntrSel = ""
	If CInt(tmpIntr) = rsIntr("index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
	strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsIntr("index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
			"document.frmMain.selIntr.appendChild(ChoiceReq);}" & vbCrLf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Appointment Information</title>
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
		function WSubmit(xxx)
		{
			var ans = window.confirm("Sumbit Appointment to Database?");
			if (ans){
				document.frmMain.action = "waction.asp?ctrl=4&Submit=1";
				document.frmMain.submit();
			}
		}
		function WNext()
		{
			if (document.frmMain.txtClilname.value == "" && document.frmMain.txtClifname.value == "")
			{
				alert("ERROR: Client is Required."); 
				return;
			}
			if (document.frmMain.selLang.value == 0)
			{
				alert("ERROR: Language is Required."); 
				return;
			}
			if (document.frmMain.txtAppDate.value == "")
			{
				alert("ERROR: Appointment Date is Required."); 
				return;
			}
			if (document.frmMain.txtAppTFrom.value == "")
			{
				alert("ERROR: Appointment Time (From:) is Required."); 
				return;
			}
			document.frmMain.action = "waction.asp?ctrl=4";
			document.frmMain.submit();
		}
		function WBack(xxx)
		{
			var ans = window.confirm("Any changes made in this page will not be saved.");
			if (ans){
				document.frmMain.action = "wMain3.asp?tmpID=" + xxx;
				document.frmMain.submit();
			}
		}
		//-->
		</script>
		</head>
		<body onload=''>
			<form method='post' name='frmMain' action='main.asp'>
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
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 4 / 6</td>
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
									<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%></td>
									</tr>
									<tr>
										<td align='right'>Department:</td>
										<td class='confirm'><%=tmpDname%></td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td class='confirm'><%=tmpDeptaddr%></td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td class='confirm'><%=tmpBContact%></td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										<td class='confirm'><%=tmpBaddr%></td>
									</tr>
									<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
										<tr>
											<td align='right' width='15%'>Rate:</td>
											<td class='confirm'><%=tmpInstRate%></td>
										</tr>
									<% End If %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>Requesting Person:</td>
										<td class='confirm'><%=tmpRP%></td>
									</tr>
									<tr>
										<td align='right'>Phone:</td>
										<td class='confirm'><%=fon%></td>
									</tr>
									<tr>
										<td align='right'>Fax:</td>
										<td class='confirm'><%=fax%></td>
									</tr>
									<tr>
										<td align='right'>E-Mail:</td>
										<td class='confirm'><%=email%></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Appointment Information</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>Client Name:</td>
										<td class='confirm'><%=tmpName%></td>
									</tr>
									<tr>
										<td align='right'>Client Address:</td>
										<td class='confirm'><%=tmpAddr%></td>
									</tr>
									<tr>
										<td align='right'>Language:</td>
										<td class='confirm'><%=tmpSalita%></td>
									</tr>
									<tr>
										<td align='right'>Appointment Date:</td>
										<td class='confirm'><%=tmpAppDate%></td>
									</tr>
									<tr>
										<td align='right'>Appointment Time:</td>
										<td class='confirm'><%=tmpAppTFrom%> - <%=tmpAppTTo%></td>
									</tr>
									<tr>
										<td align='right'>Docket Number:</td>
										<td class='confirm'><%=tmpDoc%></td>
									</tr>
									<tr>
										<td align='right'>Court Room No:</td>
										<td class='confirm'><%=tmpCRN%></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>Appointment Comment:</td>
										<td class='confirm'><%=tmpCom%></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Interpreter Information</td>
									</tr>
									<tr>
										<td align='right'>Interpreter:</td>
										<td>
											<select class='seltxt' name='selIntr' style='width: 200px;' onchange='JavaScript:IntrInfo(document.frmMain.selIntr.value);'>
												<option value='-1'>&nbsp;</option>
												<%=strIntr%>
											</select>
											<input type="button" value="..." name="btnchkavail"
											onclick='PopMe2(document.frmMain.txtAppDate.value, document.frmMain.txtAppTFrom.value, document.frmMain.selLang.value);' title='Check Availability' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input class='btnLnk' type='button' name='btnNewIntr' value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeIntr();'>
											<input type='checkbox' name='chkAll2' onclick='IntrShowMe(); IntrInfo(document.frmMain.selIntr.value);'>
											Show All
											<input type='hidden' name='HnewIntr'>
											<input type='hidden' name='Lang1'>
											<input type='hidden' name='Lang2'>
											<input type='hidden' name='Lang3'>
											<input type='hidden' name='Lang4'>
											<input type='hidden' name='Lang5'>
											<input type='hidden' name='LangCtr'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtIntrLname' value='<%=tmpIntrLname%>' onkeyup='bawal(this);'>
											<input class='trans' size='1' style='width: 5px;' name='txtcoma' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtIntrFname' value='<%=tmpIntrFname%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim2'  readonly size='12'>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='0'  <%=selIntrEmail%> onclick='chkPrim2();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtIntrEmail' value='<%=tmpIntrEmail%>' onkeyup='bawal(this);'></td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='1' <%=selIntrP1%> onclick='chkPrim2();'>&nbsp;
											Home Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrP1' value='<%=tmpIntrP1%>' onkeyup='bawal(this);'>
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtIntrExt' value='<%=tmpIntrExt%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='3' <%=selIntrFax%> onclick='chkPrim2();'>
											&nbsp;&nbsp;&nbsp;Fax:
											</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrFax' value='<%=tmpIntrFax%>' onkeyup='bawal(this);'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Please include area code on fax number</span>
										</td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='2' <%=selIntrP2%> onclick='chkPrim2();'>
											Mobile Phone:
											</td>
										<td><input class='main' size='12' maxlength='12' name='txtIntrP2' value='<%=tmpIntrP2%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddrI' value='<%=tmpNewIntrAddrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddr' value='<%=tmpIntrAddr%>' onkeyup='bawal(this);'>
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtIntrCity' value='<%=tmpIntrCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtIntrState' value='<%=tmpIntrState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtIntrZip' value='<%=tmpIntrZip%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>In-House Interpreter:</td>
										<td><input type='checkbox' name='chkInHouse' value='1' <%=tmpInHouse%>></td>
									</tr>
									<tr>
										<td align='right' width='15%'>Default Rate:</td>
										<td>
											<input class='main' size='5' maxlength='5'  readonly  name='txtIntrRate' value='<%=tmpIntrRate%>'>
											<select class='seltxt' style='width: 70px;' name='selIntrRate'>
												<option value='0' >&nbsp;</option>
												<%=strRate2%>
											</select>
										</td>
									<tr>
								
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Interpreter Comment:</td>
										<td colspan='3' >
											<textarea name='txtcomintr' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpComintr%></textarea>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WBack(<%=Request("tmpID")%>);'>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Submit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WSubmit(<%=Request("tmpID")%>);'>
											<input class='btn' type='button' value='>>' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WNext();'>
											<input type='hidden' name='tmpID' value='<%=Request("tmpID")%>'>
											<input type='hidden' name='tmpInst' value='<%=tmpInst%>'>
											<input type='hidden' name='tmpDep' value='<%=tmpDept%>'>
											<input type='hidden' name='tmpReqP' value='<%=tmpReqP%>'>
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
