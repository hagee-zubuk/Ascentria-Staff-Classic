<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%

Function GetLoc(xxx)
	Select Case xxx
		Case 0 
			GetLoc = "Front Door"
		Case 1
			GetLoc = "Cafeteria"
		Case 2
			GetLoc = "Registration Desk"
		Case 3
			GetLoc = "Department"
		Case 4
			GetLoc = "OTHER"
	End Select
End Function
'GET REQUEST
Set rsConfirm = Server.CreateObject("ADODB.RecordSet")
sqlConfirm = "SELECT * FROM Request_T WHERE [index] = " & Request("ID")
rsConfirm.Open sqlConfirm, g_strCONN, 3, 1
If Not rsConfirm.EOF Then
	TS = rsConfirm("timestamp")
	RP = rsConfirm("ReqID") 
	PID = rsConfirm("index")
	'Fon = rsConfirm("phone") 
	'Fax = rsConfirm("fax")
	'email = rsConfirm("email")
	tmpName = rsConfirm("clname") & ", " & rsConfirm("cfname")
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & "," & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	tmpCfon = rsConfirm("Cphone")
	tmpCAfon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = rsConfirm("spec_cir")
	tmpDOB = rsConfirm("DOB")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = CTime(rsConfirm("appTimeFrom")) 
	tmpAppTTo = rsConfirm("appTimeTo")
	tmpAppLoc = rsConfirm("appLoc")
	tmpInst = rsConfirm("instID")
	tmpDept = rsConfirm("DeptID")
	tmpRate = Z_FormatNumber(rsConfirm("InstRate"), 2)
	tmpDoc = rsConfirm("docNum")
	tmpCRN = rsConfirm("CrtRumNum")
	tmpIntr = rsConfirm("IntrID")
	tmpBH = Z_FormatNumber(rsConfirm("Billable"), 2)
	tmpActDate = rsConfirm("aDate")
	tmpClient = ""
	If rsConfirm("Client") = True Then tmpClient = " (LSS Client)"
	tmpActTFrom = rsConfirm("astarttime") 
	tmpActTTo = rsConfirm("aendtime")
	tmpComm = rsConfirm("comment")
	tmpEmer = "___"
	if rsConfirm("emergency") = True Then tmpEmer = "<u><b>_X_</b></u>"
	If rsConfirm("CliAdd") = True Then tmpDeptaddr = rsConfirm("CAddress") &", " & rsConfirm("CliAdrI") & ", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
	tmpHPID = Z_CZero(rsConfirm("hpid"))

	If IsNull( rsConfirm("Gender") ) Then
		tmpSex = "Unknown"
	Else
		If rsConfirm("Gender") = 1 Then
			tmpSex = "FEMALE"
		ElseIf rsConfirm("Gender") = 0 Then
			tmpSex = "MALE"
		End If
	End If

	tmpMinor2 = "ADULT"
	If rsConfirm("Child") Then tmpMinor2 = "MINOR"	
End If
rsConfirm.Close
Set rsConfirm = Nothing
'GET REQUESTING PERSON'S INFO
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & RP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpReqName = rsReq("lname") & ", " & rsReq("fname")
	tmpReqEmail = rsReq("Email")
	tmpReqFon = rsReq("phone")
	tmpReqFax = rsReq("fax")
End If
rsReq.Close
Set rsReq = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	'tmpClass = rsInst("Class")
	'tmpClassName = GetClass(rsInst("Class"))
	tmpIname = rsInst("Facility") 
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	tmpClass = rsDept("Class")
	tmpClassName = GetClass(rsDept("Class"))
	If rsDept("dept") <> "" Then  tmpIname = tmpIname & " - " & rsDept("dept")
	If tmpDeptaddr = "" Then tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
End If
rsDept.Close
Set rsDept = Nothing 
'GET LANGUAGE
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang  = "SELECT * FROM language_T WHERE [index] = " & tmpLang
rsLang.Open sqlLang , g_strCONN, 3, 1
If Not rsLang.EOF Then
	tmpSalita = rsLang("language") 
End If
rsLang.Close
Set rsLang = Nothing 
'GET INTERPRETER INFO
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & tmpIntr
rsIntr.Open sqlIntr, g_strCONN, 3, 1
If Not rsIntr.EOF Then
	tmpInName = rsIntr("last name") & ", " & rsIntr("first name")
	tmpInEmail = rsIntr("E-mail")
	tmpInFon = rsIntr("phone1")
	If rsIntr("phone2") <> "" Then tmpInFon = tmpInFon & " / " & rsIntr("phone2")
	tmpInFax = rsIntr("fax")
	tmpInaddr = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("zip code")	
	tmpInHouse = ""
	If rsIntr("InHouse") = True Then tmpInHouse = " (In-House Interpreter)"
End If
rsIntr.Close
Set rsIntr = Nothing
'GET HP INFO
If tmpHPID <> 0 Then
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
	sqlHP = "SELECT * FROM appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_strCONNHP, 3, 1
	If Not rsHp.EOF Then
		'tmpReason = rsHP("reason")
		tmpReason = GetReas(Z_Replace(rsHP("reason"),", ", "|"))
		tmpClin = rsHP("clinician")
		tmpMeet = ""
		If rsHP("mwhere") = 1 Then
			tmpMeet = UCase(GetLoc(rsHP("mlocation")))
			If tmpMeet = "OTHER" Then tmpMeet = rsHP("mother")
		End If	
		tmpMinor = ""
		If rsHP("minor") = True Then tmpMinor = "* Minor"
		tmpParents = ""
		If rsHP("parents") <> "" Then tmpParents = rsHP("parents") 
	End If
	rsHp.Close
	Set rsHP = Nothing
End If
'SAVE PRINT TIME 
If Request("PDF") <> 1 Then
	Set rsPrint = Server.CreateObject("ADODB.RecordSet")
	sqlPrint = "UPDATE request_T SET [Print] = '" & Now & "' WHERE [index] = " & Request("ID")
	rsPrint.Open sqlPrint, g_strCONN, 1, 3
	Set rsPrint = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Request Confirmation - Print</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		 
	  -->
	</script>
	</head>
		<body >
			<table cellSpacing='0' cellPadding='0' width='100%' bgColor='white' border='0' align='center'>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td valign='top' align='center'>
						<table cellSpacing='0' cellPadding='0' height='120px' width='620px' bgcolor='#D3D3D3' border='5'>
							<tr>
								<td align='center'>
									<img src='images/LBISLOGOtrans.gif' align='center' height='60px'>
									<br>
									261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>
									Tel:&nbsp;(603)&nbsp;224-8111&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
									<br><br>
									<p class='title'>Service Verification Form</p>
								</td>
							</tr>
							
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<% If tmpClass = 3 Then %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='2' align='center' class='header' style='border:1px solid #000000;' bgcolor='#D3D3D3'>
										Court Assignment
										<p class='notes'>To be completed by Language Bank Staff</p>
									</td>
								</tr>
								<tr>
									<td colspan='2' class='printForm'>
										<br>
										Project ID:&nbsp;&nbsp;<b><%=PID%></b>
										<!--&nbsp;&nbsp;<img src="_barcode.asp?code=<%=PID%>&height=20&width=1&mode=code39&text=1">//-->
										<br>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Date:<br>
										&nbsp;&nbsp;<b><%=Date%></b>
									</td>
									<td class='printForm'>
										Interpreter:<br>
										&nbsp;&nbsp;<b><%=tmpInName%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Service Recipient:<br>
										&nbsp;&nbsp;<b><%=tmpName%></b>
									</td>
									<td class='printForm'>
										Language:<br>
										&nbsp;&nbsp;<b><%=tmpSalita%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Requesting Court:<br>
										&nbsp;&nbsp;<b><%=tmpIname%></b>
									</td>
									<td class='printForm'>
										Date of Service:<br>
										&nbsp;&nbsp;<b><%=tmpAppDate%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Address for Service:<br>
										&nbsp;&nbsp;<b><%=tmpDeptaddr%></b>
									</td>
									<td class='printForm'>
										Time of Service:<br>
										&nbsp;&nbsp;<b><%=tmpAppTFrom%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Docket Number:<br>
										&nbsp;&nbsp;<b><%=tmpDoc%></b>
									</td>
									<td class='printForm'>
										Court Room No:<br>
										&nbsp;&nbsp;<b><%=tmpCRN%></b>
									</td>
								</tr>
								<tr>
									<td colspan='2' class='printForm'>
										Comment:<br>
										&nbsp;&nbsp;<b><%=tmpComm%></b>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='3' align='center' class='header' style='border:1px solid #000000;' bgcolor='#D3D3D3'>
										Job Verification
										<p class='notes'>
											To be completed at time of service<br>
											To be completed by Language Bank Interpreter and signed by requesting agency and interpreter
										</p>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' style='border-right-color:#FFFFFF;'>
										Date of Service:
									</td>
									<td class='printForm' colspan='2' width='250px' style='border-left-color:#FFFFFF;'>
										Total Hours:
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' style='border-right-color:#FFFFFF;'>
										Arrival Time:
									</td>
									<td class='printForm' colspan='2' style='border-left-color:#FFFFFF;'>
										Departure Time:
									</td>
								</tr>
								<tr height='50px'>
									<td class='printForm' valign='top' style='border-right-color:#FFFFFF;'>
										Signature of Interpreter:
									</td>
									<td class='printForm' colspan='2' valign='top' style='border-left-color:#FFFFFF;'>
										Date:
									</td>
								</tr>
								<tr height='50px'>
									<td class='printForm' valign='top' style='border-right-color:#FFFFFF;'>
										Signature of Court Representative Authorizing Service:
									</td>
									<td class='printForm' valign='top' style='border-left-color:#FFFFFF; border-right-color:#FFFFFF;'>
										Title:
									</td>
									<td class='printForm' valign='top' style='border-left-color:#FFFFFF;'>
										Date:
									</td>
								</tr>
								<tr bgcolor='#D3D3D3'>
									<td colspan='3' align='center' style='border:1px solid #000000;'>
											<table cellSpacing='0' cellPadding='0' width='620px' bgColor='#D3D3D3' border='0'>
												<tr>
													<td colspan='3' class='printForm'>
														For Billing Department - Billing Address<br>
														&nbsp;&nbsp;&nbsp;&nbsp;
														<u>Dale Trombley, Administrative Office of the Courts, 2 Noble Drive, Concord, NH 03301</u>
													</td>
												</tr>
												<tr height='25px' bgcolor='#D3D3D3'>
													<td class='printForm' valign='top' style='border-right-color:#D3D3D3;'>
														___ Block Scheduled Time: ____________<br>
														___ Regular Appoinment: _____________<br>
														<%=tmpEmer%> Emergency Appointment: _________<br>
														___ Cost of Service Hours: ____________<br>
													</td>
													<td class='printForm' colspan='2' valign='top' style='border-left-color:#D3D3D3;'>
															&nbsp;&nbsp;Travel Time: ______________<br>
															&nbsp;&nbsp;Total Allowable Mileage: ______________<br>
															&nbsp;&nbsp;Total Travel Cost: ______________<br>
															&nbsp;&nbsp;Cancellation Fee: _______________<br>
															<b>&nbsp;&nbsp;TOTAL COST: _______________</b>
														</form>
													</td>
												</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% Else %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='2' align='center' class='header' style='border:1px solid #000000;' bgcolor='#D3D3D3'>
										Assignment
										<p class='notes'>To be completed by Language Bank Staff</p>
									</td>
								</tr>
								<tr>
									<td colspan='2' class='printForm'>
										<br>
										Project ID:&nbsp;&nbsp;<b><%=PID%></b>
										<!--&nbsp;&nbsp;<img src="_barcode.asp?code=<%=PID%>&height=20&width=1&mode=code39&text=1">//-->
										<br>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Date:
										&nbsp;&nbsp;<b><%=Date%></b>
									</td>
									<td class='printForm'>
										Interpreter:
										&nbsp;&nbsp;<b><%=tmpInName%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Service Recipient(Client):
										&nbsp;&nbsp;<b><%=tmpName%></b><br>
										DOB:
										&nbsp;&nbsp;<b><%=tmpDOB%></b>
									</td>
									<td class='printForm' valign='top' style='height: 50px;'>
										Language:
										&nbsp;&nbsp;<b><%=tmpSalita%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Type of Appointment:
										&nbsp;&nbsp;<b><%=tmpClassName%></b>
									</td>
									<td class='printForm'>
										Date of Appointment:
										&nbsp;&nbsp;<b><%=tmpAppDate%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Agency Requesting Service:<br>
										&nbsp;&nbsp;<b><%=tmpIname%></b>
									</td>
									<td class='printForm'>
										Time of Service:
										&nbsp;&nbsp;<b><%=tmpAppTFrom%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' valign='top'>
										Address for Interpretation:<br>
										&nbsp;&nbsp;<b><%=tmpDeptaddr%></b>
									</td>
									<% If tmpHPID = 0 Then %>
										<td class='printForm'>
											Estimated Time Requested:<br>
											&nbsp;&nbsp;<b>&nbsp;</b>
										</td>
									<%Else%>
										<td class='printForm'>
											Clinician:<br>
											&nbsp;&nbsp;<b><%=tmpClin%></b><br>
											<% If tmpParents <> "" Then%>
												Parents:<br>
												&nbsp;&nbsp;<b><%=tmpParents%></b>
											<% End If %>
										</td>
									<%End If%>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Adult or Child:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Gender:<br>
										&nbsp;&nbsp;<b><%=tmpMinor2%></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<b><%=tmpSex%></b>
									</td>
									<% If tmpHPID = 0 Then %>
										<td class='printForm'>
											Brief Description of Job:<br>
											&nbsp;&nbsp;<b><%=tmpSC%></b>
										</td>
									<%Else%>
										<td class='printForm' valign='top'>
											Reason:<br>
											&nbsp;&nbsp;<b><%=tmpReason%></b>
										</td>
									<%End If%>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Client Phone:
										&nbsp;&nbsp;<b><%=tmpCfon%></b>
										<br><font size='1'>*please call client to remind  them of appointment 48 hours prior to appointment.</font>
										<br>
										Client Alter. Phone:
										&nbsp;&nbsp;<b><%=tmpCAfon%></b>
									</td>
									<td class='printForm'>
										Specific Request/Comments:<br>
										&nbsp;&nbsp;<b><%=tmpComm%></b>
									</td>
								</tr>
							
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='3' align='center' class='header' style='border:1px solid #000000;' bgcolor='#D3D3D3'>
										Job Verification
										<p class='notes'>
											To be completed at time of service
										</p>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' width='300px' style='border-right-color:#FFFFFF;'>
										Date of Service:
									</td>
									<td class='printForm' colspan='2' style='border-left-color:#FFFFFF;'>
										Total Hours:
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' style='border-right-color:#FFFFFF;'>
										Arrival Time:
									</td>
									<td class='printForm' colspan='2' style='border-left-color:#FFFFFF;'>
										Departure Time:
									</td>
								</tr>
								<tr height='50px'>
									<td class='printForm' valign='top' style='border-right-color:#FFFFFF;'>
										Signature of Interpreter:
									</td>
									<td class='printForm' colspan='2' valign='top' style='border-left-color:#FFFFFF;'>
										Date:
									</td>
								</tr>
								<tr height='50px'>
									<td class='printForm' valign='top' style='border-right-color:#FFFFFF;'>
										Signature of Agency Representative Requesting Service:
									</td>
									<td class='printForm' valign='top' style='border-left-color:#FFFFFF; border-right-color:#FFFFFF;'>
										Title:
									</td>
									<td class='printForm' valign='top' style='border-left-color:#FFFFFF;'>
										Date:
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' colspan='3'>
										Billing:
										&nbsp;<b><%=tmpBContact%></b><br>
										&nbsp;&nbsp;<b><%=tmpBaddr%></b>
									</td>
								</tr>
								<tr>
									<td colspan='3'>	
										<table cellspacing='0' cellpadding='0' width='100%' border='0'>
											<tr>
												<td class='header' align='center' colspan='3' bgcolor='#D3D3D3' style='border:1px solid #000000;'>
													FOR  INTERPRETERS
													<p class='notes'>
														Please fill out section below for follow up appointments:
													</p>
												</td>
											</tr>
											<tr>
												<td class='printForm' style='text-align: center;'>Date</td>
												<td class='printForm' style='text-align: center;'>Time</td>
												<td class='printForm' style='text-align: center;' width='150px'>Are you available</td>
											</tr>
											<tr>
												<td class='printForm'>1)</td>
												<td class='printForm'>&nbsp;</td>
												<td class='printForm' style='text-align: center;'>Yes* / No
											</tr>
											<tr>
												<td class='printForm'>2)</td>
												<td class='printForm'>&nbsp;</td>
												<td class='printForm' style='text-align: center;'>Yes* / No
											</tr>
											<tr>
												<td class='printForm'>3)</td>
												<td class='printForm'>&nbsp;</td>
												<td class='printForm' style='text-align: center;'>Yes* / No
											</tr>
											<tr>
												<td class='printForm' colspan='3' style='text-align: center;'>*Please note that being available does not mean appointment will be assigned to you </td>
												
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% End If %>
				<% If Request("PDF") <> 1 Then %>
					<tr>
						<td align='center' valign='bottom'>
							<br>
							<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
						</td>
					</tr>
				<% End If %>
			</table>	
		</body>
	</html>
