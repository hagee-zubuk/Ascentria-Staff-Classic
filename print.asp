<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
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
	tmpfName = UCase(rsConfirm("cfname")) 'rsConfirm("clname") & ", " & rsConfirm("cfname")
	tmplName = UCase(rsConfirm("clname"))
	tmpAddr = rsConfirm("caddress") & ", " & rsConfirm("CliAdrI") & "," & rsConfirm("cCity") & ", " &  rsConfirm("cstate") & ", " & rsConfirm("czip")
	tmpCfon = rsConfirm("Cphone")
	tmpCAfon = rsConfirm("CAphone")
	tmpDir = rsConfirm("directions")
	tmpSC = Replace(Z_FixNull(rsConfirm("spec_cir")), vbcrlf, "<br>")
	'tmpDOB = rsConfirm("DOB")
	mrrec = rsConfirm("mrrec")
	tmpLang = rsConfirm("langID")
	tmpAppDate = rsConfirm("appDate")
	tmpAppTFrom = CTime(rsConfirm("appTimeFrom")) 
	tmpAppTTo = CTime(rsConfirm("appTimeTo"))
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
	tmpEmer = "___"
	tmpEmer2 = ""
	if rsConfirm("emergency") = True Then 
		tmpEmer = "<u><b>_X_</b></u>"
		tmpEmer2 = "(EMERGENCY)"
	end if
	tmpComm = tmpEmer2 & "<br>" & Replace(Z_FixNull(rsConfirm("comment")), vbcrlf, "<br>")
	If rsConfirm("CliAdd") = True Then tmpDeptaddr = rsConfirm("CAddress") &", " & rsConfirm("CliAdrI") & ", " & rsConfirm("CCity") & ", " & rsConfirm("CState") & ", " & rsConfirm("CZip")
	tmpHPID = Z_CZero(rsConfirm("hpid"))
	tmpGender	= Z_CZero(rsConfirm("Gender"))
	If tmpGender = 0 Then 
		tmpSex = "MALE"
	Else
		tmpSex = "FEMALE"
	End If
	tmpMinor2 = "ADULT"
	If rsConfirm("Child") Then tmpMinor2 = "MINOR"	
	courtcall = ""
	If rsConfirm("courtcall") Then courtcall = "<font size='1'><i>*Call patrient/client to remind them of appointment 48 hours prior to appointment.</i></font>"
	leavemsg = "<font size='1'><i>*Do NOT leave DETAILED reminder call message on patient/client’s voice mail or with any family members.</i></font>"
	If rsConfirm("leavemsg") Then leavemsg = "<font size='1'><i>*It is OK to leave/provide full appointment info on voice mail or with the family member.</i></font>"
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
	strNotes = Trim(rsInst("Instnotes"))
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
	tmpdeptname = rsDept("dept")
	If tmpDeptaddr = "" Then tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	If Trim(rsDept("deptNotes")) <> "" Then strNotes = Trim(rsDept("deptNotes"))
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
		tmpMinor2 = ""
		If rsHP("minor") = True Then tmpMinor2 = "* Minor"
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
myPath = "C:\work\LSS-LBIS\web\Images\BC\" & PID & ".bmp"
myBC = PID & ".bmp"
'name
tmpName = tmpfName
If tmpClass = 3 Or tmpClass = 5 Then tmpName = tmpfName & " " & Left(tmplName, 1) 
'671	
%>
<html>
	<head>
		<title>Language Bank - Request Confirmation - Print</title>
		<meta content="text/html; charset=utf-8" http-equiv="content-type">
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		 
	  -->
	</script>
	</head>
		<body >
			<table cellSpacing='0' cellPadding='0' width='100%' bgColor='white' border='0' align='center'>
				<tr><td>&nbsp;</td></tr>
				<% If tmpInst = 671 Then %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' height='80px' width='620px' bgcolor='#FFFFFF'  border='0' align='center'>
								<tr>
									<td align='right' colspan='2'>
										<table cellSpacing='0' cellPadding='0'  bgcolor='#FFFFFF'  border='0' align='right'>
											<tr>
												<td align='center' width='100%' style='border: solid 1px;'>
													<img src="_barcode.asp?code=<%=PID%>&height=20&width=2&mode=code39&text=0&fileout=<%=myPath%>" style="visibility:hidden" >
													<br />
										    	&nbsp;<img src="Images/BC/<%=myBC%>">&nbsp;
										   		<br />
	                  			<b><%=PID%></b>
	                  		</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' valign='top' width='50%'>
										<p class="header">
											OFF SITE- FAX FORM TO<br />
                    	508-3636909
										</p>
									</td>
                  <td align='left' width='50%'>
                    <p class="header2">
	                  	IN-HOUSE- RETURN FORM TO INTERPRETER<br />
                    	SERVICES IMMEDIATELY AFTER<br />
                    	COMPLETING ASSIGNMENT
                 		</p>
                  </td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' width='50%'>
										<p class="header">
											Saint Vincent Hospital <br />
                    	Interpreter Services<br />
                      508-363-8695 / Ext. 28695
										</p>
									</td>
                  <td align='left' width='50%' style="border: 1px solid;" rowspan='5'>
                    <p class="header2">
	                  	Patient Name: <u><b><%=tmpName%></b></u>____________
	                  	<br /><br />
	                    Date of Birth: ______<u><b><%=tmpDOB%></b></u>______________
	                    <br /><br />
	                    Medical Record #: <u><b><%=mrrec%></b></u>____________________
	                    <br /><br />
	                   	Location of Appointment:  <u><b><%=tmpDeptaddr%></b></u>
                 		</p>
                  </td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' width='50%'>
										<p class="header">
											<table>
												<tr>
													<td>&#9744; SVH</td>
													<td>&#9744; Wellness Center</td>
													<td>&#9744; Ambulatory</td>
												</tr>
												<tr>
													<td>&#9744; Grove St.</td>
													<td>&#9744; Auburn</td>
													<td>&#9744; Pain Clinic</td>
												</tr>
											</table>
										</p>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' width='50%'>
										<p class="header">
											Language: <u><b><%=tmpSalita%></b></u>________<br />
                    	Department/Pt. Room #: <u><b><%=tmpdeptname%></b></u>
										</p>
									</td>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan='2' align='center'>
							<hr style="width: 95%;">
						</td>
					</tr>
				<% ElseIf tmpInst = 860 Then %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' height='80px' width='620px' bgcolor='#FFFFFF'  border='0' align='center'>
								<tr>
									<td align='right' colspan='2'>
										<table cellSpacing='0' cellPadding='0'  bgcolor='#FFFFFF'  border='0' align='right'>
											<tr>
												<td align='center' width='100%' style='border: solid 1px;'>
													<img src="_barcode.asp?code=<%=PID%>&height=20&width=2&mode=code39&text=0&fileout=<%=myPath%>" style="visibility:hidden" >
													<br />
										    	&nbsp;<img src="Images/BC/<%=myBC%>">&nbsp;
										   		<br />
	                  			<b><%=PID%></b>
	                  		</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='center' width='50%'>
										<p class="headerumass">
											UMASS MEMORIAL MEDICAL CENTER <br />
                    	INTERPRETER SERVICES<br />
                    	<font size='1'><i>for inquiries call 508-334-7651</i></font><br /><br />
                      <b>ENCOUNTER FORM</b>
										</p>
									</td>
                  <td align='left' width='50%' style="border: 1px solid;" rowspan='5'>
                    <p class="header2">
	                  	Patient Name: <u>_<b><%=tmpName%></b></u>____________
	                  	<br /><br />
	                  	Address:  <u>_<b><%=tmpDeptaddr%></b></u>
	                  	<br /><br />
	                    Date of Birth: <u>_<b><%=tmpDOB%></b></u>______________
	                    <br /><br />
	                    Medical Record #: <u>_<b><%=mrrec%></b></u>____________________
                 		</p>
                  </td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' width='50%'>
										<p class="header">
											<table width='100%'>
												<tr>
													<td>&#9744; University</td>
													<td>&#9744; ACC</td>
													<td>&#9744; Memorial</td>
													<td>&#9744; Hahnemann</td>
												</tr>
												<tr>
													<td colspan='4'>&#9744; Other:______________________________________</td>
												</tr>
											</table>
										</p>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='left' width='50%'>
										<p class="header">
											Language: <u>_<b><%=tmpSalita%></b></u>________<br />
                    	Location: <u>_<b><%=tmpdeptname%></b></u>
										</p>
									</td>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan='2' align='center'>
							<hr style="width: 95%;">
						</td>
					</tr>
				<% Else %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' height='120px' width='620px' bgcolor='#FFFFFF' border='5'>
								<tr>
									<td align='center'>
										<img src='images/LBISLOGOtrans.gif' align='center' height='60px'>
										<br>
										340&nbsp;Granite&nbsp;Street&nbsp;3<sup>rd</sup>&nbsp;Floor,&nbsp;Manchester,&nbsp;NH&nbsp;03102<br>
										Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
										<br><br>
										<p class='title'>Service Verification Form</p>
									</td>
								</tr>
								
							</table>
						</td>
					</tr>
				<% End If %>
				<tr><td>&nbsp;</td></tr>
				<% If tmpInst = 671 Then %>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='620px' bgcolor='#FFFFFF'  border='0' align='center'>
								<tr>
									<td valign='top' align='left'>
								   <b>ENCOUNTER INFORMATION</b>
				          </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Appointment Date/Time:<u><b><%=tmpAppDate%>&nbsp;<%=tmpAppTFrom%> - <%=tmpAppTTo%></b></u>
				        		</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;" rowspan='3'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td>&#9744; Outpatient</td>
													<td>&#9744; Inpatient</td>
												</tr>
											</table>
										</p>
                  </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Arrival Time:______________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Interpretation Begins:______________________________
				        		</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;" rowspan='3'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td>&#9744; Scheduled</td>
													<td>&#9744; ASAP</td>
												</tr>
												<tr>
													<td>&#9744; Phone</td>
													<td>&#9744; Rounds</td>
												</tr>
											</table>
										</p>
                  </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Encounter Ends:___________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td align='left' width='50%' valign='top' style="border: 1px solid;">
				        		<p class="header">
											<table style="width: 100%;">
												<tr>
													<td valign='top' align='left'>
												   <b>INTERPRETED FOR</b>
								          </td>
								        </tr>
								        <tr>
													<td>&#9744; ADMITTING/Registration</td>
													<td>&#9744; RN</td>
												</tr>
												<tr>
													<td>&#9744; ANESTHESIOLOGIST</td>
													<td>&#9744; MD</td>
												</tr>
												<tr>
													<td>&#9744; CRNA</td>
													<td>&#9744; TRIAGE</td>
												</tr>
												<tr>
													<td>&#9744; PA/NP</td>
													<td>&#9744; PT/OT</td>
												</tr>
												<tr>
													<td>&#9744; PCA</td>
													<td>&#9744; SW/CM</td>
												</tr>
												<tr>
													<td>&#9744; TECH</td>
													<td>&#9744; MENTAL HEALTH</td>
												</tr>
												<tr>
													<td>&#9744; OTHER</td>
													<td>&#9744; NUTRITIONIST</td>
												</tr>
												<tr>
													<td colspan='2'>
													______________________________
													</td>
												</tr>
											</table>
										</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;">
				        		<p class="header">
											<table style="width: 100%;">
												<tr>
													<td valign='top' align='left'>
												   <b>ENCOUNTER SUMMARY</b>
								          </td>
								        </tr>
								        <tr>
													<td>&#9744; CONSENT</td>
													<td>&#9744; PLAN OF CARE</td>
												</tr>
												<tr>
													<td>&#9744; DISCHARGE</td>
													<td>&#9744; PRE-ADMISSION TESTING</td>
												</tr>
												<tr>
													<td>&#9744; DIAGNOSIS/RESULTS</td>
													<td>&#9744; RADIOLOGY</td>
												</tr>
												<tr>
													<td>&#9744; FAMILY CONFERENCE</td>
													<td>&#9744; REGISTRATION</td>
												</tr>
												<tr>
													<td>&#9744; FOLLOW-UP CARE</td>
													<td>&#9744; ROUNDS</td>
												</tr>
												<tr>
													<td>&#9744; HISTORY</td>
													<td>&#9744; TEST/PROCEDURE</td>
												</tr>
												<tr>
													<td>&#9744; LAB</td>
													<td>&#9744; THERAPY</td>
												</tr>
												<tr>
													<td>&#9744; NURSING ADMISSION</td>
													<td>&#9744; TRAIGE</td>
												</tr>
												<tr>
													<td colspan='2'>&#9744; OTHER ______________________________</td>
												</tr>
											</table>
										</p>
				        	</td>
				        </tr>
				      </table>
			    	</td>
			    </tr>
					<tr>
						<td colspan='2' align='center'>
							<hr style="width: 95%;">
						</td>
					</tr>
          <tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='620px' bgcolor='#FFFFFF'  border='0' align='center'>
								<tr>
									<td valign='top' align='left'>
								   <b>MEDICAL PROVIDER INFORMATION</b>
				          </td>
				        </tr>
				        <tr>
				        	<td width='55%'>
				        		<p class="header">
				        			Name:_________________________________________
				        		</p>
				        	</td>
				        	<td align='left' width='45%' style="border: 1px solid;" rowspan='7'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td valign='top' align='center'>
												   <b>Provider Signature Required</b>
								          </td>
								        </tr>
												<tr><td>&#9744; Interpretation Completed</td></tr>
												<tr><td>&#9744; Provider Speaks Patient's Language</td></tr>
												<tr><td>&#9744; Request Cancelled</td></tr>
												<tr><td>&#9744; Seen w/o Interpreter</td></tr>
												<tr><td>&#9744; Patient No Show</td></tr>
												<tr><td>&#9744; Scheduling Error</td></tr>
											</table>
										</p>
                  </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Signature:_________________________<u>Time:</u>________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
									<td valign='top' align='left'>
								   <b>INTERPRETER INFORMATION</b>
				          </td>
				        </tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Name:<u>_<b><%=tmpInName%></b></u>________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Signature:_________________________<u>Time:</u>________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
									<td style='font-size: 7pt;'>
										<b>
											It is the policy of Saint Vincent Hospital to provide Interpreter<br />
											Services free of charge to all patients who do not speak English<br />
											or prefer to speak in their primary language. If you do not want<br />
											an interpreter, please sign	on the line below. Your signature<br />
											indicates that you have given	permission to release the hospital<br />
											appointed intepreter from your care.
										</b>
									</td>
									<td style='font-size: 7pt;'>
										&#9744; Patient Declines Interpeter Services (Patient Sign Below)<br /><br />
										____________________________________________<br />
										&#9744; Patient brought own interpeter<br /><br />
										&#9744; Other_____________________________________
									</td>				        
				        </tr>
				      </table>
			    	</td>
			    </tr>
			  <% ElseIf tmpInst = 860 Then %>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='620px' bgcolor='#FFFFFF'  border='0' align='center'>
								<tr>
									<td valign='top' align='center' colspan='2'>
								   <b>ENCOUNTER INFORMATION</b>
				          </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Requested Date/Time:<u>_<b><%=tmpAppDate%>&nbsp;<%=tmpAppTFrom%> - <%=tmpAppTTo%></b></u>
				        		</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;" rowspan='3'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td>&#9744; Outpatient</td>
													<td>&#9744; Inpatient</td>
												</tr>
											</table>
										</p>
                  </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Arrival Time:______________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Interpretation Begins:______________________________
				        		</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;" rowspan='3'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td>&#9744; Scheduled</td>
													<td>&#9744; ASAP</td>
												</tr>
												<tr>
													<td>&#9744; Phone</td>
													<td>&#9744; Rounds</td>
												</tr>
											</table>
										</p>
                  </td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Encounter Ends:___________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
				        	<td align='left' width='50%' valign='top'>
				        		<p class="header">
											<table style="width: 100%;">
												<tr>
													<td valign='top' align='center'>
												   <b>LIST EACH PROVIDER (w/ Credentials)</b>
								          </td>
								        </tr>
								        <tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
												<tr>
													<td>__________________________________________________</td>
												</tr>
											</table>
										</p>
				        	</td>
				        	<td align='left' valign='top' width='50%' style="border: 1px solid;">
				        		<p class="header">
											<table style="width: 100%;">
								        <tr>
													<td>&#9744; MD</td>
													<td>&#9744; PA / NP</td>
													<td>&#9744; RN / LPN</td>
												</tr>
												<tr>
													<td>&#9744; PCA</td>
													<td>&#9744; Admission</td>
													<td>&#9744; Discharge</td>
												</tr>
												<tr>
													<td>&#9744; Mental Health</td>
													<td>&#9744; PT / OT</td>
													<td>&#9744; SW / CM</td>
												</tr>
												<tr>
													<td>&#9744; Procedure</td>
													<td>&#9744; Test</td>
													<td>&#9744; Family Mtg</td>
												</tr>
												<tr>
													<td colspan='3'>&#9744; Informed Consent: ______________________________</td>
												</tr>
												<tr>
													<td colspan='3'>&nbsp;&nbsp;&nbsp;&nbsp;______________________________________________</td>
												</tr>
												<tr>
													<td colspan='3'>&nbsp;&nbsp;&nbsp;&nbsp;______________________________________________</td>
												</tr>
													<tr>
													<td colspan='3'>&#9744; Other: ________________________________________</td>
												</tr>
												<tr>
													<td colspan='3'>&nbsp;&nbsp;&nbsp;&nbsp;______________________________________________</td>
												</tr>
											</table>
										</p>
				        	</td>
				        </tr>
				      </table>
			    	</td>
			    </tr>
					<tr>
						<td colspan='2' align='center'>
							<hr style="width: 95%;">
						</td>
					</tr>
          <tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='620px' bgcolor='#FFFFFF' align='center'>
								<tr>
									<td valign='top' align='left'>
								   <b>MEDICAL PROVIDER INFORMATION</b>
				          </td>
				        </tr>
				        <tr>
				        	<td width='55%'>
				        		<p class="header">
				        			Name (Print):_____________________________________
				        		</p>
				        	</td>
				        	<td align='left' width='50%' style="border: 1px solid;" rowspan='7'>
                    <p class="header">
											<table style="width: 100%;">
												<tr>
													<td valign='top' align='center' colspan='2'>
												   <b><i>Provider Signature Required</i></b>
								          </td>
								        </tr>
												<tr><td colspan='2'>&#9744; Interpretation Completed</td></tr>
												<tr><td colspan='2'>&#9744; Provider Speaks Patient's Language</td></tr>
												<tr><td colspan='2'>&#9744; Patient declines Interpreter Services (sign below)</td></tr>
												<tr>
													<td valign='top' align='center' colspan='2'>
												   <b><i>No Signature Required</i></b>
								          </td>
								        </tr>
												<tr>
													<td>&#9744; Request Cancelled</td>
													<td>&#9744; Patient no show</td>
												</tr>
												<tr>
													<td>&#9744; Scheduling error</td>
													<td>&#9744; Seen w/o Interpreter</td>
												</tr>
												<tr><td colspan='2'>&#9744; Telephonic interpretation utilized</td></tr>
												<tr><td colspan='2'>&#9744; Interpreter not needed at this time</td></tr>
												<tr><td colspan='2'>&#9744; Patient rounds completed</td></tr>
											</table>
										</p>
                  </td>
				        </tr>
				        
				        <tr>
				        	<td>
				        		<p class="header">
				        			Signature:_______________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
									<td valign='top' align='left'>
								   <b>INTERPRETER INFORMATION</b>
				          </td>
				        </tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Name:<u>_<b><%=tmpInName%></b></u>________
				        		</p>
				        	</td>
				        </tr>
				        <tr>
				        	<td>
				        		<p class="header">
				        			Signature:_______________________________________
				        		</p>
				        	</td>
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				       </table>
			    	</td>
			    </tr>
					<tr>
						<td colspan='2' align='center'>
							<hr style="width: 95%;">
						</td>
					</tr>
          <tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='620px' bgcolor='#FFFFFF' align='center' >
				        <tr>
					        <td colspan='2'>
										&#9744; Patient states speaks English<br />
										&#9744; Patient brought own interpeter<br />
										&#9744; Other_____________________________________
									</td>				        
				        </tr>
				        <tr><td>&nbsp;</td></tr>
				        <tr>
									<td colspan='2'>
										<p>
											It is the policy of UMass Memorial Hospital to provide a hospital interpreter
											to all patients who do not speak English or prefer to speak in their primary language. 
											If you do not want an interpreter, please sign on the line below. Your signature
											indicates that you have given	permission to release the hospital
											appointed intepreter from your care.
										</p>
									</td>
									 <tr><td>&nbsp;</td></tr>
									<tr>
										<td align='center'>________________________________________________________________</td>
										<td align='center'>________________________</td>
									</tr>
									<tr>
										<td align='center'>Patient Signature</td>
										<td align='center'>Date</td>
									</tr>
				      </table>
			    	</td>
			    </tr>
				<% Else %>
				<% If tmpClass = 3 Then %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='2' align='center' class='header' style='border:1px solid #000000;' bgcolor='#FFFFFF'>
										Court Assignment
										<p class='notes'>To be completed by Language Bank Staff</p>
									</td>
								</tr>
								<tr>
									<td class='printForm'>
										Project ID:&nbsp;&nbsp;<b><%=PID%></b>
										
										<img src="_barcode.asp?code=<%=PID%>&height=20&width=2&mode=code39&text=0&fileout=<%=myPath%>" style="visibility:hidden" >
									</td>
									<td class='printForm' Style='text-align: center;'>	
										<br>
										&nbsp;<img src="Images/BC/<%=myBC%>">&nbsp;
										<br><br>
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
										Time of Service:
										<b><%=tmpAppTFrom%></b><br>
										Requested Time:
										<b><%=tmpAppTFrom%> - <%=tmpAppTTo%></b>
										<br><font size='1'>*appointment could be longer due to unforeseen situations</font>
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
									<td colspan='3' align='center' class='header' style='border:1px solid #000000;' bgcolor='#FFFFFF'>
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
								<tr height='50px'>
									<td class='printForm' valign='top' style='border-right-color:#FFFFFF;'>
										Signature of Interpreter:
									</td>
									<td class='printForm' colspan='2' valign='top' style='border-left-color:#FFFFFF;'>
										Date:
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm' style='border-right-color:#FFFFFF;'>
										Assignment Start Time:
									</td>
									<td class='printForm' colspan='2' style='border-left-color:#FFFFFF;'>
										Assignment End Time:
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
								<tr bgcolor='#FFFFFF'>
									<td colspan='3' align='center' style='border:1px solid #000000;'>
											<table cellSpacing='0' cellPadding='0' width='620px' bgColor='#FFFFFF' border='0'>
												<tr>
													<td colspan='3' class='printForm'>
														For Billing Department - Billing Address<br>
														&nbsp;&nbsp;&nbsp;&nbsp;
														<u>Dale Trombley, Administrative Office of the Courts, 2 Noble Drive, Concord, NH 03301</u>
													</td>
												</tr>
												<tr height='25px' bgcolor='#FFFFFF'>
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
								<tr>
									<td class='printForm' colspan='3' style='text-align: center;'><i>*This form contains information that must be protected and kept at secure place at all times. Please make sure that the form is kept at place that you only have access to. If verification forms are saved on your computer, please make sure you are the only person who has access to folder(s) where forms are saved. After you fax or emails us a copy of the form, please make sure to shred/delete the form</i> </td>
									
								</tr>
							</table>
						</td>
					</tr>
				<% Else %>
					<tr>
						<td valign='top' align='center'>
							<table cellSpacing='0' cellPadding='0' width='620px' bgColor='white' border='0'>
								<tr>
									<td colspan='2' align='center' class='header' style='border:1px solid #000000;' bgcolor='#FFFFFF'>
										Assignment
										<p class='notes'>To be completed by Language Bank Staff</p>
									</td>
								</tr>
								<tr>
									<td class='printForm' style="width: 310px; text-align: center;">
										
										<b><%=PID%></b>
										
										<img src="_barcode.asp?code=<%=PID%>&height=20&width=2&mode=code39&text=0&fileout=<%=myPath%>" style="visibility:hidden">

									</td>
									<td class='printForm' Style='text-align: center;'>	
										<br>
										&nbsp;<img src="Images/BC/<%=myBC%>">&nbsp;
										<br><br>
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
										Service Recipient (Client):
										&nbsp;&nbsp;<b><%=tmpName%></b>
									</td>
									<td class='printForm' valign='top'>
										Language:
										&nbsp;&nbsp;<b><%=tmpSalita%></b>
									</td>
								</tr>
								<tr>
									<td class='printForm'>
										Patient MR#:
										&nbsp;&nbsp;<b><%=mrrec%></b>
									</td>
									<td class='printForm'>
										Date of Appointment:
										&nbsp;&nbsp;<b><%=tmpAppDate%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Type of Appointment:
										&nbsp;&nbsp;<b><%=tmpClassName%></b>
									</td>
									<td class='printForm'>
										Time of Service:
										<b><%=tmpAppTFrom%></b>
									</td>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Agency Requesting Service:<br>
										&nbsp;&nbsp;<b><%=tmpIname%></b>
									</td>
									
									<td class='printForm'>
										Requested Time:
										<b><%=tmpAppTFrom%> - <%=tmpAppTTo%></b>
										<br><font size='1'><i>*appointment could be longer due to unforeseen situations</i></font>
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
											Special Circumstances/Precautions:<br>
											&nbsp;&nbsp;<b><%=tmpSC%></b>
										</td>
									<%Else%>
										<td class='printForm' valign='top'>
											Reasons / Special Circumstances/Precautions:<br>
											&nbsp;&nbsp;<b><%=tmpReason%></b><br>
											&nbsp;&nbsp;<b><%=tmpSC%></b>
										</td>
									<%End If%>
								</tr>
								<tr height='25px'>
									<td class='printForm'>
										Client Phone:
										&nbsp;&nbsp;<b><%=tmpCfon%></b>
										<br>
										Client Alter. Phone:
										&nbsp;&nbsp;<b><%=tmpCAfon%></b>
										<br>
										<%=courtcall%>
										<br>
										<%=leavemsg%>
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
									<td colspan='3' align='center' class='header' style='border:1px solid #000000;' bgcolor='#FFFFFF'>
										Job Verification
										<p class='notes'>
											To be completed at time of service
										</p>
									</td>
								</tr>
								<tr height='25px'>
									<!--<td class='printForm' width='300px' style='border-right-color:#FFFFFF;'>
										Date of Service:
									</td>//-->
									<td class='printForm' colspan='3'>
										Notes:&nbsp;&nbsp;<b><%=strNotes%></b>
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
								<tr height='25px'>
									<td class='printForm' style='border-right-color:#FFFFFF;'>
										Assignment Start Time:
									</td>
									<td class='printForm' colspan='2' style='border-left-color:#FFFFFF; width: 200px;'>
										Assignment End Time:
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
								<!--<tr height='25px'>
									<td class='printForm' colspan='3'>
										Billing:
										&nbsp;<b><%=tmpBContact%></b><br>
										&nbsp;&nbsp;<b><%=tmpBaddr%></b>
									</td>
								</tr>//-->
								<tr>
									<td colspan='3'>	
										<table cellspacing='0' cellpadding='0' width='100%' border='0'>
											
											<tr>
												<td class='printForm' colspan='3' style='text-align: center;'><i>*This form contains information that must be protected and kept at secure place at all times. Please make sure that the form is kept at place that you only have access to. If verification forms are saved on your computer, please make sure you are the only person who has access to folder(s) where forms are saved. After you fax or emails us a copy of the form, please make sure to shred/delete the form</i> </td>
												
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% End If %>
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
