<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'SAVE TIME EMAIL WAS SENT
If Request("sino") = 0 Then
	sqlSent = "UPDATE [request_T] SET SentReq = '" & Now & "' WHERE [index] = " & Request("HID")
ElseIf Request("sino") = 1 Then 
	sqlSent = "UPDATE [request_T] SET SentIntr = '" & Now & "' WHERE [index] = " & Request("HID")
End If
Server.ScriptTimeout = 3000
If Request("sino") < 2 Then
	Set rsSent = Server.CreateObject("ADODB.RecordSet")
	rsSent.Open sqlSent, g_strCONN, 1, 3
	Set rsSent = Nothing
End If
'GET REQUEST INFO
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM request_T WHERE [index] = " & Request("HID")
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	CliName = rsReq("Clname") & ", " & rsReq("Cfname")
	IntrName = GetIntr(rsReq("IntrID"))
	LangName = GetLang(rsReq("LangID"))
	LangID = rsReq("LangID")
	AppFrame = rsReq("appDate") & " (" & rsReq("appTimeFrom") & " - " & rsReq("appTimeTo") & ")" 
	AppDate = rsReq("appDate")
	InstID = rsReq("InstID")
	DeptID = rsReq("DeptID")
	'tmpDOB = rsReq("DOB")
	'tmpComment = rsReq("Comment")
	mrrec = rsReq("mrrec")
	ReqName = GetReq(rsReq("ReqID"))
	timestamp = rsReq("timestamp")
	tmpOther = rsReq("DocNum") & ",  " & rsReq("CrtRumNum")

	tmpdept =  GetDept(rsReq("DeptID"))
	tmpCon = rsReq("SentReq")
	If rsReq("CliAdd") = True Then InstAddr =  rsReq("CAddress") & ", " & rsReq("CliAdrI") & ", " & rsReq("CCity") & ", " & rsReq("CState") & ", " & rsReq("CZip")
	If rsReq("CliAdd") = True Then SubCity = rsReq("CCity")
	tmpcomintr = rsReq("intrcomment")
	tmpHPID = rsReq("HPID")
	tmpDecTT = z_fixNull(rsReq("actTT"))
	tmpDecMile = z_fixNull(rsReq("actMil"))
	tmpclaim = rsReq("claimant")
	tmpjudge = rsReq("judge")
	tmpIntr = GetIntr(rsReq("IntrID"))
	
End If
rsReq.Close
Set rsReq = Nothing
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & InstID
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	InstName = rsInst("Facility")
	subInst = rsInst("Facility")
End If
rsInst.Close
Set rsInst = Nothing
Set rsDept= Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & DeptID
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	InstName = InstName & " - " & rsDept("dept")
	If InstAddr = "" Then InstAddr = rsDept("Address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " & rsDept("State") & ", " & rsDept("Zip")
	If SubCity = "" Then SubCity = rsDept("City")
	BillAddr =  rsDept("BAddress") &", " & rsDept("BCity") & ", " & rsDept("BState") & ", " & rsDept("BZip")
	tmpBContact = rsDept("Blname") & ", " & rsDept("Bfname")
End If
rsDept.Close
Set rsDept = Nothing
If Z_CZero(tmpHPID) <> 0 Then
	Set rsHP = Server.CreateObject("ADODB.RecordSet")
	sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
	rsHP.Open sqlHP, g_strCONNHP, 3, 1
	If Not rsHP.EOF Then
		If rsHP("reqName") <> "" Then ReqName = rsHP("reqName")
	End If
	rsHP.CLose
	Set rsHP = Nothing
End If
'SEND EMAIL

Set mlMail = zSetEmailConfig()
strBody = ""
If Left(Request.ServerVariables("REMOTE_ADDR"), 11) = "192.168.111" Or _
				Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "::1" Then 
	mlMail.To = "hagee@zubuk.com"
	strBody = "<p>TO BE SENT TO: <b>" & FixEmail(Request("emailadd")) & "</b></p>" & vbCrLf
Else
	mlMail.To = FixEmail(Request("emailadd"))
	mlMail.Cc = "language.services@thelanguagebank.org"
	mlMail.Bcc = "sysdump1@ascentria.org"
End If
If Request("sino") = 0 Then 'FOR REQUESTOR
	mlMail.From = "language.services@thelanguagebank.org"
	mlMail.Subject= "Interpreter Confirmation - The Language Bank"
	strBody = strBody & "<table cellpadding='0' cellspacing='0' border='0' align='center'>" & vbCrLf & _
			"<tr><td align='center'>" & vbCrLf & _
				"<img src='https://languagebank.lssne.org/lsslbis/images/LBISLOGOBandW.jpg'>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
			"<tr><td align='center'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'><b>Appointment Confirmation:</b></font><br>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Appointment Requested by:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & ReqName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date of Request:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & timestamp & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date and Time of Confirmation:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpCon & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _

					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Interpreter:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpIntr & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _

				"</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right' width='225px'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Project ID Number:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & Request("HID") & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Client Name:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & CliName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>MR #:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & mrrec & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date of Appointment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & AppFrame & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Address of Appointment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & InstAddr & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Language:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & LangName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Department:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpdept & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Doc Num / Court / Delivery Ticket:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpOther & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Comment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpComment & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf
				If InstID = 757 Or InstID = 777 Then 'SSA
					strBody = strBody & "<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Claimant:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpclaim & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf & _
						"<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Judge:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpjudge & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf & _
						"<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Interpreter:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpIntr & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf
				End If
				strBody = strBody & "</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td align='left'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'>The request for the above appointment has been received and confirmed.  A Language Bank Interpreter will be<br>" & vbCrlf & _
				"present for this appointment.  If any of the above information is not correct, changes or you have additional questions,<br>" & vbCrLf & _
				"please contact the LanguageBank office immediately at 410-6183 or email us at <a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>"  & vbCrlf & _
				"Please refer to the Project ID Number when calling or emailing the office. If there are any difficulties in completing<br>" & vbCrLf & _
				"this assignment you will be notified.<br><br>" & vbCrLf & _
				"Language Bank Cancellation Policy:<br>" & vbCrLf & _
				"Under the following conditions you, or your agency, will still be responsible for full payment to The Language Bank:<br>" & vbCrLf & _
				"&nbsp;&nbsp;&nbsp;&nbsp;• If you or your agency cancels a request with less than 24-hour notice prior to the scheduled service.<br>" & vbCrLf & _
		    "&nbsp;&nbsp;&nbsp;&nbsp;• If a patient cancels or reschedules an appointment with less than 24-hour notice prior to the scheduled service.<br>" & vbCrLf & _
		    "&nbsp;&nbsp;&nbsp;&nbsp;• If a patient does not show up for scheduled appointment.<br>" & vbCrLf & _
				"By making this appointment you agree to our cancellation policy.<br><br>" & vbCrLf & _
				"Emergency Request (less than 24 hours notice):<br>" & vbCrLf & _
				"&nbsp;&nbsp;&nbsp;&nbsp;• Requests received less than 24 hours notice are subject to additional $20.00 fee<br><br>" & vbCrLf & _
				"DISCLAIMER: This cancellation policy does not apply to American Sign Language (ASL) appointments as Language Bank must<br>" & vbCrlf & _
				"follow NH regulations/rules related to ASL interpretation. For questions, please contact Language Bank.<br><br>" & vbCrLf & _
				"Thank you for using LanguageBank for your interpretation needs.</font><br><br>" & vbCrLf & _
				"<font size='2' face='Script MT Bold'><i>Alen Omerbegovic</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'><i><b>Alen Omerbegovic</b><br>" & vbCrLf & _
				"Program Manager<br>" & vbCrLf & _
				"LanguageBank<br>" & vbCrLf & _
				"Ascentria Care Alliance<br>" & vbCrLf & _ 
				"340 Granite Street, 3rd Floor <br>" & vbCrLf & _ 
				"Manchester, NH 03102 <br>" & vbCrLf & _ 
				"603-410-6183  603-410-6186 fax<br><br>" & vbCrLf & _ 
				"PLEASE NOTE: This email/fax is intended only for the use of the individual or entity to which it is addressed, and may contain<br>" & vbCrLf & _
				"information that is privileged, confidential and exempt from disclosure under applicable law.  If you are not the intended recipient,<br>" & vbCrLf & _
				"then dissemination, distribution or copying of this communication is strictly prohibited.  If you have received this communication in<br>" & vbCrLf & _
				"error, please notify LSS immediately at 1-800-244-8119 and return the original email to us at the above address." & vbCrLf & _
				"</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
		"</table>"
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
ElseIf Request("sino") = 1 Or Request("sino") = 3 Then 'FOR INTERPRETER
	'travelTIME and Mileage rules
	tmpMMT = Split(Request("MileTT"), "|")
	If tmpDecTT = "" Then
		tmpTravel = Replace(tmpMMT(1), "'", "") 
	Else
		tmpTravel = tmpDecTT
	End If
	If tmpDecMile = "" Then
		tmpMile = Replace(tmpMMT(0), "'", "") 
	Else
		tmpMile = tmpDecMile
	End If
	mlMail.From = "DO-NOT-REPLY@thelanguagebank.org"
	strSubj = "[LBIS]Appointment Assignment - " & AppDate & " - " & subInst & " - " & SubCity
	mlMail.Subject = strSubj
	Set theDoc = Server.CreateObject("ABCpdf6.Doc") 'converts html to pdf
	intInstID = Z_GetInfoFROMAppID(Request("HID"), "InstID")
	If intInstID = 860 Then 'UMASS Memorial Medical Center
		attachPDF = pdfStr & "VerificationForm" & Request("HID") & "UM.pdf"
		strUrl = "http://webserv6/lsslbis/umass-body.asp?ReqID=" & Request("HID")
	Else
		attachPDF = pdfStr & "VerificationForm" & Request("HID") & ".pdf"
		strUrl = "http://webserv6/lsslbis/print.asp?PDF=1&ID=" & Request("HID")
		'strUrl = "https://localhost/interpreter/print.asp?PDF=1&ID=" & Request("HID")
	End If

	thedoc.HtmlOptions.PageCacheClear
	theDoc.HtmlOptions.RetryCount = 3
	theDoc.HtmlOptions.Timeout = 120000
	theDoc.Pos.X = 10
	theDoc.Pos.Y = 10
	theID = theDoc.AddImageUrl(strUrl)
	If intInstID = 671 Then 'saint vincent
		theDoc.FontSize = 12 ' big text
		theDoc.rect.Move 50, -50
		theDoc.Page = theDoc.AddPage(1)
		theText = "<b>ATTENTION INTERPRETERS</b><br><br><br><br>" & _
			"When handling appointments at St. Vincent’s Hospital, you must follow the below procedures:<br><br><br>" & _ 
			"1)	BEFORE THE APPOINTMENT: Go to the mailroom. The mailroom is by the Loading Dock/Receiving Area.<br><br>" & _ 
			"2) Ask for Ms. Fran Goulet (phone number 508-363-9310 if you need to contact her.)<br><br>" & _ 
			"3) Sign the log book, take the badge Ms. Goulet hands you, and continue on to the appointment.<br><br>" & _ 
			"4) After the appointment, make sure that all parts of the V-Form are correctly filled out. If any part of the<br>" & _
			"V-Form is incomplete, we cannot bill St. Vincent’s!<br><br>" & _ 
			"5)	Once the V-Form is complete, leave it with the Interpretation Department on the ground floor. DO NOT ASK<br>" & _
			"FOR A COPY. We do not need copies of V-Forms from St. Vincent’s.<br><br><br>" & _ 
			"THANK YOU VERY MUCH FOR FOLLOWING THIS PROCEDURE."
		theDoc.AddHtml(theText)
	ElseIf intInstID = 849 Then 'lowel
		theDoc.FontSize = 12 ' big text
		theDoc.rect.Move 50, -50
		theDoc.Page = theDoc.AddPage(1)
		theText = "Instructions for assignments at Lowell General Hospital:<br><br><br><br>" & _
			"1.	Assignments are scheduled for 2 hours.  You must be available to stay for the full 2-hours, since new interpreting<br>" & _
			"sessions can be assigned to us little or no notice, and we may need you to complete them.<br><br>" & _
			"2.	Upon completion of an interpreting assignment, please contact Interpreter Services by dialing extension 64710 or<br>" & _
			"64709 (Saints Campus) or extension 76591 (Main Campus) for further instructions.<br><br>" & _
			"3.	If the duration of the assignment is expected to exceed 2-hours, please call the Interpreter Services office at<br>" & _
			"extension 64710 or 64709 (Saints) or extension 76591 (Main) for approval to stay longer.<br><br>" & _
			"4.	If an appointment is cancelled or the patient does not show up, please dial extension 64710 or 64709 (Saints) or<br>" & _
			"extension 76591 (Main) for further instructions. We may need you for another appointment elsewhere in the hospital.<br><br>" & _
			"5.	Upon completion of an appointment at one of our satellite clinics, please call the Interpreter Services office at<br>" & _
			"extension 64710 or 64709 (Saints) or extension 76591 (Main) to provide us with information, especially when the<br>" & _
			"appointment did not go as planned, the patient didn’t show, it started late, etc.<br><br>" & _
			"&nbsp;&nbsp;&nbsp;&nbsp;•	If nobody is available to take your call when you contact the Interpreter Services office, please leave<br>" & _
			"a message with details about the appointment.<br><br>" & _
			"Main Office 978-937-6591"
		theDoc.AddHtml(theText)
	ElseIf intInstID = 860 Then
'		Set theDoc2 = Server.CreateObject("ABCpdf6.Doc") 'converts html to pdf
'		theDoc2.Read(DirectionPath & "Instructions for Interpreters at UMass pdf version 10.10.17.pdf")
'		theDoc.Append(theDoc2)
'		Set theDoc2 = Nothing

		Set theDoc3 = Server.CreateObject("ABCpdf6.Doc")
		theDoc3.Read(DirectionPath & "umass_encounter_form.2018.pdf")
		theDoc.Append(theDoc3)

	End If

	Do
	  If Not theDoc.Chainable(theID) Then Exit Do
	  theDoc.Page = theDoc.AddPage()
	  theID = theDoc.AddImageToChain(theID)
	Loop
	
	For i = 1 To theDoc.PageCount
	  theDoc.PageNumber = i
	  theDoc.Flatten
	Next

	theDoc.Save attachPDF
	strBody = "<font size='2' face='trebuchet MS'>A request has been assigned to you.<br><br>Attached is the verification form for the request. Please fill-out the form  upon completion." & vbCrlf & _
		"If there are any questions or clarifications, please contact the LanguageBank office immediately at 410-6183 or email us at " & vbCrLf & _
		"<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>."& _
		"This email was initiated by " & Request.Cookies("LBUsrName") & "</font><br><br>" & vbCrLf & _
		"<font size='2' face='trebuchet MS'><b>Payable Mileage: " & Z_FormatNumber(tmpMile, 2) & " Miles<br>" & vbCrLf & _
		"Payable Travel Time: " & Z_FormatNumber(tmpTravel, 2) & " Hrs.</b><br><br></font>" & vbCrLf
		If InstID = 108 Then
			strBody = strBody & "<font size='2' face='trebuchet MS'>Please give the survey (attached form with the verification form) to the client you interpreted for and after he/she is done they need to give it to the DHHS Worker (Case Worker). Please do not stay around while client is filling out the survey. If survey you receive is in English, please take a moment to read the survey to the client and have him answer as soon as you leave.</font><br><br>" & vbCrLf
		End If
		If InstID = 323 And DeptID = 1924 Then
			strBody = strBody & "<font size='2' face='trebuchet MS'>Directions included (attached form with the verification form).</font><br><br>" & vbCrLf
		End If
		If InstID = 860 Then
			strBody = strBody & "<font size='2' face='trebuchet MS'>Please read the guidelines first (attached form with the verification form).</font><br><br>" & vbCrLf
		End If
		strBody = strBody & "<font size='2' face='trebuchet MS'><b>Comment:</b></font><font size='2' face='trebuchet MS' color='red'><b><i> " & tmpcomintr & "</i></b></font><br><br>" & vbCrLf & _
		"<font size='2' face='trebuchet MS'>The Language Bank</font><br><br>" & vbCrLf & _
		"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>"
	mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
	'create ICS
	'If CreateICS(Request("HID"), strSubj) Then
	'	mlMail.AddAttachment CalPath & Request("HID") & ".ICS"
	'End If
	mlMail.AddAttachment attachPDF
	If InstID = 108 Then 
		mlMail.AddAttachment SurveyPath & GetLangSurvey(LangID)
	ElseIf InstID = 323 And DeptID = 1924 Then 
		mlMail.AddAttachment DirectionPath & "DirWDH-CNS.pdf"
	'ElseIf InstID = 860 Then 
	'	mlMail.AddAttachment DirectionPath & "Instructions for Interpreters at UMass pdf version 10.10.17.pdf"
	'	mlMail.AddAttachment DirectionPath & "Interpreters guidelines.pdf"
	End If
End If

mlMail.Send
set mlMail = Nothing
If SaveHist(Request("HID"), "email.asp") Then
	
		End If
'CREATE LOG
Set fso = CreateObject("Scripting.FileSystemObject")
Set LogMe = fso.OpenTextFile(EmailLog, 8, True)
strLog = Now & vbCrLf & "----- EMAIL SENT -----" & vbCrLf & _
	"TO: " & Request("emailadd") & vbCrLf & _
	"REQUEST ID: " & Request("HID")
LogMe.WriteLine strLog
Set LogMe = Nothing
Set fso = Nothing

Session("MSG") = "E-Mail was sent to " &  Replace(Request("emailadd"), ";", " and " ) & "."
Response.Redirect "reqconfirm.asp?ID=" & Request("HID")
%>