<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'SAVE TIME EMAIL WAS SENT
If Request("sino") = 0 Then
	sqlSent = "UPDATE request_T SET SentReq = #" & Now & "# WHERE index = " & Request("HID")
ElseIf Request("sino") = 1 Then 
	sqlSent = "UPDATE request_T SET SentIntr = #" & Now & "# WHERE index = " & Request("HID")
End If
Set rsSent = Server.CreateObject("ADODB.RecordSet")
rsSent.Open sqlSent, g_strCONN, 1, 3
Set rsSent = Nothing
'GET REQUEST INFO
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM request_T WHERE index = " & Request("HID")
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	CliName = rsReq("Clname") & ", " & rsReq("Cfname")
	IntrName = GetIntr(rsReq("IntrID"))
	LangName = GetLang(rsReq("LangID"))
	AppFrame = rsReq("appDate") & " (" & rsReq("appTimeFrom") & " - " & rsReq("appTimeTo") & ")" 
	InstID = rsReq("InstID")
	DeptID = rsReq("DeptID")
	tmpComment = rsReq("Comment")
	ReqName = GetReq(rsReq("ReqID"))
	timestamp = rsReq("timestamp")
	If rsReq("DocNum") <> "" Or rsReq("CrtRumNum") <> "" Then
		tmpOther = rsReq("DocNum") & ",  " & rsReq("CrtRumNum")
	Else
		tmpOther = GetDept(rsReq("DeptID"))
	End If
	tmpCon = rsReq("SentReq")
	If rsReq("CliAdd") = True Then InstAddr =  rsReq("CAddress") &", " & rsReq("CCity") & ", " & rsReq("CState") & ", " & rsReq("CZip")
End If
rsReq.Close
Set rsReq = Nothing
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE index = " & InstID
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	InstName = rsInst("Facility")
End If
rsInst.Close
Set rsInst = Nothing
Set rsDept= Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE index = " & DeptID
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	InstName = InstName & " - " & rsDept("dept")
	If InstAddr = "" Then InstAddr = rsDept("Address") &", " & rsDept("City") & ", " & rsDept("State") & ", " & rsDept("Zip")
	BillAddr =  rsDept("BAddress") &", " & rsDept("BCity") & ", " & rsDept("BState") & ", " & rsDept("BZip")
	tmpBContact = rsDept("Blname") & ", " & rsDept("Bfname")
End If
rsDept.Close
Set rsDept = Nothing
'SEND EMAIL
Set mlMail = Server.CreateObject("CDONTS.NewMail")
mlMail.bodyformat = 0
mlMail.mailformat = 0
mlMail.To = Replace(Request("emailadd"),"'","")
mlMail.Cc = "language.services@thelanguagebank.org"
'mlMail.Bcc = "patrick@zubuk.com;paul@zubuk.com;aomerbegovic@lssnorth.org"
mlMail.From = "language.services@thelanguagebank.org"
mlMail.Subject="Interpreter Request - The Language Bank"
If Request("sino") = 0 Then 'FOR REQUESTOR
	strBody = "<table cellpadding='0' cellspacing='0' border='0' align='center'>" & vbCrLf & _
			"<tr><td align='center'>" & vbCrLf & _
				"<img src='http://web03.zubuk.com/lss-lbis-staging/images/LBISLOGOBandW.jpg'>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
			"<tr><td align='center'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'><b>Appointment Confirmation:</b></font><br>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right' width='225px'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Requesting Facility and Department:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & InstName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
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
							"<font size='2' face='trebuchet MS'>Other:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpOther & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
				"</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right' width='225px'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Billing Contact:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpBContact & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Billing Address:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & BillAddr & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
				"</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td align='left'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'>The request for the above appointment has been received and confirmed.  A Language Bank Interpreter<br>" & vbCrLf & _
				"will be present for this appointment.  If any of the above information is not correct, changes or you have<br>" & vbCrLf & _
				"additional questions, please contact the LanguageBank office immediately at 410-6183 or email us at<br>" & vbCrLf & _
				"<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>. Please refer to the Project ID Number when calling or emailing the office.<br>" & vbCrLf & _
				"If there are any difficulties in completing this assignment you will be notified.<br><br>" & vbCrLf & _
				"Thank you for using LanguageBank for your interpretation needs.</font><br><br>" & vbCrLf & _
				"<font size='2' face='Script MT Bold'><i>Alen Omerbegovic</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'><i><b>Alen Omerbegovic</b><br>" & vbCrLf & _
				"Program Manager<br>" & vbCrLf & _
				"LanguageBank<br>" & vbCrLf & _
				"Lutheran Social Services<br>" & vbCrLf & _ 
				"261 Sheep Davis Road, Suite A-1<br>" & vbCrLf & _ 
				"Concord, NH 03301<br>" & vbCrLf & _ 
				"603-410-6183  603-410-6186 fax<br><br>" & vbCrLf & _ 
				"PLEASE NOTE: This email/fax is intended only for the use of the individual or entity to which it is addressed, and may contain<br>" & vbCrLf & _
				"information that is privileged, confidential and exempt from disclosure under applicable law.  If you are not the intended recipient,<br>" & vbCrLf & _
				"then dissemination, distribution or copying of this communication is strictly prohibited.  If you have received this communication in<br>" & vbCrLf & _
				"error, please notify LSS immediately at 1-800-244-8119 and return the original email to us at the above address.  We will<br>" & vbCrLf & _
				"reimburse you for postage.</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
		"</table>"
		mlMail.Body = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
ElseIf Request("sino") = 1 Then 'FOR INTERPRETER
	Set theDoc = Server.CreateObject("ABCpdf6.Doc")
	attachPDF = pdfStr & "VerificationForm" & Request("HID") & ".pdf"
	strUrl = "http://webserv2/lss-lbis/print.asp?PDF=1&ID=" & Request("HID")
	'strUrl = "http://web03.zubuk.com/lss-lbis/print.asp?PDF=1&ID=" & Request("HID")
	'strURL = "http://oscar.ph.zubuk.com/lss-lbis/print.asp?PDF=1&ID=" & Request("HID")
	thedoc.HtmlOptions.PageCacheClear
	theDoc.HtmlOptions.RetryCount = 3
	theDoc.HtmlOptions.Timeout = 60000
	theDoc.Pos.X = 10
	theDoc.Pos.Y = 10
	theDoc.AddImageUrl strUrl
	theDoc.Save attachPDF
	strBody = "<font size='2' face='trebuchet MS'>A request has been assigned to you.<br><br>Attached is the verification form for the request. Please fill-up the form  upon completion." & vbCrlf & _
		"If there are any questions or clarifications, please contact the LanguageBank office immediately at 410-6183 or email us at " & vbCrLf & _
		"<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br><br>" & vbCrLf & _
		"The Language Bank</font><br><br>" & vbCrLf & _
		"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>"
	mlMail.Body = strBody
	mlMail.AttachFile (attachPDF)
End If

mlMail.Send
set mlMail=nothing

'CREATE LOG
Set fso = CreateObject("Scripting.FileSystemObject")
Set LogMe = fso.OpenTextFile(EmailLog, 8, True)
strLog = Now & vbCrLf & "----- EMAIL SENT -----" & vbCrLf & _
	"TO: " & Request("emailadd") & vbCrLf & _
	"REQUEST ID: " & Request("HID")
LogMe.WriteLine strLog
Set LogMe = Nothing
Set fso = Nothing

Session("MSG") = "E-Mail was sent to " &  Request("emailadd") & "."
Response.Redirect "reqconfirm.asp?ID=" & Request("HID")
%>