<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
dim strCon()
Function SearchArraysCon(xcon, strCon)
	DIM	lngMax, lngI
	SearchArraysCon = -1
	On Error Resume Next	
	lngMax = UBound(strCon)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If strCon(lngI) = Ucase(trim(xcon)) Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysCon = lngI
End Function
Function CleanFax(strFax)
	myfax = trim(strFax)
	myfax = Replace(myfax, "-", "")
	myfax = Replace(myfax, "(", "")
	myfax = Replace(myfax, ")", "") 
	CleanFax = myfax
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	server.scripttimeout = 360000
	Set rsEmail = Server.CreateObject("ADODB.RecordSet")
	sqlEmail = "SELECT Email, Fax, [index] from requester_T ORDER BY [index]"
	'sqlEmail = "SELECT DISTINCT(reqID), email, fax FROM request_T, requester_T WHERE appdate <= '" & Date & "' AND appdate >= '" & DateAdd("m", -13, Date) & "' AND (status = 1 Or status = 4) AND reqID = Requester_T.[index]"
	rsEmail.Open sqlEmail, g_strCONN, 1, 3
	x = 0
	Do Until rsEmail.EOF
		reqEmail = Z_FixNull(rsEmail("email"))
		reqFax = Z_FixNull(rsEmail("fax"))
		If reqEmail <> "" Or reqFax <> "" Then
			myEmailAdr = reqEmail
			If myEmailAdr = "" Then myEmailAdr = CleanFax(reqFax) & "@emailfaxservice.com" 
				
			xcon = myEmailAdr
			
			lngIdx = SearchArraysCon(xcon, strCon)
			If lngIdx < 0 Then 
				Redim Preserve strCon(x)
				strCon(x) = Ucase(trim(xcon))
				x = x + 1
			End If
			
		End If
		rsEmail.MoveNext
	Loop
	rsEmail.Close
	Set rsEmail = Nothing
	If Request("chkdraft") <> 1 Then 
		strMSG = Replace(Request("txtMSG"), vbCrLf, "<br>") 
	ElseIf Request("chkdraft") = 1 Then
		strMSG = "<table width='935px'>" & _
			"<tr>" & _
				"<td colspan='2' align='center'>" & _
					"<img border='0' src='https://languagebank.lssne.org/lsslbis/images/strip.jpg' />" & _
				"</td>" & _
			"</tr>" & _
			"<tr>" & _
				"<td align='left'>" & _
						"<img border='0' src='https://languagebank.lssne.org/lsslbis/images/smalllogo.jpg' />" & _
				"</td>" & _
				"<td align='right' style='font-family: times new roman, arial, trebuchet;'><nobr>July 2014</td>" & _
			"</tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr>" & _
				"<td align='left' colspan='2'>" & _
					"<p align='left' style='font-family: times new roman, arial, trebuchet;'>" & _
						"Dear Valued Partner,<br><br>" & _
						"This is an incredible time of transformation for our organization! We’re breaking the mold in human services" & _
						"with an innovative, holistic approach to client care that creates measurable, positive impact. We’re connecting" & _
						"with community partners throughout New England and building our services around those we serve.<br><br>" & _
						"As a vital component of our new strategy, we are changing our parent name to <b>Ascentria Care Alliance</b>" & _
						"effective September 1, 2014. Our new name and logo design represent ""rising together"" - putting our faith in" & _
						"action to help people of all backgrounds and beliefs achieve their full potential." & _
					"</p>" & _
					"<p align='center'>" & _
						"<img width='347px' height='94px' border='0' src='https://languagebank.lssne.org/lsslbis/images/ascentrialogo.jpg'>" & _
					"</p>" & _
					"<p align='left' style='font-family: times new roman, arial, trebuchet;'>" & _
						"The <b>Ascentria Care Alliance</b> name replaces Lutheran Social Services of New England, Inc. It represents our" & _
						"family of brands and partnerships that together comprise our client-centered care. Our business entity, Lutheran" & _
						"Community Services will change to Ascentria Community Services.<br><br>" & _
						"The names of our residential facilities and flagship services, such as the Lutheran Home of Southbury, Luther" & _
						"Ridge, Emmanuel House Residence, Emanuel Village, Good News Garage, LanguageBank and New Lands" & _
						"Farm, will not change. They will be known as members of Ascentria Care Alliance.<br><br>" & _
						"<b>Please note that this is a name change only. There is no change in ownership, our Federal Tax" & _
						"Identification numbers will not change, and our organization will continue to operate as a 501 (C) 3" & _
						"tax exempt nonprofit entity. Our mailing addresses, telephone and fax numbers will remain the same." & _
						"However on September 1, 2014 our email addresses will change to ""@ascentria.org"" and our Internet" & _
						"address will be www.ascentria.org.</b><br><br>" & _
						"A formal public announcement was made on June 12 and during the months ahead, you will notice updates to" & _
						"our website and business collateral that will reflect and emphasize our new name.<br><br>" & _
						"Should you have any questions about these changes, please visit us at lssne.org for a complete listing of" & _
						"frequently asked questions (FAQs). We look forward to continuing to work together with you in the future.<br><br>" & _
						"Sincerely," & _
					"</p>" & _
					"<p align='left'>" & _
						"<img border='0' src='https://languagebank.lssne.org/lsslbis/images/signangela.jpg'>" & _
					"</p>" & _
					"<p align='left' style='font-family: times new roman, arial, trebuchet;'>" & _
						"Angela Bovill<br>" & _
						"President and CEO" & _
					"</p>" & _
				"</td>" & _
			"</tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr><td>&nbsp;</td></tr>" & _
			"<tr>" & _
				"<td align='center' colspan='2' style='font-family: times new roman, arial, trebuchet; font-size: 10pt;'>" & _
					"14 East Worcester Street, Suite 300 • Worcester, MA 01604 • 774.243.3900<br>" & _
					"info@lssne.org • www.lssne.org" & _
				"</td>" & _
			"</tr>" & _
		"</table>"
		'strMSG = "<img border='0' src='https://languagebank.lssne.org/lsslbis/images/LBISLOGO.jpg'>" & _
		'	"<br><br>" & _
		'	"Dear Friends," & _
		'	"<br><br>" & _
		'	"Innovative organizations constantly think about ways to improve how they reach, connect with and impact the lives of their customers, staff, partners and supporters. " & _
		'	"<br><br>" & _
		'	"During the past year, Good News Garage and our parent organization Lutheran Social Services stepped back to reflect upon what we are called to do and how we do it. We spent time listening to what our stakeholders had to say about our future and we gained valuable insights that helped to shape our new mission, vision and strategic direction." & _
		'	"<br><br>" & _
		'	"This is an incredible time of transformation for our organization! We’re breaking the mold in human services with an innovative, holistic approach to client care that creates measurable, positive impact. We’re connecting with community partners throughout New England and building our services around those we serve. Our commitment to develop sustainable solutions that extend beyond today’s traditional support systems will transform how care is delivered." & _
		'"</p>" & _
		'"<p align='center'>" & _
		'	"<img width='347px' height='94px' border='0' src='https://languagebank.lssne.org/lsslbis/images/ascentrialogo.jpg'>" & _
		'"</p>" & _
		'"<p align='left'>" & _
		'	"As a vital component of our new strategy, Lutheran Social Services of New England, Inc. is changing its name to <b>Ascentria Care Alliance</b> effective September 1, 2014. Ascentria Care Alliance represents our family of brands and partners that together comprise our client-centered care." & _
		'	"<br><br>" & _
		'	"While our parent name is changing, the <b>LanguageBank</b> name and the quality services it provides will remain the same – as <b>LanguageBank, a member of Ascentria Care Alliance</b>. The Ascentria name will help us open doors for new partnerships and expand our funding opportunities." & _  
		'	"<br><br>" & _
		'	"Amidst these changes, our faith-based heritage, the cornerstone of LanguageBank and Lutheran Social Services, remains constant. Ascentria Care Alliance and our new logo design represent “rising together” – putting our faith in action to help people of all backgrounds and beliefs achieve their full potential." & _
		'	"<br><br>" & _
		'	"A formal public announcement will be made soon. During the months ahead, you will notice changes to our website and business collateral that will reflect and emphasize our new name. Should you have any questions about these changes, please visit us at lssne.org for a list of frequently asked questions (FAQs)." & _
		'	"<br><br>" & _
		'	"While we’re excited about our new name, it’s our mission and vision that make us enthusiastic about what we do every day: helping to empower others, rise together and realize new possibilities!" & _
		'	"<br><br>" & _
		'	"Sincerely," & _
		'	"<br>" & _
		'	"<img width='115px' height='53px' border='0' src='https://languagebank.lssne.org/lsslbis/images/sign.jpg'>" & _
		'	"<br>" & _
		'	"Angela Bovill" & _
		'	"<br>" & _
		'	"President and CEO"
	End If
	strSUBJ = Trim(Request("txtSub"))
	Set mlMail = CreateObject("CDO.Message")
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.socketlabs.com"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "server3874"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "UO2CUSxat9ZmzYD7jkTB"
	mlMail.Configuration.Fields.Update
	y = 0
	Do until y = ubound(strCon) + 1
	on error resume next
		'response.write strCon(y) & "<br>"
		mlMail.To = strCon(y)
		'mlMail.To = "patrick@zubuk.com;phutrek@yahoo.com"
		'mlMail.Cc = "language.services@thelanguagebank.org"
		mlMail.From = "language.services@thelanguagebank.org"
		mlMail.Subject = strSUBJ
		If Request("chkdraft") <> 1 Then 
			mlMail.HTMLBody = "<html><body><p>" & vbCrLf & strMSG & vbCrLf & "</p></body></html>"
		ElseIf Request("chkdraft") = 1 Then
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strMSG & vbCrLf & "</body></html>"
		End If
		mlMail.Send
		'response.write myEmailAdr & "<br>"
		'response.write "email sent to " & strCon(y) & "<br>"
		y = y + 1
	Loop
	Set mlMail = Nothing
	Session("MSG") = "Email Sent. COUNT: " &  ubound(strCon)

	'response.redirect "main.asp"
End If
%>
<html>
	<head>
		<title>LanguageBank - Mass Email</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SendMe()
		{
			if (document.frmMemail.txtSub.value == "")
			{
				alert("Please include a subject.")
				return;
			}
			if (document.frmMemail.txtMSG.value == "" && document.frmMemail.chkdraft.checked == false)
			{
				alert("Please include a message.")
				return;
			}
			var ans = window.confirm("Send eMail? This might take a few minutes to complete.");
			if (ans){
			document.frmMemail.action = "massmail.asp";
			document.frmMemail.submit();
			}
		}
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
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmMemail'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='center' colspan='10'>
									<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
										<table border='0' cellspacing='1'>		
											<tr>
												<td><span class='error'><%=Session("MSG")%></span></td>
											</tr>
										</table>
									</div>
								</td>
							</tr>
							<tr>
								<td align='center'>
									<table border='0' cellspacing='1'>
											<tr>
											<td valign='top'>Subject:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='txtSub' onkeyup='bawal(this);'>
												<input type="checkbox" name="chkdraft" value="1">Send LB client letter
											</td>
										</tr>
										<tr>
											<td valign='top'>Message:</td>
											<td>
												<textarea class='main' style='width: 400px;' name='txtMSG' cols='75' rows='10' onkeyup='bawal(this);'>
												</textarea>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan='2' align='right'>
												<input class='btn' type='button' value='Send' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SendMe();'>
											</td>
										</tr>
									</table>
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
		</form>
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
