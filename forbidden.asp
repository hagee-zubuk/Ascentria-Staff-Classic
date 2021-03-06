<%Language=VBScript%>
<%
Session.Abandon
myip = Request.ServerVariables("REMOTE_ADDR")
myproxy = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
%>
<html>
	<head>
		<title>Welcome to Language Bank - Vendor Site - Forbidden</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
	</head>
	<body onload='document.frmLogIn.txtUN.focus();'>
		<form method='post' name='frmLogIn' action='signin.asp'>
			<table cellSpacing='5' cellPadding='0' width="95%" border='0' align="center">
				<tr>
					<td valign='top' align="left" rowspan="2" width="80%" height="65px">
						<img src='images/LBISLOGO.jpg' border="0">
					</td>
					<td align="center" width="25%" class="tollnum">
					Toll-Free 844.579.0610
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>	
				<tr>
					<td colspan="2" class="motto" align="center">
						Understand and Be Understood.
					</td>
				</tr>
				<tr>
					<td colspan="2" width="100%">
						<table cellSpacing='5' cellPadding='0' border='0' width="100%" align="center">
							<tr>
								<td class="defborder" width="25%">&nbsp;</td>
								<td width="85%">
									<table class="defborder" width="100%" border='0'>
										<tr><td>&nbsp;</td></tr>
										<tr><td>&nbsp;</td></tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td align='center'>
												<h1>IP: <%=myip%></h1>
												<h4><%=myproxy%></h4>
												<br>
												<h3>
													Sorry. Your computer's IP address is currently not listed on our records. Please send the IP listed above or a screenshot of this page to <a href="mailto:language.services@thelanguagebank.org?Subject=IP%20Rejected%20<%=myip%>" target="_top">language.services@thelanguagebank.org</a> to verify. Please include 
													the institution you work for. We will get back to you as soon we have verified your IP address. Thank you for you using Languagebank for you interpretation needs.
												</h3>
											</td>
										</tr>
										<tr>
											<td align='center' colspan='2'>
												<span class='error'><%=Session("MSG")%></span>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr><td>&nbsp;</td></tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td colspan="2" width="100%">
												<table cellSpacing='5' cellPadding='0' border='0' width="60%" align="center">
													<tr>
														<td class="defborder2" align="center">
															<p class="hdr">Did You Know?</p>
															<p class="nrml" align="left">
																Did you know that Language Bank
																also provides written translation services?
																Transalate your written English forms,
																signage and agreements into
																languages your customers understand.<br /><br />
															</p>
														</td>
														<td class="defborder2" align="center">
															<p class="hdr">Services Available 24 x 7</p>
															<p class="nrml" align="left">
																Need Language Bank services after
																hours or during the weekend? <b>Call
																us toll-free 844.579.0610 ANYTIME</b>
																and we will gladly assist you.<br /><br />
															</p>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<table class="defborder" border='0' align="center"  width="100%">
							<tr>
								<td width="76%">&nbsp;</td>
								<td width="24%" class="footnew">
									Office Locations:<br />
									11 Shattuck Street, Worcester MA 01605<br />
									340 Granite Street, 3rd Floor, Manchester, NH 03102
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
<%
Session("MSG") = ""
%>