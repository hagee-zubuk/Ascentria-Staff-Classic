<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<%
Session.Abandon
%>
<html>
	<head>
		<title>Welcome to Language Bank - Staff Site - WS6</title>
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
										<tr>
											<td class="hdr" width="35%">
												Staff Portal Log-In
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td>
												<h1>This site needs Javascript to work properly. Please enable Javascript and/or update your browser to it's newest version.</h1>
												<h4>If it is enabled already, click <a href="default.asp">here</a> to return to the sign in page.</h4>
											</td>
										</tr>
										<tr>
											<td align='center' colspan='2'>
												<span class='error'><%=Session("MSG")%></span>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
										<tr>
											<td class="nrml" colspan="2">* Your browser should support cookies and allow pop-ups</td>
										</tr>
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